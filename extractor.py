import pdfplumber
import pandas as pd
import re
import os

# CONFIGURACIÓN
PATH_ENTRADA = "./facturas"
PATH_SALIDA = "./reportes"
ARCHIVO_REFERENCIA = "rnas_involucradas_ley_23793.xlsx"

os.makedirs(PATH_SALIDA, exist_ok=True)

# 1. Cargar base de datos de consulta (RNAS)
try:
    df_consulta = pd.read_excel(ARCHIVO_REFERENCIA, dtype={'CUIT': str})
    df_consulta['RNAS'] = df_consulta['RNAS'].astype(str)
    print("✅ Base de datos de RNAS cargada.")
except Exception as e:
    print(f"❌ Error con el Excel de referencia: {e}")

def extraer_datos_afip(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text()
        
        cuits = re.findall(r"CUIT:\s*(\d{11})", texto)
        cuit_prestador = cuits[0] if len(cuits) > 0 else ""
        cuit_rnas_receptor = cuits[1] if len(cuits) > 1 else ""
        
        receptor_match = re.search(r"Apellido y Nombre / Razón Social:\s*(.*?)(?=Condición frente al IVA|$)", texto, re.DOTALL)
        receptor_nombre = receptor_match.group(1).strip() if receptor_match else "Desconocido"

        emisor_match = re.search(r"Razón Social:\s*(.*?)(?=Fecha de Emisión|$)", texto, re.DOTALL)
        emisor_nombre = emisor_match.group(1).strip() if emisor_match else "Desconocido"

        tipo_raw = re.search(r"(Factura|Nota\s+de\s+Crédito|Nota\s+de\s+Débito)\s*([ABC])", texto, re.IGNORECASE)
        tipo_comp = f"{tipo_raw.group(1).capitalize()} {tipo_raw.group(2).upper()}" if tipo_raw else "Factura C"

        pv_match = re.search(r"Punto de Venta:\s*(\d+)", texto)
        nc_match = re.search(r"Comp\. Nro:\s*(\d+)", texto)
        
        # Convertimos a int para que Excel los tome como número
        pv = int(pv_match.group(1)) if pv_match else 0
        nc = int(nc_match.group(1)) if nc_match else 0

        cae_match = re.search(r"CAE\s*N°?[:\s]*(\d{14})", texto)
        fecha_match = re.search(r"Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})", texto)
        
        total_raw = re.search(r"Importe Total:\s*\$\s*([\d\.,]+)", texto)
        total_num = 0.0
        if total_raw:
            val = total_raw.group(1).replace(".", "").replace(",", ".")
            total_num = float(val)

        return {
            "PRESTADOR_NOMBRE": emisor_nombre,
            "RECEPTOR_RAZON_SOCIAL": receptor_nombre,
            "N.º DE COMPROBANTE": nc,
            "TIPO DE COMPROBANTE": tipo_comp,
            "PUNTO DE VENTA": pv,
            "N.º CAE": cae_match.group(1) if cae_match else "",
            "FECHA DE COMPROBANTE": fecha_match.group(1) if fecha_match else "",
            "MONTO $": total_num,
            "CUIT PRESTADOR": cuit_prestador,
            "CUIT RNAS": cuit_rnas_receptor 
        }

# 2. Procesar archivos
lista_total = []
archivos = [f for f in os.listdir(PATH_ENTRADA) if f.lower().endswith(".pdf")]
for archivo in archivos:
    try:
        lista_total.append(extraer_datos_afip(os.path.join(PATH_ENTRADA, archivo)))
    except Exception as e: print(f"Error en {archivo}: {e}")

if lista_total:
    df_facturas = pd.DataFrame(lista_total)
    df_facturas['FECHA_DT'] = pd.to_datetime(df_facturas['FECHA DE COMPROBANTE'], format='%d/%m/%Y')

    df_final = pd.merge(df_facturas, df_consulta[['CUIT', 'RNAS']], left_on='CUIT RNAS', right_on='CUIT', how='left')
    df_final = df_final.rename(columns={'RNAS': 'N.º RNOS'})

    # 4. FUNCIÓN PARA GUARDAR CON FORMATO REQUERIDO
    def guardar_excel_formateado(df_sub, nombre_archivo, incluir_razon_social=False):
        cols = ["N.º DE COMPROBANTE", "TIPO DE COMPROBANTE", "PUNTO DE VENTA", "N.º CAE", 
                "FECHA DE COMPROBANTE", "MONTO $", "CUIT PRESTADOR", "CUIT RNAS", "N.º RNOS"]
        
        if incluir_razon_social:
            cols.insert(1, "RECEPTOR_RAZON_SOCIAL")

        df_sub = df_sub.sort_values(by=['N.º RNOS', 'FECHA_DT'], ascending=[True, True])
        
        writer = pd.ExcelWriter(os.path.join(PATH_SALIDA, nombre_archivo), engine='xlsxwriter')
        df_sub.to_excel(writer, index=False, columns=cols, sheet_name='Detalle')
        
        workbook  = writer.book
        worksheet = writer.sheets['Detalle']

        # FORMATOS
        # Encabezado: Negrita, fondo blanco (sin bg_color), centrado y con bordes
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        # Formato Número Estándar
        num_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'num_format': '0'})
        # Formato Moneda
        money_fmt = workbook.add_format({'num_format': '"$" #,##0.00', 'align': 'center'})
        # Formato Texto Centrado
        text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

        for col_num, value in enumerate(cols):
            # Escribir el encabezado
            worksheet.write(0, col_num, value, header_fmt)
            
            # Aplicar formatos según tipo de dato
            if value == "MONTO $":
                worksheet.set_column(col_num, col_num, 15, money_fmt)
            elif value in ["N.º DE COMPROBANTE", "PUNTO DE VENTA"]:
                worksheet.set_column(col_num, col_num, 18, num_fmt)
            else:
                worksheet.set_column(col_num, col_num, 20, text_fmt)

        writer.close()

    # 5. SEPARACIÓN Y GUARDADO
    for cuit_p, grupo in df_final.groupby('CUIT PRESTADOR'):
        nombre_p = grupo['PRESTADOR_NOMBRE'].iloc[0]
        rnas_si = grupo[grupo['N.º RNOS'].notna()].copy()
        rnas_no = grupo[grupo['N.º RNOS'].isna()].copy()
        
        # Limpiar nombre para el archivo
        nombre_limpio = re.sub(r'[^a-zA-Z0-9 ]', '', str(nombre_p))
        
        if not rnas_si.empty:
            guardar_excel_formateado(rnas_si, f"Reporte_{nombre_limpio}.xlsx")
        
        if not rnas_no.empty:
            guardar_excel_formateado(rnas_no, f"Otras_Instituciones_{nombre_limpio}.xlsx", incluir_razon_social=True)
    
    print(f"✅ Procesamiento finalizado. Archivos generados en {PATH_SALIDA}")