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
    print("✅ Base de datos de RNAS cargada.")
except Exception as e:
    print(f"❌ Error con el Excel de referencia: {e}")

def extraer_datos_afip(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text()
        
        # Extracción de CUITs
        cuits = re.findall(r"CUIT:\s*(\d{11})", texto)
        cuit_prestador = cuits[0] if len(cuits) > 0 else ""
        cuit_rnas_receptor = cuits[1] if len(cuits) > 1 else ""
        
        # Nombre del emisor (limpio)
        emisor_match = re.search(r"Razón Social:\s*(.*?)(?=Fecha de Emisión|$)", texto, re.DOTALL)
        emisor_nombre = emisor_match.group(1).strip() if emisor_match else "Desconocido"

        # --- MEJORA: TIPO DE COMPROBANTE CON ESPACIO ---
        # Busca "Factura", "Nota de Crédito" o "Nota de Débito" + la letra
        tipo_raw = re.search(r"(Factura|Nota\s+de\s+Crédito|Nota\s+de\s+Débito)\s*([ABC])", texto, re.IGNORECASE)
        if tipo_raw:
            # Reconstruimos con un espacio fijo: "Factura C"
            tipo_comp = f"{tipo_raw.group(1).capitalize()} {tipo_raw.group(2).upper()}"
        else:
            tipo_comp = "Factura C"

        # Punto de Venta y Número de Comprobante
        punto_vta_match = re.search(r"Punto de Venta:\s*(\d+)", texto)
        nro_comp_match = re.search(r"Comp\. Nro:\s*(\d+)", texto)
        pv = str(punto_vta_match.group(1)).zfill(5) if punto_vta_match else ""
        nc = str(nro_comp_match.group(1)).zfill(8) if nro_comp_match else ""

        # --- MEJORA: EXTRACCIÓN DE CAE ---
        # Buscamos 14 dígitos seguidos que estén cerca de la palabra CAE
        cae_match = re.search(r"CAE\s*N°?[:\s]*(\d{14})", texto)
        cae = cae_match.group(1) if cae_match else ""

        # Fecha y Monto
        fecha = re.search(r"Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})", texto)
        total_raw = re.search(r"Importe Total:\s*\$\s*([\d\.,]+)", texto)
        total_num = 0.0
        if total_raw:
            val = total_raw.group(1).replace(".", "").replace(",", ".")
            total_num = float(val)

        return {
            "PRESTADOR_NOMBRE": emisor_nombre,
            "N.º DE COMPROBANTE": nc,
            "TIPO DE COMPROBANTE": tipo_comp,
            "PUNTO DE VENTA": pv,
            "N.º CAE": cae,
            "FECHA DE COMPROBANTE": fecha.group(1) if fecha else "",
            "MONTO $": total_num,
            "CUIT PRESTADOR": cuit_prestador,
            "CUIT RNAS": cuit_rnas_receptor 
        }

# 2. Procesar archivos
lista_total = []
if os.path.exists(PATH_ENTRADA):
    archivos = [f for f in os.listdir(PATH_ENTRADA) if f.lower().endswith(".pdf")]
    for archivo in archivos:
        try:
            datos = extraer_datos_afip(os.path.join(PATH_ENTRADA, archivo))
            lista_total.append(datos)
        except Exception as e:
            print(f"Error en {archivo}: {e}")

if lista_total:
    df_facturas = pd.DataFrame(lista_total)

    # 3. CRUCE CON RNAS
    df_final = pd.merge(
        df_facturas,
        df_consulta[['CUIT', 'RNAS']], 
        left_on='CUIT RNAS',
        right_on='CUIT',
        how='left'
    )
    df_final = df_final.rename(columns={'RNAS': 'N.º RNOS'})

    # 4. GUARDAR POR PRESTADOR
    for cuit, grupo in df_final.groupby('CUIT PRESTADOR'):
        nombre_p = grupo['PRESTADOR_NOMBRE'].iloc[0]
        nombre_archivo = f"Reporte_{re.sub(r'[\\/*?:<>|]', '', str(nombre_p))}.xlsx"
        
        columnas_finales = [
            "N.º DE COMPROBANTE",
            "TIPO DE COMPROBANTE",
            "PUNTO DE VENTA",
            "N.º CAE",
            "FECHA DE COMPROBANTE",
            "MONTO $",
            "CUIT PRESTADOR",
            "CUIT RNAS",
            "N.º RNOS"
        ]
        
        ruta_archivo = os.path.join(PATH_SALIDA, nombre_archivo)
        # Forzamos que las columnas de números con ceros se guarden como texto
        grupo[columnas_finales].to_excel(ruta_archivo, index=False)
        print(f"✅ Reporte generado: {nombre_archivo}")