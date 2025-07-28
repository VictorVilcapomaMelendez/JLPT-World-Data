# =============================================================================
# JLPT Data Processor Script
# =============================================================================
# Librerias necesarias:
# pip install pandas openpyxl python-dotenv pywin32 pycountry pycountry-convert

import os
import shutil
import unicodedata
import pandas as pd
import pythoncom
from pathlib import Path
from win32com.client import Dispatch
from dotenv import load_dotenv
import pycountry
import pycountry_convert as pc

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE RUTAS (Solo se usan 2 carpetas principales)
# -----------------------------------------------------------------------------
# Se debe crear un archivo .env en el mismo ambiente de python de la siguiente forma:
# Carpeta_Excel_Raw="TuRutaATuCarpetaConLosExcels del estilo C:\..."
# Carpeta_Consolidada="CualquierOtraRutaParaAlmacenarLoConsolidado"
load_dotenv()

# Carpeta donde están los archivos .xls y .xlsx originales
carpeta_data = Path(os.getenv("Carpeta_Excel_Raw", "")).resolve()

# Carpeta donde se guardan los archivos .xlsx convertidos y normalizados
carpeta_output = Path(os.getenv("Carpeta_Consolidada", "")).resolve()

# Subcarpetas internas para organización temporal
carpeta_temp = carpeta_output / "temp"
carpeta_temp.mkdir(parents=True, exist_ok=True)

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------

def contains_japanese(text):
    for ch in str(text):
        if ('\u4e00' <= ch <= '\u9fff') or ('\u3040' <= ch <= '\u309f') or ('\u30a0' <= ch <= '\u30ff'):
            return True
    return False

def obtener_continente(pais):
    try:
        country_code = pc.country_name_to_country_alpha2(pais, cn_name_format="default")
        continent_code = pc.country_alpha2_to_continent_code(country_code)
        continents = {
            "AF": "Africa", "NA": "North America", "SA": "South America",
            "AS": "Asia",   "EU": "Europe",        "OC": "Oceania"
        }
        return continents.get(continent_code, "Unknown")
    except Exception:
        return "Unknown"

def obtener_codigo_iso(pais):
    try:
        return pycountry.countries.lookup(pais).alpha_2
    except LookupError:
        return None

# -----------------------------------------------------------------------------
# PARTE 1: Convertir .xls a .xlsx y copiar los .xlsx
# -----------------------------------------------------------------------------

def convertir_y_copiar(origen: Path, destino: Path):
    pythoncom.CoInitialize()
    excel = Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    for archivo in origen.iterdir():
        if archivo.suffix.lower() == ".xls":
            ruta = str(archivo)
            nuevo = destino / (archivo.stem + ".xlsx")
            try:
                wb = excel.Workbooks.Open(ruta, ReadOnly=True)
                wb.SaveAs(str(nuevo), FileFormat=51)
                wb.Close(SaveChanges=False)
                print(f"Convertido: {archivo.name} → {nuevo.name}")
            except Exception as e:
                print(f"Error al convertir {archivo.name}: {e}")
        elif archivo.suffix.lower() == ".xlsx":
            try:
                shutil.copy2(str(archivo), str(destino / archivo.name))
                print(f"Copiado: {archivo.name}")
            except Exception as e:
                print(f"Error al copiar {archivo.name}: {e}")

    excel.Quit()
    pythoncom.CoUninitialize()

convertir_y_copiar(carpeta_data, carpeta_temp)

# -----------------------------------------------------------------------------
# PARTE 2: Normalización de los archivos
# -----------------------------------------------------------------------------

carpeta_normalizado = carpeta_output / "normalizado"
carpeta_normalizado.mkdir(parents=True, exist_ok=True)

for nombre_archivo in os.listdir(carpeta_temp):
    if not nombre_archivo.endswith(".xlsx"):
        continue
    try:
        print(f"Procesando: {nombre_archivo}")
        ruta_entrada = carpeta_temp / nombre_archivo
        nombre_base = os.path.splitext(nombre_archivo)[0].replace('_Normalizado', '')
        ruta_salida = carpeta_normalizado / f"{nombre_base}_Normalizado.xlsx"

        df = pd.read_excel(ruta_entrada, header=None)
        df = df.drop(columns=0).drop(index=0).reset_index(drop=True)
        df = df.drop(columns=[2, 3], errors='ignore')

        encabezado0 = df.iloc[0].astype(str).apply(lambda x: unicodedata.normalize("NFKC", x))
        if "合計" in list(encabezado0):
            idx_total = list(encabezado0).index("合計")
            df = df.iloc[:, :idx_total]

        if df.shape[1] > 14:
            df = df.iloc[:, :14]

        df = df.replace({'応募者': 'Applicant', '受験者': 'Examinee'})
        encabezados = [unicodedata.normalize("NFKC", str(e)) for e in df.iloc[0].tolist()]

        niveles = ['N1', 'N2', 'N3', 'N4', 'N5']
        nuevas_columnas = ['Country/Region', 'City (ENG)']
        i = 2
        for nivel in niveles:
            if i + 1 < len(encabezados):
                nuevas_columnas += [f"{nivel} Applicants", f"{nivel} Examinees"]
                i += 2

        df.columns = nuevas_columnas
        df = df.drop(index=0).reset_index(drop=True)

        for idx, val in df['Country/Region'].items():
            if contains_japanese(val) and (idx + 1) in df.index:
                df.at[idx, 'Country/Region'] = unicodedata.normalize('NFKC', str(df.at[idx + 1, 'Country/Region']))
        df['Country/Region'] = (
            df['Country/Region']
            .astype(str)
            .apply(lambda x: unicodedata.normalize('NFKC', x))
            .apply(lambda x: None if contains_japanese(x) else x)
        )
        df['Country/Region'] = df['Country/Region'].replace('nan', pd.NA).fillna(method='ffill')

        df = df[df['Country/Region'].notna()]
        df = df[df['City (ENG)'].notna() & (df['City (ENG)'].astype(str).str.strip() != '')]

        df_long = df.melt(
            id_vars=['Country/Region', 'City (ENG)'],
            var_name='metric',
            value_name='Count'
        )
        df_long[['Level', 'Type']] = df_long['metric'].str.split(' ', expand=True)
        df_long = df_long.drop(columns=['metric'])

        parts = nombre_base.split('_')
        year = int(parts[0])
        mid = int(parts[1])
        fecha = pd.Timestamp(year, mid, 1).date()
        df_long['Fecha'] = fecha
        df_long['Año'] = year
        df_long['Mes'] = 'July' if mid == 1 else ('December' if mid == 2 else '')

        df_long['Country/Region'] = df_long['Country/Region'].astype(str).str.strip()
        reemplazos_paises = {
            "Brunei": "Brunei Darussalam", "Russia": "Russian Federation", "Turkey": "Türkiye",
            "Korea": "South Korea", "Ivory Coast": "Côte d'Ivoire", "Cote d'Ivoire": "Côte d'Ivoire",
            "Cote d' Ivoire": "Côte d'Ivoire", "DR Congo": "Congo, The Democratic Republic of the",
            "Democratic Republic of the Congo": "Congo, The Democratic Republic of the",
            "Mongol": "Mongolia", "U.S.A.": "United States", "U.K.": "United Kingdom",
            "Czech": "Czech Republic", "Catarrh": "Qatar"
        }
        df_long['Country/Region'] = df_long['Country/Region'].replace(reemplazos_paises)
        df_long['City & Country/ Region'] = df_long['City (ENG)'] + ', ' + df_long['Country/Region']
        df_long['Continent'] = df_long['Country/Region'].apply(obtener_continente)
        df_long['Country Code'] = df_long['Country/Region'].apply(obtener_codigo_iso)
        df_long['Flag URL'] = df_long['Country Code'].apply(
            lambda code: f"https://flagcdn.com/w40/{code.lower()}.png" if pd.notna(code) else None
        )

        df_long.to_excel(ruta_salida, index=False)
        print(f"✔ Guardado: {ruta_salida}")

    except Exception as e:
        print(f"❌ Error procesando {nombre_archivo}: {e}")

# -----------------------------------------------------------------------------
# PARTE 3: Consolidación final
# -----------------------------------------------------------------------------

archivos_normalizados = list(carpeta_normalizado.glob("*.xlsx"))
if not archivos_normalizados:
    raise RuntimeError("No se encontró ningún archivo válido para consolidar.")

dataframes = []
for archivo in archivos_normalizados:
    if archivo.name.startswith("~$"):
        continue
    try:
        df = pd.read_excel(archivo)
        dataframes.append(df)
    except Exception as e:
        print(f"✘ No se pudo leer '{archivo}': {e}")

df_consolidado = pd.concat(dataframes, ignore_index=True)
if 'Fecha' in df_consolidado.columns:
    df_consolidado['Fecha'] = pd.to_datetime(df_consolidado['Fecha']).dt.date

ruta_final = carpeta_output / "JLPT_Historico.xlsx"
with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
    df_consolidado.to_excel(writer, index=False, sheet_name="JLPT_Historico")
    worksheet = writer.sheets["JLPT_Historico"]
    for col_cells in worksheet.columns:
        max_length = max(len(str(cell.value)) for cell in col_cells)
        col_letter = col_cells[0].column_letter
        worksheet.column_dimensions[col_letter].width = max_length + 2

print(f"✔ Archivo consolidado guardado en: {ruta_final}")
