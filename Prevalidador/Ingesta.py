from pathlib import Path
import pandas as pd
from Configuracion_parametros import carpeta, exts, clave, Campos_a_validar

#libreria para funciones de formatea para nombres de columnas extra
import re
from openpyxl.utils import get_column_letter


def validar_columnas(df, archivo):
    
    faltantes = set(Campos_a_validar) - set(df.columns)
    extras = set(df.columns) - set(Campos_a_validar)
    columnas_extra_legibles = [formatea_nombre_columna(col) for col in extras]
    if faltantes:
        raise Exception(
            f"\n❌ ERROR CRÍTICO\n"
            f"Archivo: {archivo}\n"
            f"Columnas faltantes: {faltantes}"
        )
        

    if extras:
        raise Exception(
            f"\n⚠️ Advertencia en {archivo}\n"
            f"Columnas extra: {', '.join(sorted(columnas_extra_legibles))}"
        )
        


#Obtener los archivos de OPS + columna de nombre del archivo
def obtener_datos(carpeta, clave, exts):
    datos = []
    try:    
        for i in Path(carpeta).iterdir():
            if i.is_file() and i.suffix.lower() in exts and str(clave).lower() in i.stem.lower():
            # lee cada archivo de OPS que se encuentre en la ruta
                print(f"Archivo encontrado: {i.name}")
                df = pd.read_excel(i, skiprows=7, dtype=str)
                validar_columnas(df, i.name) #validar las columnas del archivo
                df["__archivo_origen"] = i.name
                datos.append(df)
        if not datos:
            raise Exception("No hay archivos válidos para procesar")
        return datos
    except Exception as e:

        print(f"\n❌ ERROR CRÍTICO")
        print(e)
        raise

#Concatena la información del dataframe
def concatenar_datos():
    datos = obtener_datos(carpeta, clave, exts)
    df_total = pd.concat(datos, ignore_index=True)
    return df_total





#Función para convertir nombres de columnas extra a un formato más legible
def formatea_nombre_columna(nombre: str) -> str:
    
    #Si el nombre es 'Unnamed: N', devuelve 'Columna {LETRA}' (A=1).
    #Si no, devuelve el nombre original.
    
    m = re.match(r"^Unnamed:\s*(\d+)$", str(nombre), flags=re.IGNORECASE)
    if m:
        idx0 = int(m.group(1))           # índice 0-based
        letra = get_column_letter(idx0 + 1)  # Excel es 1-based
        return f"Columna {letra}"
    return str(nombre)






