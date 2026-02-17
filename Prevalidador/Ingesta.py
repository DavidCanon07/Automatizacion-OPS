from pathlib import Path
import pandas as pd
from Configuracion_parametros import carpeta, exts, clave, Campos_a_validar, largo_campos

#libreria para funciones de formatea para nombres de columnas extra
import re
from openpyxl.utils import get_column_letter

#Función para validar que el dataframe tenga las columnas esperadas y detectar columnas extra
def validar_columnas(df, archivo):
    
    faltantes = set(Campos_a_validar) - set(df.columns)
    extras = set(df.columns) - set(Campos_a_validar)
    columnas_extra_legibles = [formatea_nombre_columna(col) for col in extras]
    #Si hay columnas faltantes, se lanza una excepción crítica. Si hay columnas extra, se lanza una advertencia.
    if faltantes:
        raise Exception(
            f"\n❌ ERROR CRÍTICO\n"
            f"Archivo: {archivo}\n"
            f"Columnas faltantes: {faltantes}"
        )
        

    if extras:
        raise Exception(
            f"\nO_O Advertencia en {archivo}\n"
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
                df = pd.read_excel(i, skiprows=6, dtype=str)
                validar_columnas(df, i.name) #validar las columnas del archivo
                df["__archivo_origen"] = i.name
                datos.append(df)

        if not datos:
            raise Exception("No hay archivos válidos para procesar")
        return datos
    except Exception as e:

        print(f"\n-X- ERROR AL OBTENER LOS DATOS\n{str(e)}")
        print(e)
        raise

#Concatena la información del dataframe
def concatenar_datos():
    datos = obtener_datos(carpeta, clave, exts)
    df_total = pd.concat(datos, ignore_index=True)
    return df_total

#Función para validar el largo de campos específicos
def validar_largo_campos(df):
    
    errores_totales = []
    # Validar cada campo según su regla de largo
    for campo, regla in largo_campos.items():
        if campo not in df.columns:
            continue
        
        # Limpieza segura
        df[campo] = (
            df[campo]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
            .str.strip()
        )

        # Ignorar filas vacías
        df_validar = df[df[campo] != ""]

        # Validación
        if isinstance(regla, int):
            errores = df_validar[df_validar[campo].str.len() != regla]
        elif isinstance(regla, tuple):
            min_len, max_len = regla
            errores = df_validar[
                (df_validar[campo].str.len() < min_len) |
                (df_validar[campo].str.len() > max_len)
            ]
        else:
            continue
        if not errores.empty:
            print(f"\n❌ ERROR EN LARGO '{campo}'")
            print(errores[[campo, "__archivo_origen"]])
            errores["campo_error"] = campo
            errores_totales.append(errores)
            
    # Consolidar errores
    if errores_totales:
        errores_totales = pd.concat(errores_totales)
        ruta_error = "C:\\Users\\dacanonm\\OneDrive - Indra\\Documentos\\GitHub\\automatizacion OPS\\Prevalidador\\errores_validacion_largo_campos.xlsx"
        errores_totales.to_excel(ruta_error, index=True)
        raise Exception(
            f"Se encontraron errores de longitud en campos:\n"
            f"{errores_totales[['campo_error', '__archivo_origen']]}"
        )

    return df






#Función para convertir nombres de columnas extra a un formato más legible
def formatea_nombre_columna(nombre: str) -> str:
    
    #Si el nombre es 'Unnamed: N', devuelve 'Columna {LETRA}' (A=1).
    #Si no, devuelve el nombre original.
    #Ejemplo: 'Unnamed: 0' -> 'Columna A', 'Unnamed: 1' -> 'Columna B', etc.
    m = re.match(r"^Unnamed:\s*(\d+)$", str(nombre), flags=re.IGNORECASE)
    if m:
        idx0 = int(m.group(1))           # índice 0-based
        letra = get_column_letter(idx0 + 1)  # Excel es 1-based
        return f"Columna {letra}"
    return str(nombre)






