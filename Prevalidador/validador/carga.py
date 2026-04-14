from pathlib import Path
import pandas as pd
import time as t
from Configuracion_parametros import carpeta_archivos, escribir, exts, clave, columna_ancla, omitir
from validador.utils import formatea_nombre_columna
from Configuracion_parametros import Campos_a_validar

#-------------------------------------------------------------------------------------------------------------
#Función para validar que el dataframe tenga las columnas esperadas y detectar columnas extra
def validar_columnas(df, archivo):
    
    # Para evitar que columnas extra al final del dataframe (después de la columna "DIGITO DE VERIFICACION") sean consideradas como faltantes, se limita la validación solo hasta esa columna. Si la columna límite no existe, se valida contra todas las columnas del dataframe.
    limite = "DIGITO DE VERIFICACION"
    # Si la columna límite existe, se consideran solo las columnas hasta esa columna (inclusive). Si no existe, se consideran todas las columnas del dataframe.
    if limite in df.columns:
        idx_limite = df.columns.get_loc(limite)
        columnas_hasta_limite = list(df.columns[:idx_limite + 1])
    else:
        columnas_hasta_limite = list(df.columns)
    
    # Se comparan las columnas hasta el límite con las columnas esperadas. Las columnas extra se formatean a un nombre más legible para el reporte.
    faltantes = set(Campos_a_validar) - set(columnas_hasta_limite)
    extras = set(columnas_hasta_limite) - set(Campos_a_validar)
    columnas_extra_legibles = [formatea_nombre_columna(col) for col in extras]
    #Si hay columnas faltantes, se lanza una excepción crítica. Si hay columnas extra, se lanza una advertencia.
    if faltantes:
        raise Exception(
            f"\n✕ ERROR CRÍTICO\n"
            f"Archivo: {archivo}\n"
            f"Columnas faltantes: {faltantes}"
        )

    if extras:
        raise Exception(
            f"\n◬ Advertencia en {archivo}\n"
            f"Columnas extra: {', '.join(sorted(columnas_extra_legibles))}"
        )

#-------------------------------------------------------------------------------------------------------------
#Función para eliminar el footer dinámicamente evaluando múltiples columnas ancla. Una fila se considera 'válida' si tiene dato en AL MENOS 3 de las columnas ancla. Todo lo que venga después de la última fila válida se descarta.
def recortar_footer_dinamico(df):
    """
    Elimina el footer dinámicamente evaluando múltiples columnas ancla.
    Una fila se considera 'válida' si tiene dato en AL MENOS 3 de las columnas ancla.
    Todo lo que venga después de la última fila válida se descarta.
    """

    # Filtrar solo las columnas ancla que existen en el df
    cols_presentes = [c for c in columna_ancla if c in df.columns]
    if not cols_presentes:
        return df

    def tiene_dato(val):
        s = str(val).strip()
        return s != "" and s.lower() != "nan" and s != "none"

    # Para cada fila, contar cuántas columnas ancla tienen dato real
    conteo = df[cols_presentes].apply(lambda row: sum(tiene_dato(v) for v in row), axis=1)

    # Una fila es válida si tiene dato en al menos 3 columnas ancla (ajustable)
    filas_validas = conteo >= 3

    if filas_validas.any():
        ultima_fila_valida = filas_validas[::-1].idxmax()
        return df.loc[:ultima_fila_valida]

    return df.iloc[0:0]  # DataFrame vacío si no hay filas válidas

#-------------------------------------------------------------------------------------------------------------
#Función para obtener los datos de los archivos que cumplen con la clave y extensión, aplicando validación de columnas y recorte de footer dinámico.
def obtener_datos(carpeta_archivos, clave, exts, omitir):
    datos = []
    try:    
        for archivo in Path(carpeta_archivos).iterdir():
            if archivo.is_file() and archivo.suffix.lower() in exts and str(clave).lower() in archivo.stem.lower() and not any(om.lower() in archivo.stem.lower() for om in omitir):
                
                escribir(f"Archivo encontrado: {archivo.name}")

                extension = archivo.suffix.lower()
                
                try:
                    if extension == ".xlsb":
                        try:
                            df = pd.read_excel(archivo, skiprows=6, dtype=str, engine='calamine')
                            escribir("Archivo .xlsb leído con éxito usando 'calamine'")
                        except ImportError:
                            df = pd.read_excel(archivo, skiprows=6, dtype=str, engine='pyxlsb')
                            escribir("Archivo .xlsb leído con éxito usando 'pyxlsb'")
                    elif extension in [".xlsx", ".xlsm"]:
                        df = pd.read_excel(archivo, skiprows=6, dtype=str, engine='openpyxl')
                        escribir(f"Archivo {extension} leído con éxito usando 'openpyxl'")
                    else:
                        df = pd.read_excel(archivo, skiprows=6, dtype=str)
                        escribir(f"Archivo {extension} leído con éxito usando el motor por defecto")
                except Exception as e:
                    escribir(f"Error al leer el archivo {archivo.name} con extensión {extension}: {str(e)}")
                    escribir("Se continuará con el siguiente archivo.")
                    continue
                df = recortar_footer_dinamico(df)  # Corte dinámico del footer
                validar_columnas(df, archivo.name)
                df["__archivo_origen"] = archivo.name
                
                df = df.reset_index(drop=True)
                datos.append(df)

        if not datos:
            raise Exception("No hay archivos válidos para procesar".upper())
        return datos
    except Exception as e:
        raise

#-------------------------------------------------------------------------------------------------------------
#Concatena la información del dataframe
def concatenar_datos():
    datos = obtener_datos(carpeta_archivos, clave, exts, omitir)
    df_total = pd.concat(datos)
    return df_total

#-------------------------------------------------------------------------------------------------------------
