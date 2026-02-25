from pathlib import Path
import pandas as pd
from Configuracion_parametros import carpeta_archivos, escribir, exts, clave, Campos_a_validar, largo_campos, ruta_error_largo_campos, ruta_alertas, ruta_columna_tipo, log_exitoso
import os, time as t

#libreria para funciones de formatea para nombres de columnas extra
import re
from openpyxl.utils import get_column_letter


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

#-------------------------------------------------------------------------------------------------------------
#Obtener los archivos de OPS + columna de nombre del archivo
def obtener_datos(carpeta_archivos, clave, exts):
    datos = []
    try:    
        for i in Path(carpeta_archivos).iterdir():
            if i.is_file() and i.suffix.lower() in exts and str(clave).lower() in i.stem.lower():
            # lee cada archivo de OPS que se encuentre en la ruta
                escribir(f"Archivo encontrado: {i.name}")
                t.sleep(0.5)
                df = pd.read_excel(i, skiprows=6, skipfooter=7 ,dtype=str)
                validar_columnas(df, i.name) #validar las columnas del archivo
                df["__archivo_origen"] = i.name
                datos.append(df)

        if not datos:
            raise Exception("No hay archivos válidos para procesar".upper())
        return datos
    except Exception as e:
        # Si ocurre un error (como no encontrar archivos o columnas faltantes), se captura la excepción, se imprime el mensaje de error y se detiene el proceso.
        raise

#-------------------------------------------------------------------------------------------------------------
#Concatena la información del dataframe
def concatenar_datos():
    datos = obtener_datos(carpeta_archivos, clave, exts)
    df_total = pd.concat(datos, ignore_index=True)
    return df_total

#-------------------------------------------------------------------------------------------------------------
#Función para validar el largo de campos específicos
def validar_largo_campos(df):
    
    errores_totales = []

    for campo, regla in largo_campos.items():

        if campo not in df.columns:
            continue

        # Limpieza
        df[campo] = (
            df[campo]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
            .str.strip()
        )

        # Validar solo campos con dato
        df_validar = df[df[campo] != ""]
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
        # Si hay errores en este campo, se agrega información de causal y descripción para el reporte.
        if not errores.empty:          
            errores = errores.copy()
            errores["Causal"] = "ERROR".upper()
            errores["campo_evento"] = campo.upper()
            errores["descripcion"] = "Longitud incorrecta".upper()
            errores_totales.append(errores)
            
    # Si hay errores, se concatenan y se exportan a Excel. Luego se lanza una excepción para detener el proceso.        
    if errores_totales:
        errores_totales = pd.concat(errores_totales)
        errores_totales.index = errores_totales.index + 8
        errores_totales.to_excel(
            ruta_error_largo_campos,
            index=True,
            index_label="Fila en Excel"
        )

        raise Exception(
            f"⚠ Se encontraron errores de longitud. Revisar archivo: 02 - errores_largo_campos.xlsx"
        )
    # Si no hay errores, se imprime mensaje de validación exitosa.
    else:
        escribir("✔ Validación de largo de campos completada sin errores.")

    return df



#-------------------------------------------------------------------------------------------------------------
# validar campos no vacíos según regla de campos a validar, generando un reporte de alertas si se encuentran campos vacíos. Se asume que los primeros 2 campos no requieren validación de vacíos.
def validar_campos_vacios(df):
    alertas_totales = []

    for campo in Campos_a_validar[2:]:
        if campo not in df.columns:
            continue

        df[campo] = (
            df[campo]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0+$", "", regex=True)
            .str.strip()
        )

        alertas = df[df[campo] == ""]
        if not alertas.empty:
            alertas = alertas.copy()
            alertas["Causal"] = "ALERTA".upper()
            alertas["campo_evento"] = campo.upper()
            alertas["descripcion"] = "Campo vacío".upper()
            alertas_totales.append(alertas)



    # Solo escribimos Excel si HAY alertas
    if alertas_totales:
        alertas_totales = pd.concat(alertas_totales, ignore_index=True)
        alertas_totales.index = alertas_totales.index + 8

        with pd.ExcelWriter(ruta_alertas, engine="openpyxl", mode="w") as writer:
            alertas_totales.to_excel(
                writer,
                sheet_name="Alertas Campos Vacíos",
                index=True,
                index_label="Fila en Excel",
            )
        escribir(f"⚠ Se encontraron alertas de campos vacíos. Revisar archivo: 04 - alertas_campos.xlsx")
    else:
        escribir("✔ No se encontraron campos vacíos.")

    return df

#-------------------------------------------------------------------------------------------------------------

#Función para validar que la columna 'tipo' solo contenga 'P' o 'N' (en mayúscula), generando un reporte de errores si se encuentran valores inválidos.
def validar_columna_tipo(df):

    if "tipo" not in df.columns:
        return df

    df["tipo"] = df["tipo"].fillna("").astype(str).str.strip()
    # Detectar valores inválidos (incluye minúsculas)
    errores_tipo = df[(~df["tipo"].isin(["P", "N"])) & (df["tipo"] != "")]

    # Si hay errores, se exportan a Excel y se lanza una excepción. Si no, se imprime mensaje de validación exitosa.
    if not errores_tipo.empty:

        errores_tipo = errores_tipo.copy()
        errores_tipo["campo_evento"] = "tipo".upper()
        errores_tipo["descripcion"] = "Solo se permiten valores 'P' o 'N' en mayúscula".upper()
        
        errores_tipo.index = errores_tipo.index + 8



        with pd.ExcelWriter(
            ruta_columna_tipo,
            engine="openpyxl",
        ) as writer:

            errores_tipo.to_excel(
                writer,
                sheet_name="Errores Columna Tipo",
                index=True,
                index_label="Fila en Excel"
            )

        raise Exception(
            f"⚠ Se encontraron valores inválidos en la columna 'tipo'. Revisar archivo: 03 - errores_columna_tipo.xlsx"
        )
    else:
        escribir("✔ Validación de columna 'tipo' completada sin errores.")
    return df

#-------------------------------------------------------------------------------------------------------------
#Función para borrar archivos temporales de validación (si existen)
def borrar_archivos_temporales():
    archivos_temporales = [
        ruta_error_largo_campos, #eliminar el archivo de errores de largo de campos
        ruta_columna_tipo, #eliminar el archivo de errores de columna tipo
        ruta_alertas,#eliminar el archivo de alertas de campos vacíos
        log_exitoso #eliminar el archivo de log de validación exitosa
    ]
    for ruta in archivos_temporales:
        if os.path.exists(ruta):
            os.remove(ruta)
#-------------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------
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


#-------------------------------------------------------------------------------------------------------------


