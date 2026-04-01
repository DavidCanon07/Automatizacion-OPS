import os
import glob
import pandas as pd
from Configuracion_parametros import ruta_error_largo_campos, ruta_alertas, ruta_columna_tipo, ruta_redondeo, log_exitoso, ruta_inicio_campo, ruta_entidad_cuenta, ruta_filler, ruta_duplicados, ruta_justificacion_contable, ruta_archivo_unificado, ruta_archivo_debitos, ruta_archivo_plano_txt
#-------------------------------------------------------------------------------------------------------------
#Función para borrar archivos temporales de validación (si existen)
def borrar_archivos_temporales():
    archivos_temporales = [
        ruta_error_largo_campos, #eliminar el archivo de errores de largo de campos
        ruta_columna_tipo,       #eliminar el archivo de errores de columna tipo
        ruta_alertas,            #eliminar el archivo de alertas de campos vacíos
        ruta_redondeo,           #eliminar el archivo de errores de redondeo en campos de valor
        ruta_inicio_campo,       #eliminar el archivo de errores por número de cuenta
        ruta_entidad_cuenta,     #eliminar el archivo de errores por entidad de cuenta
        ruta_filler,             #eliminar el archivo de errores por filler
        ruta_justificacion_contable, #eliminar el archivo de errores por justificacion contable
        ruta_duplicados,         #eliminar el archivo de errores por duplicados
        ruta_archivo_unificado,  #eliminar el archivo de OPS unificada
        ruta_archivo_debitos,    #eliminar el archivo de débitos unificado
        ruta_archivo_plano_txt,  #eliminar el archivo plano de OPS
    ]
    for ruta in archivos_temporales:
        if os.path.exists(ruta):
            os.remove(ruta)

#-------------------------------------------------------------------------------------------------------------
#Función para borrar archivos de formato OPS en la carpeta de formatos (si existen)
def borrar_archivo_carpeta_formato_ops(formatos):
    folder_path = formatos
    search_pattern = os.path.join(folder_path, "*.xlsx")
    files_to_delete = glob.glob(search_pattern)
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
        except OSError as e:
            print(f"Error al eliminar {file_path}: {e}")
            
#---------------------
def borrar_carpeta_comprimido(formatos):
    folder_path = formatos
    search_pattern = os.path.join(folder_path, "*.zip")
    files_to_delete = glob.glob(search_pattern)
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
        except OSError as e:
            print(f"Error al eliminar {file_path}: {e}")
#-------------------------------------------------------------------------------------------------------------
#Función generalizada para exportar errores a Excel y lanzar excepción
def exportar_errores(df_errores: pd.DataFrame, ruta: str, mensaje: str, sheet_name: str = "Errores") -> None:
    df_errores = df_errores.copy()
    df_errores.index = df_errores.index + 8
    df_errores = df_errores.sort_index()

    #Crear carpeta si no existe
    os.makedirs(os.path.dirname(ruta), exist_ok=True)

    with pd.ExcelWriter(ruta, engine="openpyxl", mode="w") as writer:
        df_errores.to_excel(
            writer,
            sheet_name=sheet_name,
            index=True,
            index_label="Fila en Excel"
        )

    raise Exception(mensaje)

#-------------------------------------------------------------------------------------------------------------
