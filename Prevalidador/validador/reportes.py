import os
import pandas as pd
from Configuracion_parametros import ruta_error_largo_campos, ruta_alertas, ruta_columna_tipo, ruta_redondeo, log_exitoso

#-------------------------------------------------------------------------------------------------------------
#Función para borrar archivos temporales de validación (si existen)
def borrar_archivos_temporales():
    archivos_temporales = [
        ruta_error_largo_campos, #eliminar el archivo de errores de largo de campos
        ruta_columna_tipo,       #eliminar el archivo de errores de columna tipo
        ruta_alertas,            #eliminar el archivo de alertas de campos vacíos
        ruta_redondeo,           #eliminar el archivo de errores de redondeo en campos de valor
        log_exitoso              #eliminar el archivo de log de validación exitosa
    ]
    for ruta in archivos_temporales:
        if os.path.exists(ruta):
            os.remove(ruta)

#-------------------------------------------------------------------------------------------------------------
#Función generalizada para exportar errores a Excel y lanzar excepción
def exportar_errores(df_errores: pd.DataFrame, ruta: str, mensaje: str, sheet_name: str = "Errores") -> None:
    df_errores = df_errores.copy()
    df_errores.index = df_errores.index + 8
    df_errores = df_errores.sort_index()

    # Crear carpeta si no existe
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
