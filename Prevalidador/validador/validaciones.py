import pandas as pd
from Configuracion_parametros import escribir, Campos_a_validar, largo_campos, ruta_error_largo_campos, ruta_alertas, ruta_columna_tipo, ruta_redondeo, ruta_inicio_campo, ruta_caracteres_especiales
from validador.reportes import exportar_errores

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
        exportar_errores(
            pd.concat(errores_totales),
            ruta_error_largo_campos,
            "⚠ Se encontraron errores de longitud. Revisar archivo: 02 - errores_largo_campos.xlsx",
            sheet_name="Errores Largo Campos"
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
        alertas_totales = pd.concat(alertas_totales)
        alertas_totales.index = alertas_totales.index + 8
        alertas_totales = alertas_totales.sort_index()

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

        exportar_errores(
            errores_tipo,
            ruta_columna_tipo,
            "⚠ Se encontraron valores inválidos en la columna 'tipo'. Revisar archivo: 03 - errores_columna_tipo.xlsx",
            sheet_name="Errores Columna Tipo"
        )
    else:
        escribir("✔ Validación de columna 'tipo' completada sin errores.")
    return df

#-------------------------------------------------------------------------------------------------------------
def validar_redondeo_valores(df):
    
    if "valor ajuste" not in df.columns:
        return df

    # Limpieza estándar antes de validar
    col_numerica = pd.to_numeric(df["valor ajuste"], errors="coerce")

    col_formateada = col_numerica.apply(lambda x: f"{x:.2f}" if pd.notna(x) else "")
    
    con_dato = col_formateada != ""
    errores_mask = con_dato & (col_numerica.round(2) != col_numerica)
    errores = df[errores_mask].copy()

    if not errores.empty:
        errores["Causal"] = "ERROR"
        errores["campo_evento"] = "VALOR AJUSTE"
        errores["descripcion"] = "VALOR NO TIENE FORMATO NUMÉRICO CON 2 DECIMALES"

        exportar_errores(
            errores,
            ruta_redondeo,
            "⚠ Se encontraron valores con decimales que no están redondeados a 2 decimales. "
            "Revisar archivo: 05 - errores_redondeo.xlsx",
            sheet_name="Errores Valor Ajuste"
        )
    else:
        escribir("✔ Validación de redondeo en campos de valor completada sin errores.")

    return df

#-------------------------------------------------------------------------------------------------------------
def validar_inicio_numero_cuenta(df, campo: str, prefijo):
    if campo not in df.columns:
        return df

    col = df[campo].fillna("").astype(str).str.strip()
    con_dato = col != ""
    errores_mask = con_dato & ~col.str.startswith(prefijo)
    errores = df[errores_mask].copy()

    if not errores.empty:
        errores["Causal"] = "ERROR"
        errores["campo_evento"] = campo.upper()
        errores["descripcion"] = f"EL CAMPO DEBE INICIAR POR {prefijo}"

        exportar_errores(
            errores,
            ruta_inicio_campo,
            f"⚠ Se encontraron valores en '{campo}' que no inician por {prefijo}. "
            f"Revisar archivo: 06 - errores_inicio_numero_cuenta.xlsx",
            sheet_name="Errores Inicio Campo"
        )
    else:
        escribir(f"✔ Validación de inicio de campo '{campo}' completada sin errores.")

    return df

#-------------------------------------------------------------------------------------------------------------

def validar_caracteres_especiales(df):
    errores_totales = []

    #filtrar columnas de interés
    campos_filtrados = [
        campo for campo in Campos_a_validar
        if campo not in {"Unnamed: 0", "Unnamed: 1", "Detalle del ajuste realizado"}
    ]

    # Validar solo campos con dato
    for nombre_campo in campos_filtrados:         
        if nombre_campo not in df.columns:
            continue
        serie = df[nombre_campo].fillna("").astype(str).str.strip()  
        con_dato = serie != ""
        errores_mask = con_dato & serie.str.contains(r"[^a-zA-ZáéíóúüñÁÉÍÓÚÜÑ0-9\s]", regex=True)
        errores = df[errores_mask].copy()
        # Solo escribimos Excel si HAY errores
        if not errores.empty:
            errores["Causal"] = "ERROR"
            errores["campo_evento"] = nombre_campo.upper()  
            errores["descripcion"] = "EXISTEN CARACTERES ESPECIALES EN EL CAMPO"
            errores_totales.append(errores)
        else:
            escribir(f"✔ Validación de caracteres especiales en '{nombre_campo}' completada sin errores.")

    # Si hay errores en alguno de los campos, se agrega informacion de causal y descripcion para el reporte
    if errores_totales:
        exportar_errores(
            pd.concat(errores_totales),
            ruta_caracteres_especiales,
            "⚠ Se encontraron caracteres especiales en uno o más campos. "
            "Revisar archivo: 07 - errores_caracteres_especiales.xlsx",
            sheet_name="Errores Caracteres Especiales"
        )