import pandas as pd
from Configuracion_parametros import escribir, Campos_a_validar, largo_campos, ruta_error_largo_campos, ruta_alertas, ruta_columna_tipo, ruta_redondeo, ruta_inicio_campo, ruta_entidad_cuenta, ruta_filler, ruta_duplicados, justificacion_contable, ruta_justificacion_contable
from validacion.reportes import exportar_errores 

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
        escribir(f"⚠ Se encontraron alertas de campos vacíos. Revisar archivo: 04 - alertas_campos_vacios.xlsx")
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
def validar_entidad_cuenta(df):
    
    if "Entidad de la cuenta" not in df.columns:
        return df

    # Limpieza estándar
    df["Entidad de la cuenta"] = (
        df["Entidad de la cuenta"].fillna("").astype(str).str.strip().str.zfill(4)
    )

    # Detectar valores inválidos
    errores_tipo = df[
        ~(df["Entidad de la cuenta"] == "0013") & 
        (df["Entidad de la cuenta"] != "")
    ].copy()

    if not errores_tipo.empty:
        errores_tipo["Causal"] = "ERROR"                                          
        errores_tipo["campo_evento"] = "ENTIDAD DE LA CUENTA"
        errores_tipo["descripcion"] = "SOLO SE PERMITE EL CÓDIGO DE ENTIDAD '0013'"

        exportar_errores(                                                          #
            errores_tipo,
            ruta_entidad_cuenta,
            "⚠ Se encontraron valores inválidos en 'Entidad de la cuenta'. "
            "Revisar archivo: 07 - errores_entidad_de_cuenta.xlsx",
            sheet_name="Errores Entidad Cuenta"
        )
    else:
        escribir("✔ Validación de entidad completada sin errores.")

    return df

#-------------------------------------------------------------------------------------------------------------

def validar_filler(df):
    
    if "filler" not in df.columns:
        return df

    # Limpieza estándar
    df["filler"] = (
        df["filler"].fillna("").astype(str).str.strip()
    )

    # Detectar valores inválidos
    errores_tipo = df[
        ~(df["filler"] == "0") & 
        (df["filler"] != "")
    ].copy()

    if not errores_tipo.empty:
        errores_tipo["Causal"] = "ERROR"                                          
        errores_tipo["campo_evento"] = "FILLER"
        errores_tipo["descripcion"] = "SOLO SE PERMITE EL CÓDIGO DE FILLER '0'"

        exportar_errores(                                                          #
            errores_tipo,
            ruta_filler,
            "⚠ Se encontraron valores inválidos en 'Filler'. "
            "Revisar archivo: 08 - errores_filler.xlsx",
            sheet_name="Errores Filler"
        )
    else:
        escribir("✔ Validación de filler completada sin errores.")

    return df

#-------------------------------------------------------------------------------------------------------------
def validar_duplicados(df):
    
    # columnas de validación de duplicados
    columnas_llave =[
        'Entidad de la cuenta','Centro cuenta','filler','numero de la cuenta','tipo','valor ajuste','Cuenta a afectar','Justificacion contable','Cuentas contables contrapartida','Detalle del ajuste realizado','TIPO DE DOCUMENTO','NUMERO DE DOCUMENTO','DIGITO DE VERIFICACION'
    ]
    
    # iterar columnas
    cols_presentes = [c for c in columnas_llave if c in df.columns]
    # Detectar duplicados
    errores_tipo = df.duplicated(subset=cols_presentes, keep=False)
    errores_tipo = df[errores_tipo].copy()

    # Si hay errores, se exportan a Excel y se lanza una excepción. Si no, se imprime mensaje de validación exitosa.
    if not errores_tipo.empty:
        errores_tipo["Causal"] = "ALERTA"                                        
        errores_tipo["campo_evento"] = "DUPLICADOS"
        errores_tipo["descripcion"] = "EXISTEN DUPLICADOS EN EL ARCHIVO"

        exportar_errores(                                                          #
            errores_tipo,
            ruta_duplicados,
            "⚠ Se encontraron duplicados en el archivo. "
            "Revisar archivo: ALERTA_DUPLICADOS.xlsx",
            sheet_name="Alerta Duplicados"
        )
    else:
        escribir("✔ Validación de duplicados completada sin Alertas.")

    return df

#-------------------------------------------------------------------------------------------------------------

def validar_justificacion_contable(df):
    
    # Verificación correcta de columnas
    if "tipo" not in df.columns or "Justificacion contable" not in df.columns:
        return df

    # Limpieza estándar
    df["Justificacion contable"] = (
        df["Justificacion contable"].fillna("").astype(str).str.strip()
    )
    df["tipo"] = df["tipo"].fillna("").astype(str).str.strip()

    errores_totales = []

    # Validar por cada tipo (P y N) por separado
    for tipo_valor, justificaciones_validas in justificacion_contable.items():
        
        # Filtrar filas que corresponden a este tipo
        df_tipo = df[df["tipo"] == tipo_valor]
        
        # Detectar justificaciones que no están en la lista válida para ese tipo
        con_dato = df_tipo["Justificacion contable"] != ""
        errores_mask = con_dato & ~df_tipo["Justificacion contable"].isin(justificaciones_validas)
        errores = df_tipo[errores_mask].copy()

        if not errores.empty:
            errores["Causal"] = "ERROR"
            errores["campo_evento"] = "JUSTIFICACION CONTABLE"
            errores["descripcion"] = (
                f"JUSTIFICACIÓN NO PERMITIDA PARA TIPO '{tipo_valor}'"
            )
            errores_totales.append(errores)

    if errores_totales:
        exportar_errores(
            pd.concat(errores_totales),
            ruta_justificacion_contable,
            "⚠ Se encontraron valores inválidos en 'Justificacion contable'. "
            "Revisar archivo: 09 - errores_justificacion_contable.xlsx",
            sheet_name="Errores Justificacion Contable"
        )
    else:
        escribir("✔ Validación de justificación contable completada sin errores.")

    return df