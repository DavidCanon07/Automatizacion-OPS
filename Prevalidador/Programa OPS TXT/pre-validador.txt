from validador import concatenar_datos, validar_largo_campos, validar_columna_tipo, validar_campos_vacios, validar_redondeo_valores, borrar_archivos_temporales, validar_inicio_numero_cuenta, validar_entidad_cuenta, validar_filler, validar_duplicados, validar_justificacion_contable
from Configuracion_parametros import ruta_error_txt, crear_carpeta_si_no_existe, base_path, carpeta_archivos, log_exitoso, escribir, ruta_duplicados
from datetime import datetime
import os



try:

    escribir("¡Hola!:D Bienvenido al pre-validador de archivos de OPS. Iniciando proceso de validación...\n"
        "Primero, antes de comenzar limpiemos la información de validaciones anteriores para tener un entorno limpio.\n")
    # Crear carpetas necesarias para el proceso (si no existen)
    escribir("Verificando que las carpetas necesarias para el proceso existan...\n")
    crear_carpeta_si_no_existe(base_path, carpeta_archivos)
    # Limpiar archivos de validaciones anteriores
    borrar_archivos_temporales()
    
    # 01 - CARGA Y VALIDACIÓN DE COLUMNAS
    escribir("Ahora, vamos a cargar el archivo y validar que las columnas sean correctas...\n")
    df = concatenar_datos()
    
    escribir("Validación de columnas completada exitosamente."
        "\nPerfecto todo marcha bien, Empecemos la validación de campos...\n")
    
    # Eliminar columnas no deseadas
    df = df.drop(columns=["Unnamed: 0", "Unnamed: 1"], errors="ignore")
    
    # 02 - VALIDACIÓN DE LARGO DE CAMPOS
    try:
        escribir("Validando largo de campos...\n")
        validar_largo_campos(df)
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("Vaya! Se han encontrado errores en la validación de largo de campos.\n")
        escribir(error_msg)
        
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------02 - Validación de largo de campos.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
            
        exit()

    # 03 - VALIDACIÓN DE CAMPOS ESPECÍFICOS (Ejemplo: columna 'tipo')
    try:
        
        escribir("Validando valores de columna 'tipo'...\n")
        validar_columna_tipo(df)

    except Exception as e:
        error_msg = str(e)
        
        escribir("Ups! Se han encontrado errores en la validación de la columna 'tipo'.\n")
        escribir(error_msg)

        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------03 - Validación de campos específicos.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")

        exit()
    #04 - VALIDACIÓN DE CAMPOS VACÍOS
    try:
        
        escribir("Validando campos vacíos...\n")
        validar_campos_vacios(df)
    except Exception as e:
        error_msg = str(e)
        
        escribir("Oh no! Se han encontrado campos vacíos que requieren atención.\n")
        escribir(error_msg)
    # En este caso, como es una validación de alertas (no errores críticos), se registran los campos vacíos encontrados pero no se detiene el proceso. Se asume que el usuario revisará el archivo de alertas generado para corregir estos campos antes de montar el archivo en la ruta de la OPS.
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------04 - Validación de campos vacíos.--------\n")
            f.write(f"\nLog de alerta: {datetime.now()} - {error_msg}\n")

        exit()
    #05 VALIDACIÓN DE DECIMALES EN CAMPOS DE VALOR
    try:
        escribir("Validando redondeo de valores en campos de valor...\n")
        validar_redondeo_valores(df)
    except Exception as e:
        error_msg = str(e)
        
        escribir("valores con decimales que no estan redondeados a 2 decimales Detectados.\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------05 - Validación de decimales en campos de valor.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
        
        exit()
    
    #06 VALIDACIÓN DE INICIO DE CAMPO NUMERO DE CUENTA
    try:
        escribir("Validando inicio de campo 'numero de la cuenta'...\n")
        validar_inicio_numero_cuenta(df, "numero de la cuenta", ("1","2"))
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("La columna de numero de la cuenta no inicia por los prefijos permitidos 1(Cuenta de ahorros) o 2(Cuenta corriente).\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------06 - Validación de inicio de campo 'numero de la cuenta'.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
            
        exit()
        
    #07 VALIDACIÓN DE ENTIDAD CUENTA
    try:
        escribir("Validando el campo 'entidad de la cuenta'...\n")
        validar_entidad_cuenta(df)
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("Se encontraron valores no permitidos en la columna 'entidad de la cuenta' (diferentes de 0013).\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------07 - errores_entidad_de_cuenta.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
            
        exit()
        
    #08 VALIDACIÓN DE FILLER
    try:
        escribir("Validando el campo 'filler'...\n")
        validar_filler(df)
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("Se encontraron valores no permitidos en la columna 'filler' (diferentes de 0).\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------08 - errores_filler.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
            
        exit()
        
    #09 VALIDACIÓN DE TIPO
    try:
        escribir("Validando el campo 'justificacion contable' sea coincidente con el campo 'tipo'...\n")
        validar_justificacion_contable(df)
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("Se encontraron valores no permitidos y/o no coincidentes en la columna 'justificacion contable' con el campo 'tipo' (diferentes a los establecidos).\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------09 - errores_justificacion_contable.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
            
        exit()
        
    # VALIDACIÓN DE DUPLICADOS
    try:
        escribir("Validando la tapa de OPS para verificar duplicados...\n")
        validar_duplicados(df)
        
    except Exception as e:
        error_msg = str(e)
        
        escribir("Se encontraron valores duplicados -> 🚧¡REVISA TU ARCHIVO ANTES DE MONTAR EN LA RUTA! Se pueden encontrar transacciones duplicadas fuera del proceso habitual.🚧\n")
        escribir(error_msg)
        with open(
            ruta_error_txt,
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------ALERTA_DUPLICADOS.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")


#------------------------------------------------------------------------------
    # FINALIZACIÓN DEL PROCESO DE VALIDACIÓN -> valida que no tenga alertas de duplicados y que no haya errores críticos
    if os.path.exists(ruta_duplicados):
        escribir(f"\nRevisa el archivo alertas de duplicados antes de montar el archivo en la ruta de la OPS para evitar posibles duplicados fuera del proceso normal.\n")
        print("👀COMO SE PRESENTO UNA DUPLICIDAD Y VALIDA QUE SEA NORMAL(archivo 'ALERTAS_DUPLICADOS.xlsx')👀, DESPUES SIGUE LAS RECOMENDACIONES:\n"
            "1. Asegúrate de que en la ruta se crea la carpeta de tu proceso\n"
            "2. Verifica que las Tapas de OPS se monten en la ruta correcta y con el formato correcto\n"
            f"3. El archivo debe llevar la fecha del día (Ejemplo: OPS 'Nombre proceso' {datetime.now().strftime('%d-%m-%Y')}.xlsx)\n".upper())
    else:
        escribir(f"\nPerfecto! No se han encontrado errores en las validaciones(no errores críticos). El archivo está listo para ser montado en la ruta de la OPS.\n"
            "Recuerda revisar el archivo Log.txt y te comparto recomendaciones generales: \n")
        print("1. Asegúrate de que en la ruta se crea la carpeta de tu proceso\n"
            "2. Verifica que las Tapas de OPS se monten en la ruta correcta y con el formato correcto\n"
            f"3. El archivo debe llevar la fecha del día (Ejemplo: OPS 'Nombre proceso' {datetime.now().strftime('%d-%m-%Y')}.xlsx)\n".upper())
        with open(
            log_exitoso,
            "a", encoding="utf-8"
        ) as f:
            f.write(f"\nLog de validación exitosa: {datetime.now()} - Perfecto! No se han encontrado errores en las validaciones. El archivo está listo para ser montado en la ruta de la OPS.\n"
                    "\nalgunas recomendaciones: \n"
                    "1. Asegúrate de que en la ruta se crea la carpeta de tu proceso\n"
                    "2. Verifica que las Tapas de OPS se monten en la ruta correcta y con el formato correcto\n"
                    f"3. El archivo debe llevar la fecha del día (Ejemplo: OPS 'Nombre proceso' {datetime.now().strftime('%d-%m-%Y')}.xlsx)\n".upper())

#------------------------------------------------------------------------------
# EXCEPCIÓN GENERAL (ERRORES DE INGESTA O COLUMNAS)
except Exception as e:
    escribir("\n✕ ERROR DURANTE EL PROCESO DE VALIDACIÓN\n")
    escribir(str(e))
    
    with open(
        ruta_error_txt,
        "a", encoding="utf-8"
    ) as f:
        f.write("\n-------01 - Validación de columnas Y entrada de archivos.--------\n")
        f.write(f"\nLog de error: {datetime.now()} - {e}\n")
        
    escribir("El error ha sido registrado en errores.txt")
    exit()
