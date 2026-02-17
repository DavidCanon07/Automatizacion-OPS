from Ingesta import concatenar_datos, validar_largo_campos, validar_campos_vacios, validar_columna_tipo, borrar_archivos_temporales
from datetime import datetime
import time as t


try:
    t.sleep(1)
    print("¡Hola!!👋 Bienvenido al pre-validador de archivos de OPS. Iniciando proceso de validación...\n"
        "Primero, antes de comenzar limpiemos la información de validaciones anteriores para tener un entorno limpio.")
    # Limpiar archivos de validaciones anteriores
    borrar_archivos_temporales()
    t.sleep(2)
    # 01 - CARGA Y VALIDACIÓN DE COLUMNAS
    print("Ahora, vamos a cargar el archivo y validar que las columnas sean correctas...")
    df = concatenar_datos()
    t.sleep(1)
    print("Validación de columnas completada exitosamente."
        "\nPerfecto todo marcha bien,\n" 
        "Empecemos la validación de campos...")

    # Eliminar columnas no deseadas
    df = df.drop(columns=["Unnamed: 0", "Unnamed: 1"], errors="ignore")

    # 02 - VALIDACIÓN DE LARGO DE CAMPOS
    try:
        t.sleep(1)
        print("Validando largo de campos...")
        validar_largo_campos(df)

    except Exception as e:
        error_msg = str(e)
        t.sleep(1)
        print("Vaya! Se han encontrado errores en la validación de largo de campos.")
        print(error_msg)

        with open(
            r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\errores.txt",
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------02 - Validación de largo de campos.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")

        exit()

    # 03 - VALIDACIÓN DE CAMPOS ESPECÍFICOS (Ejemplo: columna 'tipo')
    try:
        t.sleep(1)
        print("Validando valores de columna 'tipo'...")
        validar_columna_tipo(df)

    except Exception as e:
        error_msg = str(e)
        t.sleep(1)
        print("Ups! Se han encontrado errores en la validación de la columna 'tipo'.")
        print(error_msg)

        with open(
            r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\errores.txt",
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------03 - Validación de campos específicos.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")

        exit()
    #04 - VALIDACIÓN DE CAMPOS VACÍOS
    try:
        t.sleep(1)
        print("Validando campos vacíos...")
        validar_campos_vacios(df)
    except Exception as e:
        error_msg = str(e)
        t.sleep(1)
        print("Oh no! Se han encontrado campos vacíos que requieren atención.")
        print(error_msg)

        with open(
            r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\errores.txt",
            "a", encoding="utf-8"
        ) as f:
            f.write("\n-------04 - Validación de campos específicos.--------\n")
            f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")

        exit()

    # Si todo salió bien, puede guardar un Log de validación exitosa.
    print("Perfecto! No se han encontrado errores en las validaciones. El archivo está listo para ser montado en la ruta de la OPS.\n"
        "Recuerda revisar el archivo Log.txt y te recuerdo recomendaciones generales: \n")
    print("1. Asegúrate de que en la ruta se crea la carpeta de tu proceso\n"
        "2. Verifica que las Tapas de OPS se monten en la ruta correcta y con el formato correcto\n"
        f"3. El archivo debe llevar la fecha del día (Ejemplo: OPS 'Nombre proceso' {datetime.now().strftime('%d-%m-%Y')}.xlsx)\n".upper())
    with open(
        r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\log.txt",
        "a", encoding="utf-8"
    ) as f:
        f.write(f"\nLog de validación exitosa: {datetime.now()} - Perfecto! No se han encontrado errores en las validaciones. El archivo está listo para ser montado en la ruta de la OPS.\n"
                "\nalgunas recomendaciones: \n"
                "1. Asegúrate de que en la ruta se crea la carpeta de tu proceso\n"
                "2. Verifica que las Tapas de OPS se monten en la ruta correcta y con el formato correcto\n"
                f"3. El archivo debe llevar la fecha del día (Ejemplo: OPS 'Nombre proceso' {datetime.now().strftime('%d-%m-%Y')}.xlsx)\n".upper())
        
# EXCEPCIÓN GENERAL (ERRORES DE INGESTA O COLUMNAS)
except Exception as e:
    error_msg = str(e)
    print(error_msg)
    
    with open(
        r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\errores.txt",
        "a", encoding="utf-8"
    ) as f:
        f.write("\n-------01 - Validación de columnas Y entrada de archivos.--------\n")
        f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
        
    print("El error ha sido registrado en errores.txt")
    exit()