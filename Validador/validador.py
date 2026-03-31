from validacion import concatenar_datos, validar_largo_campos, validar_columna_tipo, validar_campos_vacios, validar_redondeo_valores, borrar_archivos_temporales, validar_inicio_numero_cuenta, validar_entidad_cuenta, validar_filler, validar_duplicados, validar_justificacion_contable, borrar_archivo_carpeta_formato_ops
from Consolidacion import exportar_excel, cargar_estructura
from Consolidacion.Debitos import obtener_debitos, exportar_debitos
from Configuracion_parametros import ruta_error_txt, crear_carpeta_si_no_existe, base_path, carpeta_archivos, log_exitoso, escribir, ruta_duplicados, ruta_archivo_unificado, ejecucion, ruta_libro_base, hoja_base, MAPEO, ruta_formato_ops, formatos, clave_debitos, exts, clave, hoja_base_debitos, ruta_archivo_debitos, ruta_libro_base_debitos, MAPEO_DEBITOS, ruta_formato_debitos
from datetime import datetime
import os
import zipfile
import shutil


try:

    escribir("Iniciando proceso de validación de archivos para OPS...\n"
            "Primero, vamos a asegurarnos de que las carpetas necesarias para el proceso existan y estén listas para recibir los archivos...\n")
    # Crear carpetas necesarias para el proceso (si no existen)
    escribir("Verificando que las carpetas necesarias para el proceso existan...\n")
    crear_carpeta_si_no_existe(base_path, carpeta_archivos, ejecucion, formatos)
    # Limpiar archivos de validaciones anteriores
    escribir("Limpiando archivos de validaciones anteriores (si existen)...\n")
    borrar_archivos_temporales()
    escribir("Archivos temporales del paquete de la OPS (si existían).\n")
    borrar_archivo_carpeta_formato_ops(formatos)
    while True:
        opcion = input("Selecciona una opción para continuar:"
            "\n1. Consolidar archivos OPS y validar archivos"
            "\n2. Exportar archivo unificado"
            "\n3. Montar archivo excel 'Formato OPS DDMMYYYY' con los datos consolidados"
            "\n4. Consolidar Archivo Debitos 'Detalle_LATAM_SODIMAC_FALABELLA_MERCADO PAGO_DDMMYYYY'\n"
            "5. Salir\n"
            "\nSelecciona una opción: ")
        if opcion == "1":
            # 01 - CARGA Y VALIDACIÓN DE COLUMNAS
            escribir("Ahora, vamos a cargar el archivo y validar que las columnas sean correctas...\n")
            df = concatenar_datos(carpeta_archivos, clave, exts)
            
            escribir("Validación de columnas completada exitosamente."
                "\nPerfecto todo marcha bien, Empecemos la validación de campos...\n")
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
                # FINALIZACIÓN DEL PROCESO DE VALIDACIÓN -> valida que no tenga alertas de duplicados y que no haya errores críticos
            if os.path.exists(ruta_duplicados):
                escribir(f"\nRevisa el archivo alertas de duplicados antes de montar el archivo en la ruta de la OPS para evitar posibles duplicados fuera del proceso normal.\n")
                escribir("👀COMO SE PRESENTO UNA DUPLICIDAD Y VALIDA QUE SEA NORMAL(archivo 'ALERTAS_DUPLICADOS.xlsx')👀, DESPUES SIGUE LOS PROXIMOS PASOS:\n")
            else:
                escribir(f"\nPerfecto! No se han encontrado errores en las validaciones(no errores críticos). Los archivos estan listos para seguir los siguientes pasos del validador\n"
                    "Recuerda revisar el archivo Log.txt y para tener la traza historica de ejecuciones.\n")
                with open(
                    log_exitoso,
                    "a", encoding="utf-8"
                ) as f:
                    f.write(f"\nLog de validación exitosa: {datetime.now()} - Perfecto! No se han encontrado errores en las validaciones. Los archivos estan listos para seguir los siguientes pasos del validador\n")
            continue  # Volver al menú principal para que el usuario decida qué hacer a continuación
        elif opcion == "2":
            escribir("Exportando archivo unificado...")
            exportar_excel(df, ruta_archivo_unificado, "Formato OPS")
            escribir(f"Archivo exportado a: {ruta_archivo_unificado}")
            
            # VALIDACIÓN INMEDIATA DEL ARCHIVO EXPORTADO
            if not os.path.exists(ruta_archivo_unificado):
                raise Exception(f"✖️ El archivo no se creó: {ruta_archivo_unificado}")

            tamaño = os.path.getsize(ruta_archivo_unificado)
            escribir(f"✔️ Archivo creado. Tamaño: {tamaño} bytes")

            if tamaño == 0:
                raise Exception(f"✖️ El archivo está vacío: {ruta_archivo_unificado}")

            # Verificar que es un Excel válido (ZIP)
            try:
                with zipfile.ZipFile(ruta_archivo_unificado, 'r') as zf:
                    escribir("✔️ Archivo Excel válido (ZIP)")
            except zipfile.BadZipFile:
                raise Exception(f"✖️ El archivo {ruta_archivo_unificado} no es un Excel válido")
        elif opcion == "3":
            escribir("Montando archivo en formato 'Formato OPS'...")
            
            # Verificar que el archivo unificado existe
            if not os.path.exists(ruta_archivo_unificado):
                escribir(f"✖️ El archivo unificado no existe: {ruta_archivo_unificado}")
                escribir("   Primero debes exportar el archivo unificado (opción 2)")
                continue
            
            # Verificar que la plantilla existe
            if not os.path.exists(ruta_libro_base):
                escribir(f"✖️ La plantilla no existe: {ruta_libro_base}")
                continue
            
            # Copiar la plantilla (limpia) al destino con el nombre con fecha
            escribir(f"   Copiando plantilla limpia: {ruta_libro_base}")
            escribir(f"   A: {ruta_formato_ops}")
            shutil.copy2(ruta_libro_base, ruta_formato_ops)
            
            # Cargar los nuevos datos directamente (sin limpiar, porque es copia nueva)
            escribir(f"   Cargando datos al archivo destino...")
            cargar_estructura(ruta_archivo_unificado, ruta_formato_ops, hoja_base, MAPEO, fila_encabezados=8)
            
            escribir(f"\nArchivo montado exitosamente:")
            escribir(f"   Ubicación: {ruta_formato_ops}")
            escribir(f"   Hoja: {hoja_base}")
            
        elif opcion == "4":
            
            # Consolida la información de los archivos de debitos
            escribir("Consolidando archivo de debitos...")
            df_total_debitos = obtener_debitos(carpeta_archivos, clave_debitos, exts)
            
            # Exportar el archivo de debitos consolidado temporal
            escribir("Exportando archivo de debitos consolidado...")
            exportar_debitos(df_total_debitos, ruta_archivo_debitos, hoja_base_debitos)
            
            #Verificar que la plantilla de debitos exista
            if not os.path.exists(ruta_libro_base_debitos):
                escribir(f"✖️ La plantilla de debitos no existe: {ruta_libro_base_debitos}")
                continue
            
            # Copiar la plantilla (limpia) al destino con el nombre con fecha
            escribir(f"   Copiando plantilla limpia: {ruta_libro_base_debitos}")
            escribir(f"   A: {ruta_formato_debitos}")
            shutil.copy2(ruta_libro_base_debitos, ruta_formato_debitos)
            
            # Cargar los nuevos datos directamente (sin limpiar, porque es copia nueva)
            escribir(f"   Cargando datos al archivo destino de debitos...")
            cargar_estructura(ruta_archivo_debitos, ruta_formato_debitos, hoja_base_debitos, MAPEO_DEBITOS, fila_encabezados=1)
            
            escribir(f"\nArchivo de debitos montado exitosamente:")
            escribir(f"   Ubicación: {ruta_formato_debitos}")
            escribir(f"   Hoja: {hoja_base_debitos}")
            
            
        elif opcion == "5":
            escribir("Saliendo del programa. ¡Hasta luego!")
            exit()
        else:
            escribir("Opción no válida. Por favor, selecciona 1, 2, 3 o 4.")
    

#------------------------------------------------------------------------------


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
