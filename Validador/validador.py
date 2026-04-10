from validacion import *
from validacion.utils import *
from Consolidacion import exportar_excel, cargar_estructura, copiar_carpeta_ops
from Consolidacion.Debitos import obtener_debitos, exportar_debitos
from Consolidacion.Solicitud_ops import *
from Consolidacion.comprimir import comprimir_excel
from Historico_resumen.carga_historico import *
from Configuracion_parametros import *
from preguntas import *
from datetime import datetime
import os
import zipfile
import shutil


try:
    print("="*60)
    escribir(f"\nBienvenido al programa🗝️VALIDADOR🗝️ de archivos para OPS.\n".upper())
    print("="*60)
    escribir("\nEste proceso te ayudará a validar y consolidar tus archivos de manera eficiente y confiable.\n")
    escribir("Iniciando proceso de validación de archivos para OPS...\n"
            "Primero, vamos a asegurarnos de que las carpetas necesarias para el proceso existan y estén listas para recibir los archivos...\n")
    # Crear carpetas necesarias para el proceso (si no existen)
    escribir("Verificando que las carpetas necesarias para el proceso existan...\n")
    crear_carpeta_si_no_existe(base_path, carpeta_archivos, ejecucion, formatos, estructura, historico)
    # Validar estructuras base para la consolidación (si no existen, copiarlas desde la ruta_base_estructuras)
    validar_estructuras(ruta_libro_base, ruta_libro_base_debitos, ruta_libro_base_solicitud, ruta_consolidado_historico, formato_ops, formato_debitos, formato_solicitud, archivo_historico)
    # Limpiar archivos de validaciones anteriores
    escribir("Limpiando archivos de validaciones anteriores (si existen)...\n")
    borrar_archivos_temporales()
    escribir("Archivos temporales del paquete de la OPS (si existen).\n")
    borrar_archivo_carpeta_formato_ops(formatos)
    borrar_carpeta_comprimido(formatos)
    
    print("="*85)
    # Validar requisitos mínimos para la consolidación
    validar_requisitos_consolidacion()
    
    # Copiar la carpeta de la OPS del dia a la carpeta de carpetas_OPS
    escribir(f"Ahora vamos a copiar la carpeta de la OPS del día 'OPS {datetime.now().strftime('%d-%m-%Y')}' a la carpeta 'carpetas_OPS'.\n")
    copiar_carpeta_ops(ruta_dinamica, ruta_destino_desktop)
    
    #varible para contar número de alertas en la validación
    errores_totales = 0
    
    while True:
        print("\n" + "="*85)
        opcion = input(f"\nSelecciona una opción para continuar:"
            "\n1. Consolidar archivos OPS y validar archivos"
            "\n2. Exportar archivo unificado"
            "\n3. Montar archivo excel 'Formato OPS DDMMYYYY' con los datos consolidados"
            "\n4. Consolidar Archivo Debitos 'Detalle_LATAM_SODIMAC_FALABELLA_MERCADO PAGO_DDMMYYYY'\n"
            "5. Llenar archivo 'Solicitud OPS DDMMYYYY'\n"
            "6. Generar archivo plano 'OPS DDMMYYYY'\n"
            "7. Comprimir archivo 'Formato OPS DDMMYYYY'\n"
            "8. Generar historico de archivos consolidados\n"
            "9. Salir"
            "\nSelecciona una opción: ")
        if opcion == "1":
            # 01 - CARGA Y VALIDACIÓN DE COLUMNAS
            escribir("Ahora, vamos a cargar el archivo y validar que las columnas sean correctas...\n")
            df = concatenar_datos(carpeta_archivos, clave, exts, omitir)
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
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:  
                    f.write("\n-------02 - Validación de largo de campos.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                
                if os.path.exists(ruta_error_largo_campos):
                # formatar las columnas con el error detectado
                    formato_para_columnas(ruta_error_largo_campos,"Errores Largo Campos",columnas_formato=['__archivo_origen','Causal', 'campo_evento', 'descripcion'])
            # 03 - VALIDACIÓN DE CAMPOS ESPECÍFICOS (Ejemplo: columna 'tipo')
            try:
                
                escribir("Validando valores de columna 'tipo'...\n")
                validar_columna_tipo(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Ups! Se han encontrado errores en la validación de la columna 'tipo'.\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    # formatar las columnas con el error detectado
                    formato_para_columnas(ruta_columna_tipo,"Errores Columna Tipo",columnas_formato=['__archivo_origen', 'campo_evento', 'descripcion'])  
                    
                    f.write("\n-------03 - Validación de campos específicos.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")

            #04 - VALIDACIÓN DE CAMPOS VACÍOS
            try:
                escribir("Validando campos vacíos...\n")
                validar_campos_vacios(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Oh no! Se han encontrado campos vacíos que requieren atención.\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    formato_para_columnas(ruta_alertas,"Alertas Campos Vacíos",columnas_formato=['__archivo_origen', 'Causal','campo_evento', 'descripcion'])  
                    
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
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------05 - Validación de decimales en campos de valor.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                
            #06 VALIDACIÓN DE INICIO DE CAMPO NUMERO DE CUENTA
            try:
                escribir("Validando inicio de campo 'numero de la cuenta'...\n")
                validar_inicio_numero_cuenta(df, "numero de la cuenta", ("1","2"))
                
            except Exception as e:
                error_msg = str(e)
                
                escribir("La columna de numero de la cuenta no inicia por los prefijos permitidos 1(Cuenta de ahorros) o 2(Cuenta corriente).\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------06 - Validación de inicio de campo 'numero de la cuenta'.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                    
            #07 VALIDACIÓN DE ENTIDAD CUENTA
            try:
                escribir("Validando el campo 'entidad de la cuenta'...\n")
                validar_entidad_cuenta(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Se encontraron valores no permitidos en la columna 'entidad de la cuenta' (diferentes de 0013).\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------07 - errores_entidad_de_cuenta.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                    
            #08 VALIDACIÓN DE FILLER
            try:
                escribir("Validando el campo 'filler'...\n")
                validar_filler(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Se encontraron valores no permitidos en la columna 'filler' (diferentes de 0).\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------08 - errores_filler.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                    
            #09 VALIDACIÓN DE TIPO
            try:
                escribir("Validando el campo 'justificacion contable' sea coincidente con el campo 'tipo'...\n")
                validar_justificacion_contable(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Se encontraron valores no permitidos y/o no coincidentes en la columna 'justificacion contable' con el campo 'tipo' (diferentes a los establecidos).\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------09 - errores_justificacion_contable.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                    
            # VALIDACIÓN DE DUPLICADOS nivel archivo
            try:
                escribir("Validando la tapa de OPS para verificar duplicados...\n")
                validar_duplicados(df)
            except Exception as e:
                error_msg = str(e)
                escribir("Se encontraron valores duplicados -> 🚧¡REVISA TU ARCHIVO ANTES DE MONTAR EN LA RUTA! Se pueden encontrar transacciones duplicadas fuera del proceso habitual.🚧\n")
                escribir(error_msg)
                
                errores_totales += 1
                with open(
                    ruta_error_txt,
                    "a", encoding="utf-8"
                ) as f:
                    f.write("\n-------ALERTA_DUPLICADOS.--------\n")
                    f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
                # FINALIZACIÓN DEL PROCESO DE VALIDACIÓN -> valida que no tenga alertas de duplicados y que no haya errores críticos
            if errores_totales == 0:
                escribir(f"\nPerfecto! No se han encontrado errores en las validaciones(no errores críticos). Los archivos estan listos para seguir los siguientes pasos del validador\n"
                    "Recuerda revisar el archivo Log.txt y para tener la traza historica de ejecuciones.\n")
                with open(
                    log_exitoso,
                    "a", encoding="utf-8"
                ) as f:
                    f.write(f"\nLog de validación exitosa: {datetime.now()} - Perfecto! No se han encontrado errores en las validaciones. Los archivos estan listos para seguir los siguientes pasos del validador\n")
                continue
            elif errores_totales > 1:
                escribir(f"\nSe han encontrado {errores_totales} errores en las validaciones. Los archivos no estan listos para seguir los siguientes pasos del validador\n")
                break  # Salir del bucle si se encuentran errores críticos
            if os.path.exists(ruta_duplicados):
                escribir(f"\nRevisa el archivo alertas de duplicados antes de montar el archivo en la ruta de la OPS para evitar posibles duplicados fuera del proceso normal.\n")
                escribir("👀COMO SE PRESENTO UNA DUPLICIDAD Y VALIDA QUE SEA NORMAL(archivo 'ALERTAS_DUPLICADOS.xlsx')👀, DESPUES SIGUE LOS PROXIMOS PASOS:\n")
                continue
            else:
                escribir(f"\nNo se han encontrado duplicados. Los archivos estan listos para seguir los siguientes pasos del validador\n")
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
                
                escribir(f"Realizando comparación del archivo unificado que corresponde a la OPS con fecha {datetime.now().strftime('%d-%m-%Y')} VS el archivo de Historico de las OPS")
                comparar_consolidados(ruta_consolidado_historico, ruta_archivo_unificado, ruta_retorno_duplicados, "Historico OPS")
                formato_para_columnas(ruta_retorno_duplicados, "Alerta Duplicados", columnas_formato=['Número de fila','Fecha_historico','descripcion'])
                
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
            cargar_estructura(ruta_archivo_unificado, ruta_formato_ops, hoja_base, MAPEO, fila_encabezados_destino=8, fila_encabezados_origen=1)
            
            escribir(f"\nArchivo montado exitosamente:")
            escribir(f"   Ubicación: {ruta_formato_ops}")
            escribir(f"   Hoja: {hoja_base}")
            
        elif opcion == "4":
            # Consolida la información de los archivos de debitos
            escribir("Consolidando archivo de debitos...")
            df_total_debitos = obtener_debitos(carpeta_archivos, clave_debitos, exts, omitir)
            
            if df_total_debitos.empty:
                escribir("No se encontraron archivos de débitos para consolidar.")
                continue
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
            cargar_estructura(ruta_archivo_debitos, ruta_formato_debitos, hoja_base_debitos, MAPEO_DEBITOS, fila_encabezados_destino=1, fila_encabezados_origen=1)
            
            escribir(f"\nArchivo de debitos montado exitosamente:")
            escribir(f"   Ubicación: {ruta_formato_debitos}")
            escribir(f"   Hoja: {hoja_base_debitos}")
            
            
        elif opcion == "5":
            # Verificar que la plantilla de solicitud exista
            if not os.path.exists(ruta_libro_base_solicitud):
                escribir(f"✖️ La plantilla de solicitud no existe: {ruta_libro_base_solicitud}")
                continue
            
            # Copiar la plantilla (limpia) al destino con el nombre con fecha
            escribir(f"   Copiando plantilla limpia: {ruta_libro_base_solicitud}")
            escribir(f"   A: {ruta_solicitud_ops}")
            shutil.copy2(ruta_libro_base_solicitud, ruta_solicitud_ops)
            
            
            escribir("Llenando archivo 'Solicitud OPS DDMMYYYY'...")
            obtener_solicitud_ops(ruta_formato_ops, ruta_solicitud_ops, hoja_base, hoja_base_solicitud,  MAPEO_SOLICITUD, arreglo_fecha)
            
            # Extraer cuentas y descripciones para llenar la solicitud OPS
            extraer_cuentas_y_descripciones(ruta_formato_ops, ruta_solicitud_ops, hoja_base, hoja_base_solicitud, celda_descripcion="D32", celda_cuentas="D26", celda_codigo_MIR="D23", columna_inicio=10, fila_inicio=9, separador="\n", separador_codigo_MIR=" - ")
            
            escribir(f"\nArchivo de solicitud montado exitosamente:")
            escribir(f"   Ubicación: {ruta_solicitud_ops}")
            escribir(f"   Hoja: {hoja_base_solicitud}")
            
        elif opcion == "6":
            escribir("Generando archivo plano 'OPS DDMMYYYY'...")
            
            exportar_txt_limpio(ruta_formato_ops, hoja_base, "R", ruta_archivo_plano_txt)
            
            escribir(f"\nArchivo plano generado exitosamente:")
            escribir(f"   Ubicación: {ruta_archivo_plano_txt}")
        elif opcion == "7":
            escribir("Comprimiendo archivo 'Formato OPS DDMMYYYY'...")
            comprimir_excel(ruta_formato_ops, formatos)
        elif opcion == "8":
            
            if not os.path.exists(ruta_archivo_unificado):
                escribir(f"✖️ El archivo consolidado no se encuentra: {ruta_archivo_unificado}")
                continue
            
            escribir("Generando historico de archivos consolidados...")
            cargar_estructura(ruta_archivo_unificado, ruta_consolidado_historico, Hoja_base_historico, MAPEO, fila_encabezados_destino=1, fila_encabezados_origen=1)
            
            
            escribir("Depurando historico de archivos consolidados...")
            depurar_historico(ruta_consolidado_historico, "FECHA", Hoja_base_historico, dias_ventana=30)
            
        elif opcion == "9":
            escribir("Saliendo del programa 🗝️VALIDADOR🗝️. ¡Hasta luego!")
            exit()
        else:
            escribir("Opción no válida. Por favor, selecciona 1, 2, 3, 4, 5 o 6.")
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
