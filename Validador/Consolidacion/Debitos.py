from pathlib import Path
import pandas as pd
import time as t
from validacion.utils import escribir

#consolidar debitos de LATAM y consolidar con el resto de debitos
def obtener_debitos(carpeta_archivos,clave_debitos,exts, omitir):

    datos_debitos = []
    ruta_base = Path(carpeta_archivos)
    archivos_procesados = 0
    archivos_con_error = 0
    errores_detalle = []
    
    #Recorrer todas las subcarpetas y archivos recursivamente
    for archivo in ruta_base.rglob("*"):
        if archivo.is_file() and archivo.suffix.lower() in exts and str(clave_debitos).lower() in archivo.stem.lower() and not any(om.lower() in archivo.stem.lower() for om in omitir):
                
            carpeta_nombre = archivo.parent.name  # nombre de la carpeta
            
            escribir(f"Carpeta: {carpeta_nombre}")
            escribir(f"Archivo encontrado: {archivo.name}")
            t.sleep(0.5)
            
            extension = archivo.suffix.lower()
            try:
                if extension == ".xlsb":
                    try:
                        df = pd.read_excel(archivo, dtype=str, engine='calamine')
                        escribir("Archivo .xlsb leído con éxito usando 'calamine'")
                    except ImportError:
                        df = pd.read_excel(archivo, dtype=str, engine='pyxlsb')
                        escribir("Archivo .xlsb leído con éxito usando 'pyxlsb'")
                elif extension in [".xlsx", ".xlsm"]:
                    df = pd.read_excel(archivo, dtype=str, engine='openpyxl')
                    escribir(f"Archivo {extension} leído con éxito usando 'openpyxl'")
                else:
                    df = pd.read_excel(archivo, dtype=str)
                    escribir(f"Archivo {extension} leído con éxito usando el motor por defecto")
            except Exception as e:
                escribir(f"Error al leer el archivo {archivo.name} con extensión {extension}: {str(e)}")
                escribir("Se continuará con el siguiente archivo.")
                continue
            try:    
                #FORMATAR columnas
                if "Fecha Canje" in df.columns:
                    df["Fecha Canje"] = pd.to_datetime(df["Fecha Canje"], errors='coerce').dt.strftime("%d/%m/%Y")
                
                # Formatear Fecha Comprobante
                if "Fecha Comprobante" in df.columns:
                    df["Fecha Comprobante"] = pd.to_datetime(df["Fecha Comprobante"], errors='coerce').dt.strftime("%d/%m/%Y")
                
                # ========== FORMATEAR COLUMNAS NUMÉRICAS ==========
                # Lista de columnas numéricas a convertir
                columnas_numericas = [
                    "VALOR COMPRA", "VALOR IVA", "VALOR PROPINA", 
                    "TOTAL= Venta + IVA + Propina", "Valor Comisión", 
                    "Rete IVA", "Rete Fuente", "Rete ICA", "Valor Neto"
                ]
                
                # Convertir cada columna numérica
                for col in columnas_numericas:
                    if col in df.columns:
                        # Limpiar: eliminar comas y convertir a float
                        df[col] = df[col].astype(str).str.replace(",", "").str.replace(" ", "")
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                
                df = df.reset_index(drop=True)
                datos_debitos.append(df)
                archivos_procesados += 1
                escribir(f"filas procesadas: {len(df)}")
            except Exception as e:
                archivos_con_error += 1
                error_msg = f"Error al procesar el archivo {archivo.name}: {str(e)}"
                errores_detalle.append(error_msg)
                escribir(error_msg)
                escribir("Se continuará con el siguiente archivo.")
                continue
    
    # Mostrar resumen de errores al finalizar el proceso (ESTOS BLOQUES ESTÁN FUERA DEL for)
    if errores_detalle:
        escribir("\nResumen de errores en archivos de débitos:")
        for error in errores_detalle:
            escribir(f"- {error}")
        escribir(f"\nTotal archivos procesados: {archivos_procesados}")
        escribir(f"Total archivos con error: {archivos_con_error}") 
            
    if not datos_debitos:
        escribir("No se encontraron archivos de débitos válidos para procesar.")
        return pd.DataFrame()  # Retornar DataFrame vacío si no hay datos
    
    # Mostrar resumen
    escribir(f"\n✔️ Total archivos de debitos encontrados: {len(datos_debitos)}")

    #concatenar los datos de debitos
    df_total =pd.concat(datos_debitos, ignore_index=True)
    escribir(f"Total filas consolidadas en debitos: {len(df_total)}")
    return df_total



def exportar_debitos(df_total_debitos, ruta, sheet_name):
    
    try:
        with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
            df_total_debitos.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n✔️ Archivo de debitos exportado correctamente: {ruta}")
    except Exception as e:
        print(f"\n✖️ Error al exportar archivo de debitos: {e}")
        raise