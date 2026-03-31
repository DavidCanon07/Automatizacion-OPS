from pathlib import Path
import pandas as pd
import time as t
from Configuracion_parametros import escribir



#consolidar debitos de LATAM y consolidar con el resto de debitos
def obtener_debitos(carpeta_archivos,clave_debitos,exts):

    datos_debitos = []
    ruta_base = Path(carpeta_archivos)
    
    try:
        #Recorrer todas las subcarpetas y archivos recursivamente
        for archivo in ruta_base.rglob("*"):
            if archivo.is_file() and archivo.suffix.lower() in exts and str(clave_debitos).lower() in archivo.stem.lower():
                
                carpeta_nombre = archivo.parent.name  # nombre de la carpeta
                
                escribir(f"Carpeta: {carpeta_nombre}")
                escribir(f"Archivo encontrado: {archivo.name}")
                t.sleep(0.5)
                
                df = pd.read_excel(archivo, dtype=str)
                
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
                
        if not datos_debitos:
            raise Exception("No hay archivos de debitos válidos para procesar".upper())
        
        # Mostrar resumen
        escribir(f"\n✔️ Total archivos de debitos encontrados: {len(datos_debitos)}")

        #concatenar los datos de debitos
        df_total =pd.concat(datos_debitos, ignore_index=True)
        return df_total
    
    except Exception as e:
        raise


def exportar_debitos(df_total_debitos, ruta, sheet_name):
    
    try:
        with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
            df_total_debitos.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n✔️ Archivo de debitos exportado correctamente: {ruta}")
    except Exception as e:
        print(f"\n✖️ Error al exportar archivo de debitos: {e}")
        raise