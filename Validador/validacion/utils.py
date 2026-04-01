import re
import time as t
from openpyxl.utils import get_column_letter
import shutil
from pathlib import Path

#-------------------------------------------------------------------------------------------------------------
#Función para convertir nombres de columnas extra a un formato más legible
def formatea_nombre_columna(nombre: str) -> str:
    
    #Si el nombre es 'Unnamed: N', devuelve 'Columna {LETRA}' (A=1).
    #Si no, devuelve el nombre original.
    #Ejemplo: 'Unnamed: 0' -> 'Columna A', 'Unnamed: 1' -> 'Columna B', etc.
    m = re.match(r"^Unnamed:\s*(\d+)$", str(nombre), flags=re.IGNORECASE)
    if m:
        idx0 = int(m.group(1))           # índice 0-based
        letra = get_column_letter(idx0 + 1)  # Excel es 1-based
        return f"Columna {letra}"
    return str(nombre)

#-------------------------------------------------------------------------------------------------------------


#Función para imprimir texto con efecto de máquina de escribir
def escribir(texto, velocidad=0.01):
    for char in texto:
        print(char, end="", flush=True)
        t.sleep(velocidad)
    print()  # Salto de línea solo al final
    


def input_con_efecto(texto, velocidad=0.01):
    """Muestra texto con efecto máquina de escribir y retorna input del usuario"""
    for char in texto:
        print(char, end="", flush=True)
        t.sleep(velocidad)
    print()  # Salto de línea
    return input()  # input sin prompt


#Función para crear carpetas si no existen
def crear_carpeta_si_no_existe(base_path, carpeta_archivos, ejecucion, formatos, estructura):
    try:
        for path in [base_path, carpeta_archivos, ejecucion, formatos, estructura]:
            if not path.exists():
                escribir(f"Creando carpeta: {path}")
                path.mkdir(parents=True, exist_ok=True)
            else:
                continue
    except Exception as e:
        escribir(f"Error al crear carpetas: {e}")
        
#Función para validar que las estructuras necesarias para la consolidación existan en la carpeta de estructuras
def validar_estructuras(ruta_libro_base, ruta_libro_base_debitos, ruta_libro_base_solicitud, 
                        formato_ops, formato_debitos, formato_solicitud):
    try:
        #convertir a Path si no lo son
        ruta_libro_base = Path(ruta_libro_base)
        ruta_libro_base_debitos = Path(ruta_libro_base_debitos)
        ruta_libro_base_solicitud = Path(ruta_libro_base_solicitud)
        
        # Verificar cada estructura
        if not ruta_libro_base.is_file():
            escribir(f"Copiando estructura base: {ruta_libro_base.name}")
            shutil.copy2(formato_ops, ruta_libro_base)
        
        if not ruta_libro_base_debitos.is_file():
            escribir(f"Copiando estructura débitos: {ruta_libro_base_debitos.name}")
            shutil.copy2(formato_debitos, ruta_libro_base_debitos)
        
        if not ruta_libro_base_solicitud.is_file():
            escribir(f"Copiando estructura solicitud: {ruta_libro_base_solicitud.name}")
            shutil.copy2(formato_solicitud, ruta_libro_base_solicitud)
            
    except Exception as e:
        escribir(f"Error al validar estructuras: {e}")