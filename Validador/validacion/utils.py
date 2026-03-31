import re
from openpyxl.utils import get_column_letter

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
