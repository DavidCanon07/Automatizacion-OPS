from pathlib import Path
import time as t

# Configuración de parámetros para el pre-validador
Campos_a_validar = ['Unnamed: 0','Unnamed: 1','Entidad de la cuenta','Centro cuenta','filler','numero de la cuenta','tipo','valor ajuste','Cuenta a afectar','Justificacion contable','Cuentas contables contrapartida','Detalle del ajuste realizado','TIPO DE DOCUMENTO','NUMERO DE DOCUMENTO','DIGITO DE VERIFICACION']

# Configuración de validación de largo de campos (longitud máxima permitida)
largo_campos = {'Entidad de la cuenta': 4, 'Centro cuenta': 4, 'filler': 1, 'numero de la cuenta': 9,'tipo': 1, 'Cuenta a afectar': 20, 'Cuentas contables contrapartida': (9,12),'TIPO DE DOCUMENTO': 1, 'DIGITO DE VERIFICACION': 1}

#Configuración para búsqueda de archivos
base_path = Path(r"C:\Prevalidador")
carpeta_archivos = base_path / "archivos"

def crear_carpeta_si_no_existe(base_path, carpeta_archivos):
    try:
        for path in [base_path, carpeta_archivos]:
            if not path.exists():
                escribir(f"Creando carpeta: {path}")
                path.mkdir(parents=True, exist_ok=True)
            else:
                continue
    except Exception as e:
        escribir(f"Error al crear carpetas: {e}")

#Configuración para búsqueda de archivos
clave = 'OPS'
exts = {'.xlsx', '.xls'}

# Rutas para guardar archivos de errores y alertas
ruta_error_txt = r"C:\Prevalidador\errores.txt"
ruta_error_largo_campos = r"C:\Prevalidador\02 - errores_largo_campos.xlsx"
ruta_columna_tipo = r"C:\Prevalidador\03 - errores_columna_tipo.xlsx"
ruta_alertas = r"C:\Prevalidador\04 - alertas_campos.xlsx"
log_exitoso = r"C:\Prevalidador\log.txt"

# Función para imprimir texto con efecto de máquina de escribir
def escribir(texto, velocidad=0.02):
    for char in texto:
        print(char, end="", flush=True)
        t.sleep(velocidad)
    print()  # Salto de línea solo al final



