from pathlib import Path

# Configuración de parámetros para el pre-validador
Campos_a_validar = ['Unnamed: 0','Unnamed: 1','Entidad de la cuenta','Centro cuenta','filler','numero de la cuenta','tipo','valor ajuste','Cuenta a afectar','Justificacion contable','Cuentas contables contrapartida','Detalle del ajuste realizado','TIPO DE DOCUMENTO','NUMERO DE DOCUMENTO','DIGITO DE VERIFICACION']

# Configuración de validación de largo de campos (longitud máxima permitida)
largo_campos = {'Entidad de la cuenta': 4, 'Centro cuenta': 4, 'filler': 1, 'numero de la cuenta': 9,'tipo': 1, 'Cuenta a afectar': 20, 'Cuentas contables contrapartida': (9,12),'TIPO DE DOCUMENTO': 1, 'DIGITO DE VERIFICACION': 1}

#Configuración para búsqueda de archivos
base_path = Path(r"C:\Pre-validador")
carpeta_archivos = base_path / "archivos"

def crear_carpeta_si_no_existe(base_path, carpeta_archivos):
    try:
        for path in [base_path, carpeta_archivos]:
            if not path.exists():
                print(f"Creando carpeta: {path}")
                path.mkdir(parents=True, exist_ok=True)
            else:
                print(f"La carpeta '{path}' ya existe. \nContinuando con el proceso de validación...")
    except Exception as e:
        print(f"Error al crear carpetas: {e}")


carpeta = Path(r'C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\archivos')
clave = 'OPS'
exts = {'.xlsx', '.xls'}

# Rutas para guardar archivos de errores y alertas
ruta_error_txt = r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\errores.txt"
ruta_error_largo_campos = r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\02 - errores_largo_campos.xlsx"
ruta_columna_tipo = r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\03 - errores_columna_tipo.xlsx"
ruta_alertas = r"C:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\automatizacion OPS\Prevalidador\04 - alertas_campos.xlsx"



