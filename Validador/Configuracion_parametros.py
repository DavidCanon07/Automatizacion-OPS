from pathlib import Path
from datetime import datetime
import time as t

# Definición de columnas ancla para la validación de filas válidas
columna_ancla = ['Entidad de la cuenta', 'Centro cuenta', 'numero de la cuenta', 'tipo']

# Configuración de parámetros para el pre-validador
Campos_a_validar = ['Unnamed: 0','Unnamed: 1','Entidad de la cuenta','Centro cuenta','filler','numero de la cuenta','tipo','valor ajuste','Cuenta a afectar','Justificacion contable','Cuentas contables contrapartida','Detalle del ajuste realizado','TIPO DE DOCUMENTO','NUMERO DE DOCUMENTO','DIGITO DE VERIFICACION']


# Configuración de validación de largo de campos (longitud máxima permitida)
largo_campos = {'Entidad de la cuenta': 4, 'Centro cuenta': 4, 'filler': 1, 'numero de la cuenta': 9,'tipo': 1, 'Cuenta a afectar': 20, 'Cuentas contables contrapartida': (9,12),'TIPO DE DOCUMENTO': 1, 'DIGITO DE VERIFICACION': 1}

#Conceptos atados a la justificación y tipo de ajuste
justificacion_contable = {
    "N": [
        "CARGO AJUSTE POR NUEVO MODELO DE RECUPERACIONES",
        "CARGO DEPOSITOS ELECTRONICOS",
        "CARGO DEVOLUCION QR",
        "CARGO NO APLICADO POR TX INTERNACIONAL CUENTA 400",
        "CARGO POR DOBLE ABONO A LA TARJETA",
        "CARGO POR TRANSFERENCIA ERRADA",
        "CARGO REEMBOLSOS",
        "CARGO RELIQUIDACION COMISIONES",
        "CARGO REVERSO PAGO",
        "CARGO SALDO A FAVOR",
        "CARGO SALDO A FAVOR ADELANTO DE NOMINA",
        "CARGOS",
        "RECUPERACION TRXS PENDIENTE DE COBRO A CLIENTES, APLICATIVO CONCISO"
    ],
    "P": [
        "ABONO AJUSTE POR NUEVO MODELO DE RECUPERACIONES",
        "ABONO DEPOSITOS ELECTRONICOS",
        "ABONO DEVOLUCION QR",
        "ABONO POR REVERSO PAGO NACIONAL AUTORIZACION",
        "ABONO RECLAMO ATM",
        "ABONO REEMBOLSOS",
        "ABONO RELIQUIDACION COMISIONES",
        "ABONO SALDO A FAVOR",
        "ABONO REVERSO PAGO",
        "ABONO SALDO A FAVOR ADELANTO DE NOMINA",
        "ABONOS",
        "ABONOS RECLAMACION",
        "DEVOLUCION INTERN VISA",
        "REINTEGRO TRANSFERENCIA P2P SIN COMPENSAR"
    ]
}

#Configuración para búsqueda de archivos
base_path = Path(r"C:\validador")
carpeta_archivos = base_path / "carpetas_OPS"
ejecucion = base_path / "Control de ejecuciones"
formatos = base_path / "Archivos OPS"

#Función para crear carpetas si no existen
def crear_carpeta_si_no_existe(base_path, carpeta_archivos, ejecucion, formatos):
    try:
        for path in [base_path, carpeta_archivos, ejecucion, formatos]:
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
ruta_error_txt = r"C:\validador\Control de ejecuciones\errores.txt"
ruta_error_largo_campos = r"C:\validador\02 - errores_largo_campos.xlsx"
ruta_columna_tipo = r"C:\validador\03 - errores_columna_tipo.xlsx"
ruta_alertas = r"C:\validador\04 - alertas_campos.xlsx"
ruta_redondeo = r"C:\validador\05 - errores_redondeo.xlsx"
log_exitoso = r"C:\validador\Control de ejecuciones\log.txt"
ruta_inicio_campo = r"C:\validador\06 - errores_inicio_numero_cuenta.xlsx"
ruta_entidad_cuenta = r"C:\validador\07 - errores_entidad_cuenta.xlsx"
ruta_filler = r"C:\validador\08 - errores_filler.xlsx"
ruta_justificacion_contable = r"C:\validador\09 - errores_justificacion.xlsx"
ruta_duplicados = r"C:\validador\ALERTA_DUPLICADOS.xlsx"


# Función para imprimir texto con efecto de máquina de escribir
def escribir(texto, velocidad=0.01):
    for char in texto:
        print(char, end="", flush=True)
        t.sleep(velocidad)
    print()  # Salto de línea solo al final


#----------------------Parametros para la consolidación de archivos----------------------
ruta_archivo_unificado = r"C:\validador\archivos_unificados.xlsx"
ruta_libro_base = r"C:\validador\Estructura OPS manipular.xlsx"
hoja_base = "Formato OPS"

#ruta carpeta con archivos retorno pos validación
ruta_formato_ops = formatos / f"Formato OPS {datetime.now().strftime('%d%m%Y')}.xlsx"
ruta_formato_ops = str(ruta_formato_ops)  # Convertir a string para openpyxl

columnas_obligatorias = [
    "N°",
    "FECHA",
    "Proceso",
    "Entidad de la cuenta",
    "Centro cuenta",
    "filler",
    "numero de la cuenta",
    "tipo",
    "valor ajuste",
    "Justificacion contable",
    "Cuentas contables contrapartida",
    'Detalle del ajuste realizado',
    'TIPO DE DOCUMENTO',
    'NUMERO DE DOCUMENTO',
    'DIGITO DE VERIFICACION'
]


MAPEO = {
    "N°": "N°",
    "FECHA": "FECHA",
    "Proceso": "Proceso",
    "Entidad de la cuenta": "Banco",      
    "Centro cuenta": "Centro",                    
    "filler": "Filler",                                  
    "numero de la cuenta": "CUENTA",        
    "tipo": "Tipo",                                      
    "valor ajuste": "Valor",                      
    "Justificacion contable": "Justificación",  
    "Cuentas contables contrapartida": "Cuenta Contrapartida",  
    "Detalle del ajuste realizado": "Detalle del ajuste realizado",        
    "TIPO DE DOCUMENTO": "TD",            
    "NUMERO DE DOCUMENTO": "DOCUMENTO",        
    "DIGITO DE VERIFICACION": "DV"   
}
