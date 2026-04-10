from pathlib import Path
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
        "Cargo depositos electronicos",
        "Cargo devolución QR",
        "Cargo no aplicado por tx internacional CUENTA 400",
        "CARGO POR DOBLE ABONO A LA TARJETA",
        "CARGO POR TRANSFERENCIA ERRADA",
        "Cargo reembolsos",
        "Cargo reliquidación comisiones",
        "CARGO REVERSO PAGO",
        "CARGO SALDO A FAVOR",
        "CARGO SALDO A FAVOR ADELANTO DE NOMINA",
        "CARGOS",
        "RECUPERACION TRXS PENDIENTE DE COBRO A CLIENTES",
        "APLICATIVO CONCISO"
    ],
    "P": [
        "ABONO AJUSTE POR NUEVO MODELO DE RECUPERACIONES",
        "Abono depositos electronicos",
        "Abono devolución QR",
        "ABONO POR REVERSO PAGO NACIONAL AUTORIZACION",
        "ABONO RECLAMO ATM",
        "Abono reembolsos",
        "Abono reliquidación comisiones",
        "ABONO SALDO A FAVOR",
        "ABONO REVERSO PAGO",
        "ABONO SALDO A FAVOR ADELANTO DE NOMINA",
        "ABONOS",
        "Abonos Reclamacion",
        "DEVOLUCION INTERN VISA",
        "REINTEGRO TRANSFERENCIA P2P SIN COMPENSAR"
    ]
}

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
ruta_redondeo = r"C:\Prevalidador\05 - errores_redondeo.xlsx"
log_exitoso = r"C:\Prevalidador\log.txt"
ruta_inicio_campo = r"C:\Prevalidador\06 - errores_inicio_numero_cuenta.xlsx"
ruta_entidad_cuenta = r"C:\Prevalidador\07 - errores_entidad_cuenta.xlsx"
ruta_filler = r"C:\Prevalidador\08 - errores_filler.xlsx"
ruta_justificacion_contable = r"C:\Prevalidador\09 - errores_justificacion.xlsx"
ruta_duplicados = r"C:\Prevalidador\ALERTA_DUPLICADOS.xlsx"

# Función para imprimir texto con efecto de máquina de escribir
def escribir(texto, velocidad=0.01):
    for char in texto:
        print(char, end="", flush=True)
        t.sleep(velocidad)
    print()  # Salto de línea solo al final



