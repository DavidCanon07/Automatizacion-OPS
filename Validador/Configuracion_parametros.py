from pathlib import Path
from datetime import datetime

#---------------------Parametros para validación de archivos----------------------

# Definición de columnas ancla para la validación de filas válidas
columna_ancla = ['Entidad de la cuenta', 'Centro cuenta', 'numero de la cuenta', 'tipo']

# Configuración de parámetros para el pre-validador
Campos_a_validar = ['Unnamed: 0','Unnamed: 1','Entidad de la cuenta','Centro cuenta','filler','numero de la cuenta','tipo','valor ajuste','Cuenta a afectar','Justificacion contable','Cuentas contables contrapartida','Detalle del ajuste realizado','TIPO DE DOCUMENTO','NUMERO DE DOCUMENTO','DIGITO DE VERIFICACION']


# Configuración de validación de largo de campos (longitud máxima permitida)
largo_campos = {'Entidad de la cuenta': 4, 'Centro cuenta': 4, 'filler': 1, 'numero de la cuenta': 9,'tipo': 1, 'Cuenta a afectar': 20, 'Cuentas contables contrapartida': [9,12],'TIPO DE DOCUMENTO': 1, 'DIGITO DE VERIFICACION': 1}

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


#----------------Rutas y configuraciones generales----------------------
#Configuración para búsqueda de archivos
base_path = Path(r"C:\validador")
carpeta_archivos = base_path / "carpetas_OPS"
ejecucion = base_path / "Control de ejecuciones"
estructura = base_path / "Estructuras"
formatos = base_path / "Archivos OPS"
historico = base_path / "Historico"



#Configuración para búsqueda de archivos
clave = 'OPS'
omitir = ['soporte OPS', 'soportes OPS', 'nueva OPS', 'CHECK LIST OPS']
clave_debitos = 'Detalle_LATAM'
exts = {'.xlsx', '.xls', '.xlsb', '.xlsm'}

# Rutas para guardar archivos de errores y alertas
ruta_error_txt = r"C:\validador\Control de ejecuciones\errores.txt"
log_exitoso = r"C:\validador\Control de ejecuciones\log.txt"

#ruta carpeta con archivos retorno pos validación
ruta_error_largo_campos = r"C:\validador\02 - errores_largo_campos.xlsx"
ruta_columna_tipo = r"C:\validador\03 - errores_columna_tipo.xlsx"
ruta_alertas = r"C:\validador\04 - alertas_campos.xlsx"
ruta_redondeo = r"C:\validador\05 - errores_redondeo.xlsx"
ruta_inicio_campo = r"C:\validador\06 - errores_inicio_numero_cuenta.xlsx"
ruta_entidad_cuenta = r"C:\validador\07 - errores_entidad_cuenta.xlsx"
ruta_filler = r"C:\validador\08 - errores_filler.xlsx"
ruta_justificacion_contable = r"C:\validador\09 - errores_justificacion.xlsx"
ruta_duplicados = r"C:\validador\ALERTA_DUPLICADOS.xlsx"





#----------------------Parametros para la consolidación de archivos----------------------
ruta_archivo_unificado = r"C:\validador\temp_archivos_unificados.xlsx"
ruta_libro_base = r"C:\validador\Estructuras\Estructura_Formato_OPS.xlsx"
hoja_base = "Formato OPS"

#ruta carpeta con archivos retorno pos validación
ruta_formato_ops = formatos / f"FORMATO OPS {datetime.now().strftime('%d%m%Y')}.xlsx"
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

#----------------------Parametros para consolidar los debitos----------------------

ruta_archivo_debitos = r"C:\validador\temp_archivos_unificados_debitos.xlsx"
ruta_libro_base_debitos = r"C:\validador\Estructuras\Estructura_Debitos.xlsx"
hoja_base_debitos = "LATAM"

#ruta carpeta con archivos retorno pos validación
ruta_formato_debitos = formatos / f"Detalle_LATAM_SODIMAC_FALABELLA_MERCADO PAGO_{datetime.now().strftime('%d%m%Y')}.xlsx"
ruta_formato_debitos = str(ruta_formato_debitos)  # Convertir a string para openpyxl

MAPEO_DEBITOS = {
    'Número de Terminal':'Número de Terminal',
    'Codigo Autorización':'Codigo Autorización',
    'Autorización2':'Autorización2',
    'Fecha Canje':'Fecha Canje',
    'Fecha Comprobante':'Fecha Comprobante',
    'Fecha Consignación':'Fecha Consignación',
    'Franquicia':'Franquicia',
    'Transaccion':'Transaccion',
    'Bin Fuente':'Bin Fuente',
    'Tipo de Deposito':'Tipo de Deposito',
    'Tipo transacción':'Tipo transacción',
    'Naturaleza':'Naturaleza',
    'Cuenta consignación':'Cuenta consignación',
    'Codigo establecimiento':'Codigo establecimiento',
    'Número de NIT':'Número de NIT',
    'Nombre establecimient':'Nombre establecimient',
    'CODIGO MCC':'CODIGO MCC',
    'Tipo de Negocio':'Tipo de Negocio',
    'TARJETA':'TARJETA',
    'TIPO BIN':'TIPO BIN',
    'CHCCFT':'CHCCFT',
    'VALOR COMPRA':'VALOR COMPRA',
    'VALOR IVA':'VALOR IVA',
    'VALOR PROPINA':'VALOR PROPINA',
    'TOTAL= Venta + IVA + Propina':'TOTAL= Venta + IVA + Propina',
    'Valor Comisión':'Valor Comisión',
    'Rete IVA':'Rete IVA',
    'Rete Fuente':'Rete Fuente',
    'Rete ICA':'Rete ICA',
    'Valor Neto':'Valor Neto',
    'VALOR INTERCAMBIO':'VALOR INTERCAMBIO',
    'IATA':'IATA',
    'NOMBRE AGENCIA':'NOMBRE AGENCIA',
    'CIUDAD AGENCIA':'CIUDAD AGENCIA',
    'TARJETA COMPLETA':'TARJETA COMPLETA',
    'IDENTIFICADORUNICO':'IDENTIFICADORUNICO',

}

#------------------------Parametros para rellenar la solicitud de la OPS------------------------
ruta_libro_base_solicitud = r"C:\validador\Estructuras\Estructura_Solicitud_OPS.xlsx"
hoja_base_solicitud = "Solicitud_OPS"
arreglo_fecha = datetime.now().strftime("%d/%m/%Y") #Fecha actual para colocar en la solicitud de OPS

ruta_solicitud_ops = formatos / f"Solicitud OPS {datetime.now().strftime('%d%m%Y')}.xlsx"
ruta_solicitud_ops = str(ruta_solicitud_ops)  # Convertir a string para openpyxl

MAPEO_SOLICITUD = {
    'J1':'D14',
    'I1':'D15',
    'J2':'D17',
    'I2':'D18',
    
}

#------------------------Ruta para archivo plano OPS------------------------

ruta_archivo_plano_txt = formatos / f"OPS {datetime.now().strftime('%d%m%Y')}.txt"
ruta_archivo_plano_txt = str(ruta_archivo_plano_txt)  # Convertir a string para openpyxl




#------------------------Path base para copiar y pegar las estructuras base en la carpeta de estructuras------------------------

ruta_base_estructuras = r"C:\Programa_validador\estructuras_base"
formato_ops = Path(ruta_base_estructuras) / "Estructura_Formato_OPS.xlsx"
formato_debitos = Path(ruta_base_estructuras) / "Estructura_Debitos.xlsx"
formato_solicitud = Path(ruta_base_estructuras) / "Estructura_Solicitud_OPS.xlsx"
archivo_historico = Path(ruta_base_estructuras) / "historico_OPS.xlsx"


#----------------------parametros para historico_resumen----------------------

ruta_consolidado_historico = r"C:\validador\Historico\historico_OPS.xlsx"
Hoja_base_historico = "Historico OPS"

ruta_retorno_duplicados = base_path / f"ALERTA_DUPLICADOS_historico {datetime.now().strftime('%d-%m-%Y')}.xlsx"
ruta_retorno_duplicados = str(ruta_retorno_duplicados)

#-----------------------Parametros para Lectura directa desde ruta (automatización de ingesta)-----------------------

ruta_origen_cifrado = r"C:\Programa_validador"
ruta_dinamica = Path(ruta_origen_cifrado) / f"OPS {datetime.now().strftime('%d-%m-%Y')}"
ruta_dinamica = str(ruta_dinamica)

ruta_destino_desktop = r"C:\validador\carpetas_OPS"
ruta_destino_desktop = str(ruta_destino_desktop)