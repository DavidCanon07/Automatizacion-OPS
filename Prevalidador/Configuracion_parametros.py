from pathlib import Path


Campos_a_validar = ["Unnamed: 0","Unnamed: 1","Entidad de la cuenta","Centro cuenta","filler","Numero de la cuenta","tipo","Valor Ajuste","Cuenta a afectar","justificacion contable","Cuentas contables contrapartida","Detalle del ajuste realizado","TIPO DE DOCUMENTO","NUMERO DE DOCUMENTO"," DIGITO DE VERIFICACION"]

largo_campos = {"Entidad de la cuenta": 4, "Centro cuenta": 4, "filler": 1, "Numero de la cuenta": 9,"tipo": 1, "Cuenta a afectar": 20, "Cuentas contables contrapartida": 12, "Cuentas contables contrapartida": 9, "TIPO DE DOCUMENTO": 1, " DIGITO DE VERIFICACION": 1}


carpeta = Path("C:\\Users\\dacanonm\\OneDrive - Indra\\Documentos\\GitHub\\automatizacion OPS\\Prevalidador\\archivos")
clave = "OPS"
exts = {".xlsx", ".xls"}




