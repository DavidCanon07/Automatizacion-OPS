from pathlib import Path


Campos_a_validar = ["Unnamed: 0","Unnamed: 1","Entidad de la cuenta","Centro origen","Filler","Numero de la cuenta","Valor Ajuste","Cuenta","Tipo de operación","justificación contable","Cuenta contrapartida","Comentario","TIPO DE DOCUMENTO","NUMERO DE DOCUMENTO"," DIGITO DE VERIFICACIÓN"]

largo_campos = {"Entidad de la cuenta": 4, "Centro origen": 4, "Filler": 1, "Numero de la cuenta": 9, "Cuenta": 20, "Tipo de operación": 1, "Cuenta contrapartida": 12, "cuenta contrapartida": 9, "TIPO DE DOCUMENTO": 1, " DIGITO DE VERIFICACIÓN": 1}


carpeta = Path("C:\\Users\\DAVID\\OneDrive\\1. CURSOS\\Proyectos\\automatizacion OPS\\Prevalidador\\archivos")
clave = "OPS"
exts = {".xlsx", ".xls"}




