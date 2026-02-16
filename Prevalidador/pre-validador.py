from Ingesta import concatenar_datos
from datetime import datetime



try: 
    
    df = concatenar_datos()
    print("Validación de campos iniciada...")
    df = df.fillna("-")
    print(df)
    df = df.drop(columns=["Unnamed: 0","Unnamed: 1"], errors="ignore")
    print(df)
    df.to_excel("C:\\Users\\dacanonm\\OneDrive - Indra\\Documentos\\GitHub\\automatizacion OPS\\Prevalidador\\archivos\\datos_concatenados.xlsx", index=False)
    
except Exception as e:
    error_msg = str(e)
    print(error_msg)
    
    #guardar el error en un archivo de texto
    with open("C:\\Users\\dacanonm\\OneDrive - Indra\\Documentos\\GitHub\\automatizacion OPS\\Prevalidador\\archivos\\errores.txt", "a", encoding="utf-8") as f:
        f.write(f"\n-------01 - Validación de columnas--------\n")
        f.write(f"\nLog de error: {datetime.now()} - {error_msg}\n")
    print("El error ha sido registrado en errores.txt")
    exit()  # Salir con código de error para indicar que hubo un problema

