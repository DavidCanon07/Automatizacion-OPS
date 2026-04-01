import zipfile
import os

def comprimir_excel(archivo_excel, carpeta_destino):
    """
    Comprime un archivo Excel en un archivo .zip
    
    Args:
        archivo_excel: Ruta del archivo Excel a comprimir
        carpeta_destino: Carpeta donde guardar el zip (opcional)
    
    Returns:
        Ruta del archivo zip creado
    """
    # Definir nombre del zip
    nombre_base = os.path.splitext(os.path.basename(archivo_excel))[0]
    if carpeta_destino:
        zip_path = os.path.join(carpeta_destino, f"{nombre_base}.zip")
    else:
        zip_path = f"{nombre_base}.zip"
    
    # Crear archivo zip
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(archivo_excel, os.path.basename(archivo_excel))
    
    print(f"✔️ Archivo comprimido: {zip_path}")
    return zip_path
