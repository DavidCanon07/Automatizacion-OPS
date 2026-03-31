import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font

def cargar_estructura(ruta_archivo_unificado, ruta_libro_base, hoja_base, MAPEO):
    
    # Verificar si el archivo destino existe, si no, crearlo
    if not os.path.exists(ruta_libro_base):
        print(f"  El archivo destino no existe. Creando: {ruta_libro_base}")
        # Crear un libro nuevo con la hoja base
        wb_new = Workbook()
        ws_new = wb_new.active
        ws_new.title = hoja_base
        wb_new.save(ruta_libro_base)
        wb_new.close()
        
    # Mapear el libro base
    wb_original = load_workbook(ruta_archivo_unificado, read_only=True, data_only=True)
    ws_original = wb_original.active

    wb_base = load_workbook(ruta_libro_base)
    ws_base = wb_base[hoja_base]
    
    # índice de columnas en el libro origen(unificado)
    encontrar_origen = [str(c or "").strip() for c in next(ws_original.iter_rows(values_only=True))]
    # índice de columnas en el libro base
    encontrar_base = [str(c or "").strip() for c in next(ws_base.iter_rows(values_only=True, min_row=8))]
    
    print(f" Encabezados ORIGEN: {encontrar_origen[:5]}...")
    print(f" Encabezados DESTINO: {encontrar_base[:5]}...")
    
    # encontrar la primera fila vacia en la base sin el cuadro de suma 
    fila_libre = ws_base.min_row + 8
    
    copiadas = 0
    filas_procesadas = 0
    
    # Copiar los datos del libro origen al libro base
    for row in ws_original.iter_rows(min_row=2, values_only=True):
        filas_procesadas += 1
        datos_copiados = False
        
        for col_o, col_d in MAPEO.items():
            if col_o in encontrar_origen and col_d in encontrar_base:
                idx_origen = encontrar_origen.index(col_o)
                idx_destino = encontrar_base.index(col_d)
                
                valor = row[idx_origen] if idx_origen < len(row) else None
                
                if valor is not None and valor != "":
                    ws_base.cell(row=fila_libre, column=idx_destino + 1, value=valor)
                    datos_copiados = True
        
        if datos_copiados:
            copiadas += 1
            fila_libre += 1
    
    # Guardar el libro base con los datos copiados
    wb_original.close()
    wb_base.save(ruta_libro_base)
    
    print(f"✔️ Se procesaron {filas_procesadas} filas desde origen")
    print(f"✔️ Se copiaron {copiadas} filas al libro base")
    print(f"✔️ Archivo guardado en: {ruta_libro_base}")