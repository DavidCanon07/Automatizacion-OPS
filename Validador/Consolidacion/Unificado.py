import pandas as pd
import re
import openpyxl
def limpiar_numero_latino(valor):
    """
    Convierte un número en formato latino (1.234,56) a float (1234.56)
    Maneja múltiples formatos:
    - "1.234,56" → 1234.56
    - "1234,56" → 1234.56
    - "1,234.56" (inglés) → 1234.56
    - "1234.56" → 1234.56
    - "1.234" (con punto, puede ser miles o decimal)
    """
    if pd.isna(valor) or valor == "":
        return None
    
    # Si ya es número, devolverlo
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Convertir a string y limpiar espacios
    valor_str = str(valor).strip()
    if not valor_str:
        return None
    
    # Detectar formato según patrones
    # Caso 1: Tiene punto y coma (formato latino claro: 1.234,56)
    if '.' in valor_str and ',' in valor_str:
        # Eliminar puntos de miles, reemplazar coma decimal por punto
        valor_str = valor_str.replace('.', '').replace(',', '.')
    
    # Caso 2: Solo tiene comas (formato latino sin miles: 1234,56)
    elif ',' in valor_str and '.' not in valor_str:
        # Reemplazar coma decimal por punto
        valor_str = valor_str.replace(',', '.')
    
    # Caso 3: Solo tiene puntos
    elif '.' in valor_str and ',' not in valor_str:
        # Intentar determinar si el punto es decimal o de miles
        partes = valor_str.split('.')
        
        if len(partes) == 2:
            # Un solo punto: puede ser decimal o miles
            if len(partes[1]) <= 2 and partes[1].isdigit():
                # El segmento después del punto es pequeño (1-2 dígitos) → es decimal
                pass  # Mantener como está
            else:
                # El segmento después del punto es largo → probablemente son miles
                valor_str = valor_str.replace('.', '')
        else:
            # Múltiples puntos → son miles, eliminar todos
            valor_str = valor_str.replace('.', '')
    
    # Intentar convertir a float
    try:
        resultado = float(valor_str)
        return resultado
    except ValueError:
        # Si falla, intentar extraer números con regex
        numeros = re.findall(r'[\d,\.]+', valor_str)
        if numeros:
            try:
                # Probar con el primer grupo de números
                return limpiar_numero_latino(numeros[0])
            except:
                pass
        print(f" No se pudo convertir: '{valor}'")
        return valor


def exportar_excel(df, ruta, sheet_name):
    columnas_exportar = [
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
    
    # Filtrar solo las columnas que existen
    columnas_existentes = [col for col in columnas_exportar if col in df.columns]
    
    # Crear copia
    df_exportar = df[columnas_existentes].copy()
    
    # CONVERTIR LA COLUMNA DE VALORES A FLOAT
    if 'valor ajuste' in df_exportar.columns:
        print("\n PROCESANDO COLUMNA 'valor ajuste'...")
        
        # Mostrar ejemplos antes de convertir
        print("   Ejemplos antes de conversión:")
        for i, val in enumerate(df_exportar['valor ajuste'].head(5)):
            print(f"      {i+1}: '{val}' (tipo: {type(val).__name__})")
        
        # Aplicar conversión
        df_exportar['valor ajuste'] = df_exportar['valor ajuste'].apply(limpiar_numero_latino)
        
        # Mostrar ejemplos después de convertir
        print("\n   Ejemplos después de conversión:")
        for i, val in enumerate(df_exportar['valor ajuste'].head(5)):
            print(f"      {i+1}: {val} (tipo: {type(val).__name__})")
        
        # Estadísticas
        num_numericos = df_exportar['valor ajuste'].apply(lambda x: isinstance(x, (int, float))).sum()
        num_texto = len(df_exportar) - num_numericos
        print(f"\n  Estadísticas de conversión:")
        print(f"      ✔️ Valores numéricos: {num_numericos}")
        print(f"      ✖️ Valores no convertidos: {num_texto}")
    
    # Agregar columna de índice
    df_exportar.insert(0, 'N°', range(1, len(df_exportar) + 1))
    
    # Agregar columna de concatenación PARA COMPARAR con el historico
    df_exportar['Concatenado'] = (
        df_exportar['Proceso'].astype(str) + 
        df_exportar['Entidad de la cuenta'].astype(str) + 
        df_exportar['Centro cuenta'].astype(str) + 
        df_exportar['filler'].astype(str) + 
        df_exportar['numero de la cuenta'].astype(str) + 
        df_exportar['tipo'].astype(str) + 
        df_exportar['valor ajuste'].apply(lambda x: f"{x:.2f}".replace('.', ',') if pd.notna(x) else '0,00') + 
        df_exportar['Detalle del ajuste realizado'].astype(str)
)
    
    # Exportar
    try:
        with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
            df_exportar.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n✔️ Archivo exportado correctamente: {ruta}")
    except Exception as e:
        print(f"\n✖️ Error en exportación: {e}")
        raise