import pandas as pd
from pathlib import Path
import os
from Configuracion_parametros import *
from openpyxl import load_workbook


print(ruta_formato_ops)

wb_destino = load_workbook(ruta_formato_ops, read_only=False, data_only=False, )
ws_destino = wb_destino[hoja_base]

wb_destino.save(r"C:\Prevalidador\Formato OPS.xlsx")
wb_destino.close()

