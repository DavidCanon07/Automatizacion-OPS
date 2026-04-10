import pandas as pd

"""
df = pd.read_excel(r"C:\validador\Historico\historico_OPS.xlsx", sheet_name="Historico OPS")


# DD/MM/YYYY (ej: 31/03/2026)
if "FECHA" in df.columns:
                    df["FECHA"] = pd.to_datetime(df["FECHA"], format='%d/%m/%Y').dt.date

df.info()

df.to_excel(r"C:\validador\Historico\historico_OPS_copia.xlsx", sheet_name="Historico OPS", index=False)


"""

import win32com
print(win32com.__gen_path__)