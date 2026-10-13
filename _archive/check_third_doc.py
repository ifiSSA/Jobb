import os
import pandas as pd

_IFI = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "excel")
df = pd.read_excel(os.path.join(_IFI, 'resultat_mengdeuttak_med_OMEGA.xlsx'), engine='openpyxl')
print('First 25 rows:')
print(df.head(25).to_string())
print('\nStatistics:')
print(f'Total rows: {len(df)}')
omega_col = "OMEGA_Tittel" if "OMEGA_Tittel" in df.columns else df.columns[4]
print(f'Rows with data: {df[omega_col].notna().sum()}')
print(f'Rows without data: {df[omega_col].isna().sum()}')
print('\nKolonner:')
print(df.columns.tolist())
