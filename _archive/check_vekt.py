import pandas as pd

# Last inn df_armering
df_armering = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\Armering_LP09.xlsx', sheet_name='Armering', engine='openpyxl')
print('df_armering Total_Vekt sum:', df_armering['PGF_RIB_Info.Vekt'].sum())
print('df_armering rows:', len(df_armering))

# Last inn output
df_out = pd.read_excel('resultat_mengdeuttak_med_OMEGA.xlsx')
print('\nOutputdokument Total_Vekt sum:', df_out['Total_Vekt'].sum())
print('Output rows:', len(df_out))

# Sjekk vekt-kolonner
print('\nVekt-kolonner sum:')
for col in ['Vekt_0_008', 'Vekt_0_01', 'Vekt_0_012', 'Vekt_0_016', 'Vekt_0_02']:
    print(f'{col}: {df_out[col].sum()}')

print('\nSum av alle Vekt_*:', df_out[['Vekt_0_008', 'Vekt_0_01', 'Vekt_0_012', 'Vekt_0_016', 'Vekt_0_02']].sum(axis=1).sum())
