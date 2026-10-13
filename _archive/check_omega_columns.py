import pandas as pd

df_omega = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\LYA_OMEGA_1303_komplett.xlsx', sheet_name='BOQNT', engine='openpyxl')

print(f'OMEGA har {len(df_omega.columns)} kolonner:')
for i, col in enumerate(df_omega.columns):
    print(f'{i+1}. {col}')

# Sjekk spesifikt for posten 1303.01.06.281.1
test_post = '1303.01.06.281.1'
test_row = df_omega[df_omega['Navn (Prosesskode)'] == test_post]

if not test_row.empty:
    print(f'\n\nPosten {test_post} finnes i OMEGA:')
    for col in test_row.columns:
        val = test_row[col].iloc[0]
        print(f'  {col}: {val}')
else:
    print(f'\nPosten {test_post} finnes IKKE i OMEGA')
