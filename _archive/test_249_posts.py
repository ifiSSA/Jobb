import os
import pandas as pd

df = pd.read_excel(r"C:\Users\ifi\OneDrive - Scan Survey\VEK_OMEGA_1302_komplett.xlsx", sheet_name="BOQNT")

# Sjekk betong-filteret
betong = df[(df['Tittel'].astype(str).str.contains('BETONG', case=False, na=False)) & 
            (df['Enhet'].astype(str).str.lower() == 'm³')]

posts_249 = betong[betong['Navn (Prosesskode)'].astype(str).str.contains('1302.03.05.249.1', na=False)]
print(f"Funnet {len(posts_249)} poster med 1302.03.05.249.1 i betong-filteret")

# Sjekk om de er i resultatfilen
df_res = pd.read_excel(os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "excel"), "resultat_mengdeuttak_med_OMEGA_20260129_155253.xlsx"))

# Finn rader som har 1302.03.05.249.1.1 eller 1302.03.05.249.1.2 i Mengdelistepost_Betong
rader_med_249 = df_res[
    df_res['Mengdelistepost_Betong'].astype(str).str.contains('1302.03.05.249.1', na=False)
]

print(f"\nFunnet {len(rader_med_249)} rader i resultatfilen med 1302.03.05.249.1 i Mengdelistepost_Betong")
if len(rader_med_249) > 0:
    print(rader_med_249[['Produksjonsenhet', 'Mengdelistepost_Hovedkode', 'Mengdelistepost_Betong', 'Total_Volum', 'Total_Vekt']].to_string())
