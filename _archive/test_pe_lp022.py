import pandas as pd

df = pd.read_excel(r"C:\Users\ifi\OneDrive - Scan Survey\VEK_OMEGA_1302_komplett.xlsx", sheet_name="BOQNT")

# Filtrer betong-poster (samme som i Mengdeuttak_Armering.py)
betong = df[(df['Tittel'].astype(str).str.contains('BETONG', case=False, na=False)) & 
            (df['Enhet'].astype(str).str.lower() == 'm³')]

# Sjekk om 1302.03.06.241.3.21 er i listen
post = betong[betong['Navn (Prosesskode)'] == '1302.03.06.241.3.21']
print(f"Betong-post 1302.03.06.241.3.21 funnet: {len(post)} ganger")
if len(post) > 0:
    print(post[['Navn (Prosesskode)', 'Tittel', 'Enhet']])

# Sjekk armering
armering = df[df["Tittel"].astype(str).str.contains('Armering', case=False, na=False)]
arm_post = armering[armering['Navn (Prosesskode)'] == '1302.03.06.241.3.20']
print(f"\nArmering-post 1302.03.06.241.3.20 funnet: {len(arm_post)} ganger")
if len(arm_post) > 0:
    print(arm_post[['Navn (Prosesskode)', 'Tittel', 'Enhet']])

# Sjekk alle poster som starter med 1302.03.06.241.3
alle = df[df['Navn (Prosesskode)'].astype(str).str.startswith('1302.03.06.241.3')]
print(f"\nAlle poster som starter med 1302.03.06.241.3: {len(alle)}")
