import pandas as pd

# Last inn output
df_out = pd.read_excel('resultat_mengdeuttak_med_OMEGA.xlsx')

# Sjekk 1: Sjekk antal rader per Mengdelistepost_Armering
print("Armering-poster (unike):")
arm_posts = df_out["Mengdelistepost_Armering"].dropna().unique()
print(f"Antal: {len(arm_posts)}")
print(f"Dupliserte: {len(arm_posts) != df_out['Mengdelistepost_Armering'].dropna().nunique()}")

# Sjekk 2: Sjekk antal rader per Mengdelistepost_Betong
print("\nBetong-poster (unike):")
bet_posts = df_out["Mengdelistepost_Betong"].dropna().unique()
print(f"Antal: {len(bet_posts)}")

# Sjekk 3: Sjekk om samme Armering-post finnes på fleire rader
print("\nDupliserte Armering-poster:")
arm_counts = df_out["Mengdelistepost_Armering"].value_counts()
duplicates = arm_counts[arm_counts > 1]
if len(duplicates) > 0:
    print(duplicates)
else:
    print("Ingen dupliserte Armering-poster")

print("\nDupliserte Betong-poster:")
bet_counts = df_out["Mengdelistepost_Betong"].value_counts()
duplicates = bet_counts[bet_counts > 1]
if len(duplicates) > 0:
    print(duplicates)
else:
    print("Ingen dupliserte Betong-poster")

# Sjekk 4: Sjekk ein spesifikk Armering-post
print("\nEksempel - 1303.01.05.216.1.2:")
eksempel = df_out[df_out["Mengdelistepost_Armering"] == "1303.01.05.216.1.2"]
print(eksempel[["Produksjonsenhet", "Mengdelistepost_Hovedkode", "Kontraktsmengde_Armering", "Rapportert_Mengde_Kumulativ_Armering"]].to_string())
