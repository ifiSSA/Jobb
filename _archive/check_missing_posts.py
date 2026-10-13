import pandas as pd

# Les filene
df_RIB = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\RIB_LP09.xlsx', sheet_name='RIB', engine='openpyxl')
df_armering = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\Armering_LP09.xlsx', sheet_name='Armering', engine='openpyxl')
df_omega = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\LYA_OMEGA_1303_komplett.xlsx', sheet_name='BOQNT', engine='openpyxl')

# Få Mengdelistepost-kolonnene fra RIB
mengle_col = 'FOB_Info.FOB_Mengdelistepost'

# Parse RIB mengde-strenger og få alle unike poster
rib_mengde_flat = set()
rib_posts_to_rows = {}  # Mapping fra post til rad-indekser
for idx, val in enumerate(df_RIB[mengle_col].dropna()):
    for item in str(val).split(';'):
        item = item.strip()
        if item:
            rib_mengde_flat.add(item)
            if item not in rib_posts_to_rows:
                rib_posts_to_rows[item] = []
            rib_posts_to_rows[item].append(idx)

print(f'Totalt unike posts fra RIB: {len(rib_mengde_flat)}')

# OMEGA posts
omega_codes = set(df_omega['Navn (Prosesskode)'].astype(str).str.strip().unique())
print(f'Totalt posts i OMEGA: {len(omega_codes)}')

# Finn RIB-posts som IKKE finnes i OMEGA
missing_in_omega = rib_mengde_flat - omega_codes
print(f'\nPosts fra RIB som IKKE finnes i OMEGA: {len(missing_in_omega)}')

print("\n" + "="*70)
print("DETALJER PÅ POSTS SOM MANGLER I OMEGA:")
print("="*70)

missing_rader_indices = set()
for post in sorted(missing_in_omega):
    print(f'\nPost: {post}')
    
    # Finn alle RIB-rader som inneholder denne posten
    mask = df_RIB[mengle_col].astype(str).str.contains(post, na=False)
    rader = df_RIB[mask]
    
    total_vekt = rader['PGF_Kalkyle.PGF_Volum'].sum() if 'PGF_Kalkyle.PGF_Volum' in df_RIB.columns else 0
    total_volum = rader['PGF_Kalkyle.PGF_Volum'].sum() if 'PGF_Kalkyle.PGF_Volum' in df_RIB.columns else 0
    
    for idx, row in rader.iterrows():
        missing_rader_indices.add(idx)
        prod_enhet = row.get('FOB_Info.FOB_Produksjonsenhet', 'N/A')
        vekt = row.get('PGF_Kalkyle.PGF_Volum', 0) if 'PGF_Kalkyle.PGF_Volum' in df_RIB.columns else 0
        volum = row.get('PGF_Kalkyle.PGF_Volum', 0) if 'PGF_Kalkyle.PGF_Volum' in df_RIB.columns else 0
        print(f'  - Rad {idx}: Produksjonsenhet={prod_enhet}, Volum={vekt:.3f}')
    
    print(f'  TOTALT: Volum={total_vekt:.3f}')

print("\n" + "="*70)
print(f"\nTotalt antall RIB-rader som inneholder posts som mangler i OMEGA: {len(missing_rader_indices)}")
print(f"Rad-indekser: {sorted(missing_rader_indices)}")

# Sjekk armering også
print("\n" + "="*70)
print("SJEKK: Inneholder armering disse postene?")
print("="*70)

for post in sorted(missing_in_omega):
    arm_rows = df_armering[df_armering['FOB_Info.FOB_Mengdelistepost'].astype(str).str.contains(post, na=False)]
    if not arm_rows.empty:
        total_arm_vekt = arm_rows['PGF_RIB_Info.Vekt'].sum() if 'PGF_RIB_Info.Vekt' in df_armering.columns else 0
        print(f'Post {post}: JA - {len(arm_rows)} rader, total vekt armering={total_arm_vekt:.2f} kg')
    else:
        print(f'Post {post}: NEI')
