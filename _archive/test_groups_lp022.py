import pandas as pd

# Simuler en rad fra final for PE-LP022-V-0007
meng_cols = ["Mengdelistepost_1", "Mengdelistepost_2", "Mengdelistepost_3"]

row_data = {
    "Mengdelistepost_1": "1302.03.06.241.3",
    "Mengdelistepost_2": "1302.03.06.241.3.1",
    "Mengdelistepost_3": "1302.03.06.241.3.4",
    "PGF_Info.PGF_Fagkapittel": "06 Betongkonstruksjoner, prefab"
}

# OMEGA-data
df = pd.read_excel(r"C:\Users\ifi\OneDrive - Scan Survey\VEK_OMEGA_1302_komplett.xlsx", sheet_name="BOQNT")

omega_armering_codes = set(df[df["Tittel"].astype(str).str.contains('Armering', case=False, na=False)]["Navn (Prosesskode)"].astype(str).unique())
omega_betong_codes = set(df[(df["Tittel"].astype(str).str.contains('BETONG', case=False, na=False)) & 
                            (df["Enhet"].astype(str).str.lower() == "m³")]["Navn (Prosesskode)"].astype(str).unique())

print(f"1302.03.06.241.3.20 i omega_armering_codes: {'1302.03.06.241.3.20' in omega_armering_codes}")
print(f"1302.03.06.241.3.21 i omega_betong_codes: {'1302.03.06.241.3.21' in omega_betong_codes}")

def find_all_mengdelistepost_groups(row, omega_armering_codes, omega_betong_codes, meng_cols):
    # Få alle Mengdelistepost-verdier fra raden
    meng_values = [str(row[col]).strip() if pd.notna(row.get(col)) else None for col in meng_cols]
    meng_values = [v for v in meng_values if v and v != "nan"]
    
    print(f"\nmeng_values: {meng_values}")
    
    if not meng_values:
        return [(None, None, None)]
    
    # Finn alle unike familie-prefikser
    families = {}
    
    for code in meng_values:
        # Skip hvis denne koden er i armering/betong-listene
        if code in omega_armering_codes or code in omega_betong_codes:
            print(f"  Skipping {code} (finnes i omega_armering_codes eller omega_betong_codes)")
            continue
        
        parts = code.split(".")
        if len(parts) >= 4:
            family_prefix = ".".join(parts[:4])
            
            if family_prefix not in families:
                families[family_prefix] = []
            families[family_prefix].append(code)
            print(f"  Added {code} to family {family_prefix}")
    
    if not families:
        print("  No families found!")
        return [(None, None, None)]
    
    print(f"\nFamilies: {families}")
    
    result_groups = []
    
    for family_prefix in sorted(families.keys()):
        codes_in_family = families[family_prefix]
        
        codes_sorted = sorted(codes_in_family, key=lambda x: x.count("."))
        hovedkode = codes_sorted[0]
        
        print(f"\nFamily {family_prefix}:")
        print(f"  Hovedkode: {hovedkode}")
        
        armering_post = None
        betong_post = None
        
        for post in meng_values:
            if post.startswith(hovedkode + "."):
                print(f"  Checking {post} (starts with {hovedkode}.)")
                if post in omega_armering_codes and armering_post is None:
                    armering_post = post
                    print(f"    -> Armering post!")
                elif post in omega_betong_codes and betong_post is None:
                    betong_post = post
                    print(f"    -> Betong post!")
        
        print(f"  Result: ({hovedkode}, {armering_post}, {betong_post})")
        result_groups.append((hovedkode, armering_post, betong_post))
    
    return result_groups

groups = find_all_mengdelistepost_groups(row_data, omega_armering_codes, omega_betong_codes, meng_cols)
print(f"\n\nFinal groups: {groups}")

# Sjekk fagkapittel filter
fagkapittel = row_data.get("PGF_Info.PGF_Fagkapittel", "")
har_betong_i_fagkapittel = "betong" in str(fagkapittel).lower()
print(f"\nFagkapittel: {fagkapittel}")
print(f"Har betong i fagkapittel: {har_betong_i_fagkapittel}")

groups_med_betong = []
for h, a, b in groups:
    print(f"\nChecking group ({h}, {a}, {b})")
    if b is not None:
        print(f"  -> Has betong-post, including")
        groups_med_betong.append((h, a, b))
    elif har_betong_i_fagkapittel and h is not None:
        print(f"  -> Has betong in fagkapittel but no betong-post, checking løsmasser...")
        # Sjekk løsmasser
        hoved_str = str(h)
        omega_subset = df[df["Navn (Prosesskode)"].astype(str).str.startswith(hoved_str)]
        if len(omega_subset) > 0:
            losmassereferanse = omega_subset["Tittel"].astype(str).str.contains("løsmasser", case=False).any()
            print(f"    Løsmasser referanse: {losmassereferanse}")
            if not losmassereferanse:
                print(f"    -> Not løsmasser, including")
                groups_med_betong.append((h, a, b))
        else:
            print(f"    -> No OMEGA subset found, including")
            groups_med_betong.append((h, a, b))

print(f"\n\nGroups med betong: {groups_med_betong}")
