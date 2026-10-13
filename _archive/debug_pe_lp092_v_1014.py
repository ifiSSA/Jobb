import pandas as pd

final = pd.read_csv('resultat_mengdeuttak.csv')
meng_cols = [c for c in final.columns if c.startswith('Mengdelistepost_')]

# Finn PE-LP092-V-1014
rad = final[final['FOB_Produksjonsenhet_rep'] == 'PE-LP092-V-1014'].iloc[0]

print("PE-LP092-V-1014:")
print(f"Total_Armering_Vekt: {rad['Total_Armering_Vekt']}")
print("\nAlle Mengdelistepost-verdiar:")
meng_values = []
for col in meng_cols:
    val = rad[col]
    if pd.notna(val):
        print(f"  {col}: {val}")
        meng_values.append(str(val).strip())

print(f"\nSummert list: {meng_values}")

# No sjekk basekode-logikken
def find_base_code(code, all_codes):
    """Finn basekoden som denne koden tilhøyrer"""
    code_str = str(code)
    code_parts = code_str.split('.')
    
    # Sjekk om denne koden er ein basekode
    for other in all_codes:
        other_str = str(other)
        other_parts = other_str.split('.')
        
        # other startar med code (med punkt-respekt) og er lengre?
        if other_str != code_str and len(other_parts) > len(code_parts):
            if other_parts[:len(code_parts)] == code_parts:
                # other startar med code (med punkt-respekt)
                return code_str
    
    # Eller sjekk om denne koden startar med ein annan kode (med punkt-respekt)
    for other in all_codes:
        other_str = str(other)
        other_parts = other_str.split('.')
        
        if other_str != code_str and len(code_parts) > len(other_parts):
            if code_parts[:len(other_parts)] == other_parts:
                # code startar med other (med punkt-respekt)
                return find_base_code(other_str, all_codes)
    
    # Denne er ein basekode
    return code_str

# Gruppér
families = {}
for post in meng_values:
    base = find_base_code(post, meng_values)
    if base not in families:
        families[base] = []
    families[base].append(post)

print(f"\nFamilier funnet:")
for family, posts in families.items():
    print(f"  {family}: {posts}")

print(f"\nAntall familier: {len(families)}")
