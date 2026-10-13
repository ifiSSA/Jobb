import pandas as pd

# Last inn data
final = pd.read_csv('resultat_mengdeuttak.csv')
df_omega = pd.read_excel(r'C:\Users\ifi\OneDrive - Scan Survey\LYA_OMEGA_1303_komplett.xlsx', sheet_name='BOQNT', engine='openpyxl')

# Filter OMEGA
df_omega_filtered = df_omega[df_omega["Navn (Prosesskode)"].astype(str).str.startswith("1303", na=False)]
df_omega_armering = df_omega_filtered[df_omega_filtered["Tittel"].astype(str).str.contains("Armering", case=False, na=False)]
df_omega_betong = df_omega_filtered[df_omega_filtered["Tittel"].astype(str).str.contains("Betong", case=False, na=False)]

omega_armering_codes = set(df_omega_armering["Navn (Prosesskode)"].astype(str).unique())
omega_betong_codes = set(df_omega_betong["Navn (Prosesskode)"].astype(str).unique())

# Funksjon
def find_all_mengdelistepost_groups(row, omega_armering_codes, omega_betong_codes):
    """
    Returnér liste av tuples: [(hovedkode, armering_post, betong_post), ...]
    """
    
    # Få alle Mengdelistepost-verdier fra raden
    meng_cols = [c for c in final.columns if c.startswith("Mengdelistepost_")]
    meng_values = [str(row[col]).strip() if pd.notna(row[col]) else None for col in meng_cols]
    meng_values = [v for v in meng_values if v and v != "nan"]
    
    print(f"Alle mengdelistepost-verdier: {meng_values}")
    
    if not meng_values:
        return [(None, None, None)]
    
    # Første verdi er hovedkoden
    hovedkode = meng_values[0]
    
    # Identifiser hovedkode-familien: bruk lengda av hovedkoden som grunnlag
    # Hvis hovedkode er "1303.01.05.241.17" (15 tegn), familien er "1303.01.05.241.17"
    # Hvis posts starter med det, de er i samme familie
    def get_family(code, ref_code=None):
        if ref_code is None:
            # Bruk hele koden som familie
            return str(code)
        else:
            # For andre koder, bruk lengda av ref_code for å bestemme familie
            ref_len = len(str(ref_code))
            return str(code)[:ref_len] if len(str(code)) >= ref_len else str(code)
    
    hovedfamily = hovedkode  # Bruk hele hovedkoden som familienavn
    print(f"Hovedkode: {hovedkode}, Familie: {hovedfamily}")
    
    # Finn alle poster som er Armering eller Betong (direkte matches)
    armering_posts = {post for post in meng_values if post in omega_armering_codes}
    betong_posts = {post for post in meng_values if post in omega_betong_codes}
    
    print(f"Armering-poster funnet: {armering_posts}")
    print(f"Betong-poster funnet: {betong_posts}")
    
    # Gruppér posts etter familie: 
    # Først, identifiser alle "top-level" koder (dei utan suffiks, eller der andre koder starter med dei)
    # Då grupperer me alle koder etter kva top-level-kode dei tilhøyrer
    
    def find_base_code(code, all_codes):
        """Finn basekoden som denne koden tilhøyrer"""
        code_str = str(code)
        # Sjekk om denne koden er ein basekode (ingen annan lengre kode starter med denne)
        for other in all_codes:
            other_str = str(other)
            if other_str != code_str and other_str.startswith(code_str):
                # Dette er ein basekode
                return code_str
        # Eller om den starter med ein annan kode
        for other in all_codes:
            other_str = str(other)
            if other_str != code_str and code_str.startswith(other_str):
                # Finn den lengste koden som dette starter med
                # Rekursivt søk
                return find_base_code(other_str, all_codes)
        # Denne er ein basekode
        return code_str
    
    families = {}
    for post in meng_values:
        base = find_base_code(post, meng_values)
        if base not in families:
            families[base] = []
        families[base].append(post)
    
    
    print(f"Familier: {families}")
    
    result = []
    
    # For hver familie, lag ei gruppe BERRE hvis han har armering eller betong
    for family, posts in families.items():
        family_arm = None
        family_bet = None
        family_hovedkode = family  # Bruk basekoden som hovedkode for denne familien
        
        for post in posts:
            if post in armering_posts and family_arm is None:
                family_arm = post
            elif post in betong_posts and family_bet is None:
                family_bet = post
        
        # Legg til denne gruppen BERRE hvis den har armering eller betong
        if family_arm or family_bet:
            result.append((family_hovedkode, family_arm, family_bet))
            if family == hovedfamily:
                print(f"Hovedkode-gruppe: ({family_hovedkode}, {family_arm}, {family_bet})")
            else:
                print(f"Annen gruppe: ({family_hovedkode}, {family_arm}, {family_bet})")
    
    
    return result if result else [(None, None, None)]

# Test med PE-LP092-V-1012
row = final[final['FOB_Produksjonsenhet_rep'] == 'PE-LP092-V-1012'].iloc[0]
print("\n=== PE-LP092-V-1012 ===")
result = find_all_mengdelistepost_groups(row, omega_armering_codes, omega_betong_codes)
print(f"\nResultat: {result}")
