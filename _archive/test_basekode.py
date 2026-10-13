import pandas as pd

# Test av basekode-logikk
def find_base_code_old(code, all_codes):
    """Gammal versjon"""
    code_str = str(code)
    # Sjekk om denne koden er ein basekode
    for other in all_codes:
        other_str = str(other)
        if other_str != code_str and other_str.startswith(code_str):
            return code_str
    # Eller om den starter med ein annan kode
    for other in all_codes:
        other_str = str(other)
        if other_str != code_str and code_str.startswith(other_str):
            return find_base_code_old(other_str, all_codes)
    return code_str

# Test
test_codes = ['1303.01.05.241.1', '1303.01.05.241.10', '1303.01.05.241.1.2', '1303.01.05.241.10.3']

print("Gamle logikk:")
for code in test_codes:
    base = find_base_code_old(code, test_codes)
    print(f"{code} -> {base}")

# Nå, det som skjer:
# 1303.01.05.241.1 -> sjekkar om 1303.01.05.241.10 startar med 1303.01.05.241.1
#   JA! Fordi "1303.01.05.241.10" startar med "1303.01.05.241.1" (string-wise)
#   So den returnerer 1303.01.05.241.1 som base
# 1303.01.05.241.10 -> sjekkar om 1303.01.05.241.1 startar med 1303.01.05.241.10
#   NEI. Sjekkar om 1303.01.05.241.10 startar med 1303.01.05.241.1
#   JA! So den kallar find_base_code_old('1303.01.05.241.1')
#   Og får attende 1303.01.05.241.1

print("\nProblemet: '1303.01.05.241.10' startar med '1303.01.05.241.1' (string-wise)")
print("Men dei bør IKKJE vere i same familie!")
print("1303.01.05.241.1 og 1303.01.05.241.1.2 skulle vere i same familie")
print("Men 1303.01.05.241.1 og 1303.01.05.241.10 IKKJE")

def find_base_code_new(code, all_codes):
    """Ny versjon med punkt-respekt"""
    code_str = str(code)
    code_parts = code_str.split('.')
    
    # Sjekk om denne koden er ein basekode (det finnes ikkje lengre koder som startar med denne, med punkt-respekt)
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
                return find_base_code_new(other_str, all_codes)
    
    # Denne er ein basekode
    return code_str

print("\nNy logikk (med punkt-respekt):")
for code in test_codes:
    base = find_base_code_new(code, test_codes)
    print(f"{code} -> {base}")
