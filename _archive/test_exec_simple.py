"""Enkel test: Sjekk exec() og globale variabler"""

# Test 1: Sett variabler og sjekk om de finnes i globals()
excel_file_armering = r"C:\test\armering.xlsx"
excel_file_RIB = r"C:\test\rib.xlsx"
excel_file_OMEGA = r"C:\test\omega.xlsx"

print(f"Test 1 - Direkte variabler:")
print(f"  'excel_file_armering' in globals(): {'excel_file_armering' in globals()}")
print(f"  excel_file_armering = {excel_file_armering}")

# Test 2: Simuler hva exec() gjor
print(f"\nTest 2 - Exec med et lite script:")
test_code = """
print(f"  Inne i exec: 'excel_file_armering' in globals() = {'excel_file_armering' in globals()}")
if 'excel_file_armering' not in globals():
    print("  PROBLEM: Variabelen finnes IKKE i globals() inne i exec!")
    excel_file_armering = "DEFAULT"
else:
    print(f"  OK: Variabelen finnes, verdi = {excel_file_armering}")
"""
exec(test_code)

# Test 3: Exec med globals() eksplisitt
print(f"\nTest 3 - Exec med globals() eksplisitt:")
exec(test_code, globals())
