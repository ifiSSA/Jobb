"""Test-script for å sjekke hvordan exec() håndterer globale variabler"""

import tempfile
import subprocess
import os
import sys

_ARMERING = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Mengdeuttak_Armering.py")

# De samme filene som i Mengdeuttak_Armering.py
arm_file = r"C:\Users\ifi\OneDrive - Scan Survey\VEK_Armering_LP01_LP02AB.xlsx"
rib_file = r"C:\Users\ifi\OneDrive - Scan Survey\VEK_RIB_LP01_LP02AB.xlsx"
omega_file = r"C:\Users\ifi\OneDrive - Scan Survey\VEK_OMEGA_1302_komplett.xlsx"

# Wrapper-script (samme som GUI bruker)
wrapper_script = f"""
import sys

# Sett filstiene direkte i dette scriptets globale scope
excel_file_armering = r"{arm_file}"
excel_file_RIB = r"{rib_file}"
excel_file_OMEGA = r"{omega_file}"

# Debug: Sjekk hva som er i globals()
print("=== DEBUG: Variabler satt i wrapper ===")
print(f"excel_file_armering: {{excel_file_armering}}")
print(f"excel_file_RIB: {{excel_file_RIB}}")
print(f"excel_file_OMEGA: {{excel_file_OMEGA}}")
print(f"'excel_file_armering' in globals(): {{'excel_file_armering' in globals()}}")

# Kjør hovedscriptet
print("\\n=== Kjører hovedscriptet ===")
with open(r"{_ARMERING}", "r", encoding="utf-8") as f:
    script_content = f.read()
    # Debug: Sjekk igjen rett før exec
    print(f"Rett før exec - 'excel_file_armering' in globals(): {{'excel_file_armering' in globals()}}")
    exec(script_content)
"""

# Skriv wrapper-scriptet til temp-fil
with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False, encoding='utf-8') as f:
    f.write(wrapper_script)
    temp_script = f.name

print(f"Temp-script: {temp_script}")
print("=" * 60)

# Kjør wrapper-scriptet
result = subprocess.run(
    [sys.executable, temp_script],
    capture_output=True,
    text=True,
    encoding='utf-8',
    errors='replace'
)

print("STDOUT:")
if result.stdout:
    print(result.stdout[:3000] if len(result.stdout) > 3000 else result.stdout)
else:
    print("(tom)")
print("\nSTDERR:")
if result.stderr:
    print(result.stderr[:1000] if len(result.stderr) > 1000 else result.stderr)
else:
    print("(tom)")
print(f"\nReturn code: {result.returncode}")

# Slett temp-filen
import os
try:
    os.remove(temp_script)
except:
    pass
