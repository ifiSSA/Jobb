import pandas as pd

excel_file_OMEGA = r"C:\Users\ifi\OneDrive - Scan Survey\LYA_OMEGA_1303_komplett.xlsx"

# Sjekk arkene
xls = pd.ExcelFile(excel_file_OMEGA)
print("Sheets:", xls.sheet_names)

# Les første ark
df_omega = pd.read_excel(excel_file_OMEGA, sheet_name=0, engine="openpyxl")
print("\nKolonner:")
print(df_omega.columns.tolist())

print("\nShape:", df_omega.shape)

print("\nFørste 5 rader:")
print(df_omega.head())

# Sjekk "Tittel" og "Navn (Prosesskode)" spesifikt
if "Tittel" in df_omega.columns:
    print("\nUnike Tittel-verdier som inneholder 'Betong' eller 'Armering':")
    mask = df_omega["Tittel"].astype(str).str.contains("Betong|Armering", case=False, na=False)
    print(df_omega[mask][["Tittel", "Navn (Prosesskode)" if "Navn (Prosesskode)" in df_omega.columns else ""]].drop_duplicates().head(10))
