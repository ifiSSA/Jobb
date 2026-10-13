import openpyxl
import csv

# Hent Excel-filnavnet
excel_fil = "C:\\Users\\ifi\\Downloads\\Spreadsheet1 (4).xlsx"

# Filnavn på CSV-filen hvor resultatene skal lagres
output_fil = "C:\\Users\\ifi\\Downloads\\resultat4.csv"

try:
    # Åpne Excel-filen
    wb = openpyxl.load_workbook(excel_fil)

    # Ordbok for å lagre verdier fra E65 med antall forekomster og sheets
    e65_verdier = {}

    # Iterer gjennom alle arkene i Excel-filen
    for sheet in wb.sheetnames:
        ark = wb[sheet]
        verdi_e65 = ark['E65'].value  # Hent verdi fra E65

        if verdi_e65:
            if verdi_e65 in e65_verdier:
                e65_verdier[verdi_e65]['antall'] += 1
                e65_verdier[verdi_e65]['ark'].append(sheet)
            else:
                e65_verdier[verdi_e65] = {'antall': 1, 'ark': [sheet]}

    # Skriv de unike verdiene til en CSV-fil
    with open(output_fil, "w", newline='', encoding='utf-8') as fil:
        csv_writer = csv.writer(fil)
        # Skriv topptekst
        csv_writer.writerow(["Fastmerke", "Antall ganger brukt", "Sheet-navn"])

        # Skriv hver verdi og dens detaljer
        for verdi, data in e65_verdier.items():
            csv_writer.writerow([verdi, data['antall'], ", ".join(data['ark'])])

    print(f"Resultatene er lagret i {output_fil}")

except FileNotFoundError:
    print(f"Filen {excel_fil} ble ikke funnet. Sørg for at filnavnet er riktig.")
except Exception as e:
    print(f"En feil oppstod: {e}")
