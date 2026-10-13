import os
import shutil
import argparse

def kopier_dwg_filer(kilde, destinasjon):
    # Sørg for at destinasjonsmappen eksisterer
    if not os.path.exists(destinasjon):
        os.makedirs(destinasjon)

    # Gå gjennom alle undermapper og filer i kildemappen
    for rot, _, filer in os.walk(kilde):
        for fil in filer:
            if fil.lower().endswith('.dwg'):
                kilde_fil = os.path.join(rot, fil)
                dest_fil = os.path.join(destinasjon, fil)

                # Kopierer filen til destinasjonsmappen
                shutil.copy2(kilde_fil, dest_fil)
                print(f"Kopiert: {kilde_fil} til {dest_fil}")

def main():
    parser = argparse.ArgumentParser(description="Kopier DWG-filer fra en mappe til en annen.")
    parser.add_argument("kildemappe", help="Sti til kildemappen som skal leses.")
    parser.add_argument("destinasjonsmappe", help="Sti til destinasjonsmappen hvor filene skal kopieres.")

    args = parser.parse_args()

    kopier_dwg_filer(args.kildemappe, args.destinasjonsmappe)

if __name__ == "__main__":
    main()