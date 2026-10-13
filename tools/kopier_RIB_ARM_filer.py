import os
import shutil
import argparse
from datetime import datetime
from pathlib import Path

# --- Standardverdier (brukes om ikke oppgitt via CMD) ---
DEFAULT_KILDE = r"C:\Users\ifi\DC\ACCDocs\Multiconsult Norge AS\Fornebubanen\Project Files\04_LEVERANSER\05_ARBEIDSUNDERLAG_ENTREPRENOR\K10\01_Samlet leveranse (gjeldende)"
DEFAULT_MAL = r"C:\Users\ifi\Documents\Modell_lokalt"

# --- Kommandolinje-argumenter ---
parser = argparse.ArgumentParser(
    description="Kopierer ARM/RIB IFC-filer fra kildemappe til lokal målmappe, organisert per stasjon.",
    formatter_class=argparse.RawTextHelpFormatter,
    epilog=(
        "Eksempel:\n"
        "  python kopier_RIB_ARM_filer.py\n"
        "  python kopier_RIB_ARM_filer.py --kilde \"C:\\sti\\til\\kilde\"\n"
        "  python kopier_RIB_ARM_filer.py --kilde \"C:\\sti\\til\\kilde\" --mal \"C:\\lokal\\mappe\"\n"
    ),
)
parser.add_argument(
    "--kilde",
    default=DEFAULT_KILDE,
    help=f"Kildemappe å søke i.\nStandard: {DEFAULT_KILDE}",
)
parser.add_argument(
    "--mal",
    default=DEFAULT_MAL,
    help=f"Lokal målmappe å kopiere til.\nStandard: {DEFAULT_MAL}",
)
args = parser.parse_args()

kilde_mappe = args.kilde
mal_mappe = args.mal

# Søkekriterier - RIB filer som slutter på .ifc
sokemonstre_type = ["RIB", "rib", "Rib"]


def fmt_dato(tidsstempel):
    """Formaterer et Unix-tidsstempel til lesbar streng."""
    return datetime.fromtimestamp(tidsstempel).strftime("%Y-%m-%d %H:%M:%S")


def fmt_storrelse(antall_bytes):
    """Formaterer filstørrelse til KB eller MB."""
    if antall_bytes >= 1_048_576:
        return f"{antall_bytes / 1_048_576:.1f} MB"
    return f"{antall_bytes / 1024:.1f} KB"


def hent_stasjonskode(fil_navn):
    """
    Henter stasjonskoden fra filnavn for sortering i undermapper.
    Formatet er: U_F_STA_SKY_RIB_XXX_K10-LP026.ifc
    Stasjonskoden er fjerde segment (index 3) skilt med underscore.
    Eksempel: U_F_STA_SKY_RIB_XXX_K10-LP026.ifc -> SKY
              U_F_STA_LYA_RIB_XXX_K10-LP039.ifc -> LYA
    """
    deler = fil_navn.split("_")
    if len(deler) >= 4:
        return deler[3].upper()
    return "UKJENT"


def versjonssjekk(kilde_sti, mal_sti):
    """
    Sammenligner kilde og lokal fil. Returnerer en av:
      'samme'       - identisk storrelse og kilde ikke nyere enn lokal
      'kilde_nyere' - kildefilen er endret etter lokal kopi
      'lokal_nyere' - lokal fil er nyere enn kilde (uventet)
      'ulik_storrelse' - stoerrelse er ulik (ny versjon)
    """
    ks = os.stat(kilde_sti)
    ms = os.stat(mal_sti)

    storrelse_lik = ks.st_size == ms.st_size
    kilde_nyere = ks.st_mtime > ms.st_mtime + 1   # 1 sek toleranse for filsystem-avrunding

    if not storrelse_lik:
        return "ulik_storrelse", ks, ms
    if kilde_nyere:
        return "kilde_nyere", ks, ms
    if ms.st_mtime > ks.st_mtime + 1:
        return "lokal_nyere", ks, ms
    return "samme", ks, ms


# --- Oppstart ---
os.makedirs(mal_mappe, exist_ok=True)

print("=" * 70)
print("KOPIERING AV ARM/RIB IFC FILER - ORGANISERT PER STASJON")
print("=" * 70)
print(f"\nKildemappe : {kilde_mappe}")
print(f"Malmappe   : {mal_mappe}")
print(f"\nSoker etter RIB-filer som ender paa .ifc")
print("Organiserer etter stasjonskode (del 4 i filnavn, skilt med _)")
print("=" * 70 + "\n")

# --- Finn filer ---
funnet_filer = []
kopierte_filer = []
hoppet_over_filer = []
feilede_filer = []

print("[INFO] Soker gjennom filer...")

for root, dirs, files in os.walk(kilde_mappe):
    for fil_navn in files:
        inneholder_type = any(m in fil_navn for m in sokemonstre_type)
        if inneholder_type and fil_navn.lower().endswith(".ifc"):
            funnet_filer.append((os.path.join(root, fil_navn), fil_navn))

print(f"[INFO] Funnet {len(funnet_filer)} RIB-filer\n")

if len(funnet_filer) == 0:
    print("[ADVARSEL] Ingen RIB IFC-filer funnet - sjekk at kildemappen er riktig.")
else:
    for fil_sti, fil_navn in funnet_filer:
        try:
            stasjon = hent_stasjonskode(fil_navn)
            stasjon_mappe = os.path.join(mal_mappe, stasjon)
            os.makedirs(stasjon_mappe, exist_ok=True)
            mal_sti = os.path.join(stasjon_mappe, fil_navn)

            if os.path.exists(mal_sti):
                status, ks, ms = versjonssjekk(fil_sti, mal_sti)

                if status == "samme":
                    print(
                        f"[UENDRET]   {stasjon}/{fil_navn}"
                        f"  ({fmt_storrelse(ks.st_size)}, "
                        f"kilde: {fmt_dato(ks.st_mtime)}, "
                        f"lokal: {fmt_dato(ms.st_mtime)})"
                    )
                    hoppet_over_filer.append((stasjon, fil_navn))
                    continue

                elif status == "kilde_nyere":
                    print(
                        f"[OPPDATERER] {stasjon}/{fil_navn}"
                        f"  (kilde {fmt_dato(ks.st_mtime)} > lokal {fmt_dato(ms.st_mtime)}, "
                        f"kilde {fmt_storrelse(ks.st_size)})"
                    )

                elif status == "lokal_nyere":
                    print(
                        f"[ADVARSEL]   {stasjon}/{fil_navn}"
                        f"  - lokal fil er NYERE enn kilde! "
                        f"(lokal {fmt_dato(ms.st_mtime)} > kilde {fmt_dato(ks.st_mtime)}) "
                        f"- kopierer likevel"
                    )

                elif status == "ulik_storrelse":
                    print(
                        f"[OPPDATERER] {stasjon}/{fil_navn}"
                        f"  (storrelse endret: kilde {fmt_storrelse(ks.st_size)} "
                        f"vs lokal {fmt_storrelse(ms.st_size)}, "
                        f"kilde {fmt_dato(ks.st_mtime)})"
                    )
            else:
                ks = os.stat(fil_sti)
                print(
                    f"[NY FIL]    {stasjon}/{fil_navn}"
                    f"  ({fmt_storrelse(ks.st_size)}, {fmt_dato(ks.st_mtime)})"
                )

            shutil.copy2(fil_sti, mal_sti)
            kopierte_filer.append((stasjon, fil_navn))

        except Exception as e:
            print(f"[FEIL] Kunne ikke kopiere {fil_navn}: {e}")
            feilede_filer.append((fil_navn, str(e)))

    # --- Oppsummering ---
    print("\n" + "=" * 70)
    print("OPPSUMMERING")
    print("=" * 70)
    print(f"Totalt funnet   : {len(funnet_filer)} filer")
    print(f"Kopiert/oppdatert: {len(kopierte_filer)} filer")
    print(f"Uendret (hoppet over): {len(hoppet_over_filer)} filer")
    print(f"Feilet          : {len(feilede_filer)} filer")

    if feilede_filer:
        print(f"\nFiler som feilet:")
        for fil, feil in feilede_filer:
            print(f"  - {fil}: {feil}")

    if kopierte_filer:
        stasjoner = {}
        for stasjon, fil in kopierte_filer:
            stasjoner.setdefault(stasjon, []).append(fil)

        print(f"\nKopierte/oppdaterte filer per stasjon:")
        for stasjon in sorted(stasjoner.keys()):
            print(f"  [{stasjon}]")
            for fil in stasjoner[stasjon]:
                print(f"    - {fil}")

    print("\n" + "=" * 70)
    print(f"[SUCCESS] Filer kopiert og organisert til: {mal_mappe}")
    print("=" * 70)
