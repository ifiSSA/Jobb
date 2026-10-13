"""
masse_oppsummering.py
---------------------
Leser 'Summasjon leveranser K2B' fra Veidekke massehåndtering .xlsm og
lager en daglig oppsummering gruppert per Dato + VD Prosess.
Hvert lossested faar sin egen kolonne med tonn for den kombinasjonen.

Bruk:
    python masse_oppsummering.py
    python masse_oppsummering.py --kilde "C:\\sti\\til\\fil.xlsm"
    python masse_oppsummering.py --kilde "C:\\sti\\til\\fil.xlsm" --ut "C:\\resultat.xlsx"
"""

import os
import argparse
import datetime
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- Standard stier ---
DEFAULT_MAPPE   = r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-03 K2B - Tunnel og grunnarbeider Vækerø - Lysaker\Målebrev\Dokumentasjon Tonn fra VD"
DEFAULT_UT      = os.path.join(DEFAULT_MAPPE, "FOB_Daglig_oppsummering_v11.xlsx")
DEFAULT_FORRIGE = os.path.join(DEFAULT_MAPPE, "FOB_Daglig_oppsummering_v10.xlsx")

# --- Kjente lossestedene i ønsket rekkefølge ---
KJENTE_LOSSESTEDENE = [
    "NOAH Engadalen",
    "Lindum Oredalen",
    "Lindum Egge",
    "Lindum Drammen",
    "NG Kopstad",
    "NOMAS, Ryghkollen",
    "ROAF Berger",
    "Røyken Grovfyllplass",
    "Elvestad",
    "Veidekke, Elvestad",
    "Feiring, Hobøl",
]
PALETTE = [
    "D6E4F0",  # lysebla
    "D5F5E3",  # lysegron
    "FCF3CF",  # lysgul
    "FADBD8",  # lyserosa
    "E8DAEF",  # lyslilla
    "D6EAF8",  # himmelblaa
    "FDEBD0",  # lys oransje
    "D1F2EB",  # mint
]


def finn_kildefil(mappe):
    """Finner forste .xlsm-fil i mappen som ikke er en tempfil."""
    for f in os.listdir(mappe):
        if f.endswith(".xlsm") and not f.startswith("~"):
            return os.path.join(mappe, f)
    return None


def normaliser_dato(d):
    if isinstance(d, datetime.datetime):
        return d.date()
    if isinstance(d, datetime.date):
        return d
    if isinstance(d, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(d.strip(), fmt).date()
            except ValueError:
                pass
    return None


def normaliser_artikkel(a):
    if not a:
        return "Ukjent"
    a = a.strip()
    if a.lower() == "slam":
        return "Slam"
    if a.lower() in ("uorganisk slam", "slam uorganisk"):
        return "Slam uorganisk"
    return a


def les_data(src_path):
    """Leser kildefilen og returnerer gruppert dict."""
    wb   = openpyxl.load_workbook(src_path, read_only=True, keep_vba=False, data_only=True)
    ws   = wb["Summasjon leveranser K2B"]
    rows = list(ws.iter_rows(values_only=True))

    # nokkel: (dato, vd_prosess)
    # verdi:  {"tonn_total": float, "turer": int, "massetyper": [str],
    #          "tonn_per_lossested": {lossested: float}}
    grouped = defaultdict(lambda: {
        "tonn_total": 0.0,
        "turer": 0,
        "massetyper": [],
        "tonn_per_lossested": defaultdict(float),
    })

    skipped = 0
    for row in rows[1:]:
        if row[0] is None:
            continue
        dato = normaliser_dato(row[0])
        if dato is None:
            skipped += 1
            continue

        lossested  = str(row[7] or "Ukjent").strip()
        artikkel   = normaliser_artikkel(row[11])
        antall     = row[12] if isinstance(row[12], (int, float)) else 0.0
        vd_prosess = str(row[14] or "Ukjent").strip()

        key = (dato, vd_prosess)
        grouped[key]["tonn_total"]  += antall
        grouped[key]["turer"]       += 1
        grouped[key]["massetyper"].append(artikkel)
        grouped[key]["tonn_per_lossested"][lossested] += antall

    print(f"[INFO] Hoppet over {skipped} rader med uleselig dato")
    return grouped


def les_kommentarer(v10_path):
    """Leser 'Kommentar FOB'-kolonne fra forrige versjon av oppsummeringen.
    Returnerer dict med nokkel (dato, vd_prosess) -> kommentar-tekst."""
    if not v10_path or not os.path.isfile(v10_path):
        return {}
    wb = openpyxl.load_workbook(v10_path, read_only=True, data_only=True)
    if "Daglig oppsummering" not in wb.sheetnames:
        return {}
    ws = wb["Daglig oppsummering"]
    rows = list(ws.iter_rows(values_only=True))
    # Forventet header: Dato, Kommentar FOB, VD Prosess, ...
    kommentarer = {}
    for row in rows[1:]:
        if row[0] is None:
            continue
        dato      = normaliser_dato(row[0])
        kommentar = row[1]
        vd        = str(row[2] or "").strip()
        if dato and kommentar and vd:
            kommentarer[(dato, vd)] = str(kommentar).strip()
    return kommentarer


def bygg_excel(grouped, out_path, kommentarer=None):
    """Lager Excel-filen med farger, tykke dato-skillelinjer og en kolonne per lossested."""

    sortert     = sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1]))
    alle_datoer = sorted(set(k[0] for k in grouped))
    dato_farge  = {d: PALETTE[i % len(PALETTE)] for i, d in enumerate(alle_datoer)}

    # Bestem kolonner: kjente lossestedene forst, deretter eventuelle ukjente fra data
    alle_lossestedene_i_data = set()
    for val in grouped.values():
        alle_lossestedene_i_data.update(val["tonn_per_lossested"].keys())
    ukjente = sorted(alle_lossestedene_i_data - set(KJENTE_LOSSESTEDENE))
    lossested_kolonner = KJENTE_LOSSESTEDENE + ukjente

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daglig oppsummering"

    # --- Stiler ---
    navy      = PatternFill("solid", fgColor="1E3070")
    navy_lett  = PatternFill("solid", fgColor="2E4494")  # litt lysere for lossested-header
    hfont     = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    nfont     = Font(name="Calibri", size=10)
    mid       = Alignment(horizontal="center", vertical="center", wrap_text=False)
    lft       = Alignment(horizontal="left",   vertical="center", wrap_text=False)

    thin_s  = Side(style="thin",   color="B0A8A0")
    thick_s = Side(style="medium", color="4A4A4A")

    def bdr(top_thick=False, bot_thick=False):
        return Border(
            left   = thin_s,
            right  = thin_s,
            top    = thick_s if top_thick  else thin_s,
            bottom = thick_s if bot_thick else thin_s,
        )

    # --- Faste kolonner: A=Dato, B=Kommentar FOB, C=VD Prosess, D=Massetyper,
    #                     E=Antall turer, F=Tonn totalt, G=!, H+ = ett lossested per kolonne ---
    FASTE_HEADERS = ["Dato", "Kommentar FOB", "VD Prosess", "Massetyper", "Antall turer", "Tonn totalt", "!"]
    alle_headers  = FASTE_HEADERS + lossested_kolonner
    n_faste       = len(FASTE_HEADERS)

    for c, h in enumerate(alle_headers, 1):
        cell           = ws.cell(row=1, column=c, value=h)
        cell.font      = hfont
        cell.fill      = navy if c <= n_faste else navy_lett
        cell.alignment = mid
        cell.border    = Border(left=thin_s, right=thin_s, top=thick_s, bottom=thick_s)

    # --- Datoskille ---
    dato_seq = [k[0] for k, _ in sortert]
    siste    = {i for i in range(len(dato_seq) - 1) if dato_seq[i] != dato_seq[i + 1]}

    # --- 3-dagers rullende slam-sum ---
    # En dato markeres roed hvis den inngaar i et hvilken som helst 3-dagers
    # vindu (dato-2 .. dato, dato-1 .. dato+1, dato .. dato+2) med sum > 150t.
    SLAM_GRENSE = 150.0
    slam_per_dato = defaultdict(float)
    for (dato_k, vd_k), val_k in grouped.items():
        if "slam" in vd_k.lower():
            slam_per_dato[dato_k] += val_k["tonn_total"]

    roede_datoer = set()
    for dato_k in slam_per_dato:
        for start_offset in range(3):  # vindu starter 0, 1 eller 2 dager for dato_k
            vindu_start = dato_k - datetime.timedelta(days=start_offset)
            vindu_sum = sum(
                slam_per_dato.get(vindu_start + datetime.timedelta(days=j), 0.0)
                for j in range(3)
            )
            if vindu_sum > SLAM_GRENSE:
                for j in range(3):
                    roede_datoer.add(vindu_start + datetime.timedelta(days=j))

    # Utvid: dagen for og etter et rodt vindu markeres roed hvis det ble kjoert slam den dagen
    utvidet = set()
    for d in roede_datoer:
        for delta in (-1, 1):
            nabo = d + datetime.timedelta(days=delta)
            if slam_per_dato.get(nabo, 0.0) > 0:
                utvidet.add(nabo)
    roede_datoer |= utvidet

    roed_fill = PatternFill("solid", fgColor="FFFF4444")

    # --- Datarender ---
    tonn_col_idx = FASTE_HEADERS.index("Tonn totalt") + 1  # 1-basert = 5

    for idx, (key, val) in enumerate(sortert):
        r        = idx + 2
        dato, vd = key
        er_slam  = "slam" in vd.lower()
        roed     = er_slam and dato in roede_datoer
        fill     = roed_fill if roed else PatternFill("solid", fgColor=dato_farge[dato])

        is_first = (idx == 0) or (dato_seq[idx - 1] != dato)
        is_last  = idx in siste
        rad_bdr  = bdr(top_thick=is_first, bot_thick=is_last)

        massetyper_str  = ", ".join(dict.fromkeys(val["massetyper"]))
        kommentar_tekst = (kommentarer or {}).get((dato, vd), None)

        faste_data = [dato, kommentar_tekst, vd, massetyper_str, val["turer"], round(val["tonn_total"], 2), "!" if roed else None]
        for c, v in enumerate(faste_data, 1):
            cell           = ws.cell(row=r, column=c, value=v)
            cell.font      = Font(bold=True, color="FFFFFF" if c == 7 and roed else "000000", name="Calibri", size=10) if c == 7 and roed else nfont
            cell.fill      = fill
            cell.border    = rad_bdr
            cell.alignment = mid if c in (1, 5, 6, 7) else lft

        # Lossested-kolonner
        for li, los in enumerate(lossested_kolonner):
            c    = n_faste + 1 + li
            tonn = val["tonn_per_lossested"].get(los, None)
            v    = round(tonn, 2) if tonn else None
            cell           = ws.cell(row=r, column=c, value=v)
            cell.font      = nfont
            cell.fill      = fill
            cell.border    = rad_bdr
            cell.alignment = mid
            if v is not None:
                cell.number_format = "#,##0.00"

    # --- Formatering ---
    for cell in ws["A"][1:]:
        cell.number_format = "DD.MM.YYYY"

    from openpyxl.utils import get_column_letter
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 40
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 5
    for li in range(len(lossested_kolonner)):
        col_letter = get_column_letter(n_faste + 1 + li)
        ws.column_dimensions[col_letter].width = 18

    total_cols = n_faste + len(lossested_kolonner)
    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{len(sortert)+1}"

    wb.save(out_path)
    print(f"[SUCCESS] Lagret: {out_path}")
    print(f"[INFO] {len(alle_datoer)} datoer, {len(sortert)} rader, {len(lossested_kolonner)} lossested-kolonner")
    if ukjente:
        print(f"[INFO] Ukjente lossestedene lagt til bakerst: {ukjente}")


def main():
    parser = argparse.ArgumentParser(
        description="Lager daglig masseoppsummering fra Veidekke xlsm.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "Eksempel:\n"
            "  python masse_oppsummering.py\n"
            "  python masse_oppsummering.py --kilde \"C:\\sti\\fil.xlsm\"\n"
            "  python masse_oppsummering.py --kilde \"C:\\sti\\fil.xlsm\" --ut \"C:\\resultat.xlsx\"\n"
        ),
    )
    parser.add_argument("--kilde", default=None,
                        help="Sti til .xlsm-kildefil. Om ikke oppgitt brukes forste .xlsm i standardmappen.")
    parser.add_argument("--ut",      default=DEFAULT_UT,
                        help=f"Sti til output .xlsx.\nStandard: {DEFAULT_UT}")
    parser.add_argument("--forrige", default=DEFAULT_FORRIGE,
                        help=f"Sti til forrige versjon (.xlsx) med 'Kommentar FOB'.\nStandard: {DEFAULT_FORRIGE}")
    args = parser.parse_args()

    if args.kilde:
        src = args.kilde
        if not os.path.isfile(src):
            print(f"[FEIL] Finner ikke kildefilen: {src}")
            return
    else:
        src = finn_kildefil(DEFAULT_MAPPE)
        if not src:
            print(f"[FEIL] Ingen .xlsm-fil funnet i: {DEFAULT_MAPPE}")
            return

    print(f"[INFO] Kildefil:  {src}")
    print(f"[INFO] Outputfil: {args.ut}")

    kommentarer = les_kommentarer(args.forrige)
    if kommentarer:
        print(f"[INFO] Leste {len(kommentarer)} kommentarer fra: {args.forrige}")
    else:
        print(f"[INFO] Ingen kommentarer lastet (forrige-fil ikke funnet eller tom)")

    grouped = les_data(src)
    bygg_excel(grouped, args.ut, kommentarer)


if __name__ == "__main__":
    main()
