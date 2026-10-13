# -*- coding: utf-8 -*-
"""
Les PDF-rapporten "Poster i modeller som mangler i kontrakt og PF/EO-logg"
og eksporter tabellen til Excel.

PDF-formatet:
    Tittelbanner  : "Poster i modeller som mangler i kontrakt og PF/EO-logg"
    Tabell-header : Mengdelistepost | Model
    Tabellen gaar over flere sider.

VIKTIG -- kolonneskille:
    PDF-en har ingen loddrette streker mellom kolonnene; de er kun skilt med
    mellomrom. pdfplumber's extract_tables() slaar da begge kolonnene sammen
    til en celle. Derfor leses tabellen ordvis via x-koordinater:

      1. Finn header-linjen (ordene 'Mengdelistepost' og 'Model').
      2. Bruk x0 til 'Model'-ordet som kolonnegrense.
      3. Del hver datalinje: ord til venstre for grensen -> Mengdelistepost,
         ord til hoyre -> Model.

    Grensen huskes paa tvers av sider, saa fortsettelsessider uten gjentatt
    header leses ogsaa riktig.

Bruk:
    # Vindu for aa velge filer (ingen argumenter, eller dobbeltklikk):
    python pdf_manglende_poster_til_excel.py

    # Kommandolinje:
    python pdf_manglende_poster_til_excel.py "rapport.pdf"
    python pdf_manglende_poster_til_excel.py "rapport.pdf" -o "ut.xlsx"
    python pdf_manglende_poster_til_excel.py "rapport.pdf" --debug
    python pdf_manglende_poster_til_excel.py "rapport.pdf" --grense 280

    # Alle PDF-er i en mappe samles i ett ark:
    python pdf_manglende_poster_til_excel.py "C:\\sti\\til\\mappe"

    # Slaa opp FOB_Status fra en eller flere CSV-filer:
    python pdf_manglende_poster_til_excel.py "rapport.pdf" --csv a.csv b.csv

    # RB fra en eller flere kontrakt-Excel-filer (alle ark leses):
    python pdf_manglende_poster_til_excel.py "rapport.pdf" --rb k1.xlsx k2.xlsx

CSV-oppslag (--csv):
    CSV-ene maa ha kolonnene FOB_Info/FOB_Mengdelistepost og
    FOB_Info/FOB_Status (punktum i stedet for skraastrek gaar ogsaa).
    Statusen slaas opp per mengdelistepost og skrives i kolonnen
    'FOB_Status'. Celler med flere poster skilt med ';' splittes.
    Har samme post flere ulike statuser, slaas de sammen: 'S4; S5'.
    Poster som ikke finnes i CSV-ene faar tom status.

Kontraktsjekk (--rb):
    Hver post slaas opp i kontrakten; treffet skrives i 'Kontraktpost'.
    Poster med kjente uttrekksfeil sjekkes ogsaa i rettet form, og
    kommentaren sier om de finnes slik eller ikke:
      - mangler '1' i start           303-03.87.1  -> 1303-03.87.1
      - 010-090, 100 osv. etter '-'   1303-030.87.1 -> 1303-03.87.1
                                      1303-100.87.1 -> 1303-10.87.1
        (000 rettes ikke)
    Begge feil kan opptre samtidig: 303-030.87.13891 -> 1303-03.87.13891.

Ark i utfila:
    Manglende poster    -- alle rader
    Med svar            -- posten finnes i kontrakten (krever --rb)
    S5 ikke i kontrakt  -- status S5 og posten finnes ikke (krever --rb og --csv)
"""

from __future__ import annotations

import argparse
import contextlib
import json
import os
import re
import sys

import pandas as pd
import pdfplumber
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# -- Kolonner --
KOL_POST = "Mengdelistepost"
KOL_MODEL = "Model"
KOL_STATUS = "FOB_Status"
KOL_KOMMENTAR = "Kommentar"
KOL_KONTRAKTPOST = "Kontraktpost"   # posten slik den ble funnet i kontrakten
KOLONNER = [KOL_POST, KOL_MODEL, KOL_KOMMENTAR, "Kildefil", "Side"]

BREDDER = {
    KOL_POST: 30, KOL_MODEL: 60, KOL_STATUS: 16, KOL_KOMMENTAR: 34,
    KOL_KONTRAKTPOST: 30, "Kildefil": 34, "Side": 7,
}

# Merknad for poster som ikke starter paa '1'
_KOMM_MANGLER_1 = "mangler 1 i start"

# Leddet rett etter '-' har faatt en 0 for mye i PDF-uttrekket:
#     303-030.87.13891  ->  1303-03.87.13891
# Tresifret ledd som ender paa 0 har faatt en 0 for mye: 010-090 -> 01-09,
# 100 -> 10 osv. Unntak: 000 rettes ikke (00 er ikke et gyldig ledd).
_FEIL_LEDD = re.compile(r"-((?!00)\d\d)0(?=\.|$)")
_KOMM_FEIL_LEDD = "{feil} skal vaere {rett}"

# Tillegg naar posten er sjekket mot kontrakten i rettet form
_KOMM_FINNES_RETTET = "finnes i kontrakt som {post}"
_KOMM_IKKE_RETTET = "ikke i kontrakt rettet heller"

# Arknavn i utfila
_ARK_ALLE = "Manglende poster"
_ARK_SVAR = "Med svar"               # posten finnes i kontrakten
_ARK_S5 = "S5 ikke i kontrakt"       # status S5 og posten finnes ikke
_STATUS_S5 = "S5"

# -- CSV-oppslag (FOB_Status) --
# Kolonnenavn i CSV-en er "FOB_Info/FOB_Mengdelistepost" og
# "FOB_Info/FOB_Status". Vi matcher paa suffikset etter '/' eller '.', slik at
# begge skrivemaatene (og evt. andre prefiks) treffer.
_CSV_SUFFIX_POST = "fob_mengdelistepost"
_CSV_SUFFIX_STATUS = "fob_status"

# Flere mengdelisteposter i samme celle er skilt med semikolon
_CSV_POST_SKILLE = ";"

# -- RB-oppslag (eget Excel-ark) --
# Arket har formatet:
#     Post nummer           | RB
#     1304.04.04.000.10.1   |
#                           | 0
#                           | 2
# Posten staar paa foerste rad i blokka; RB-verdiene ligger paa radene under
# med tom postkolonne. Vi tar den NEDERSTE RB-verdien i hver blokk.
#
# Prosessarket har andre kolonnenavn, men samme blokkformat:
#     ProsessID: prosess, sted og element | ... | Revisjonsindeks | ...
#
# Alle navn er normalisert: mellomrom fjernet, lowercase.
_RB_KOL_POST = (
    "postnummer",                          # NS3420-ark
    "prosessid:prosess,stedogelement",     # Prosessark
)
_RB_KOL_VERDI_EKSAKT = ("rb",)
_RB_KOL_VERDI_PREFIKS = ("revisjonsind",)  # 'Revisjonsindeks' e.l.
_RB_PREFIKS = "RB"             # skrives foran tallet i kommentaren

# Antall rader fra toppen det lettes etter header-raden i
_MAKS_HEADER_RAD = 10

# Skilletegn mellom flere merknader i Kommentar-kolonnen
_KOMM_SKILLE = "; "

# Ord som identifiserer header-raden (substring-match, ikke eksakt likhet --
# headeren kan vaere "Mengdelistepost (post)" e.l.)
_HDR_ORD_POST = "mengdelistepost"
_HDR_ORD_MODEL = "model"

# Tittelbanner og annen tekst som ikke er data
_IKKE_DATA = (
    "poster i modeller som mangler",
    "pf/eo-logg",
)

# Linjegruppering: ord innenfor denne hoydeforskjellen (pt) regnes som samme linje
_LINJE_TOLERANSE = 3.0

# Trekkes fra 'Model'-ordets x0 for aa gi litt slingringsmonn paa grensen
_GRENSE_MARGIN = 3.0

# -- Excel-stiler (samme palett som pdf_maalebrev_til_excel.py) --
_HDR_FILL = PatternFill("solid", fgColor="1B4AA3")
_HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
_FONT_DATA = Font(name="Calibri", size=10)
_FILL_ANNENHVER = PatternFill("solid", fgColor="EEF2FA")
_VTOPP = Alignment(vertical="top", wrap_text=True)


def _rens(verdi) -> str:
    """Rens tekst: trim og gjor linjeskift/dobbel-space om til ett mellomrom."""
    if verdi is None:
        return ""
    return " ".join(str(verdi).split())


def _grupper_linjer(ord_liste: list) -> list:
    """Grupper ord i tekstlinjer etter y-posisjon.

    Returnerer liste med linjer, hver som liste av ord sortert venstre->hoyre.
    """
    if not ord_liste:
        return []

    sortert = sorted(ord_liste, key=lambda w: (w["top"], w["x0"]))
    linjer = []
    naavaerende = [sortert[0]]

    for ord_ in sortert[1:]:
        if abs(ord_["top"] - naavaerende[0]["top"]) <= _LINJE_TOLERANSE:
            naavaerende.append(ord_)
        else:
            linjer.append(sorted(naavaerende, key=lambda w: w["x0"]))
            naavaerende = [ord_]
    linjer.append(sorted(naavaerende, key=lambda w: w["x0"]))

    return linjer


def _finn_kolonnegrense(linje: list):
    """Returner x-grensen mellom kolonnene hvis linjen er header, ellers None.

    Grensen er x0 til 'Model'-ordet, minus en liten margin.
    """
    tekster = [_rens(w["text"]).lower() for w in linje]
    if not any(_HDR_ORD_POST in t for t in tekster):
        return None

    for ord_ in linje:
        if _HDR_ORD_MODEL in _rens(ord_["text"]).lower():
            # Ikke forveksle 'Model' med 'modeller' i tittelbanneret
            if _HDR_ORD_POST in _rens(ord_["text"]).lower():
                continue
            return ord_["x0"] - _GRENSE_MARGIN
    return None


def _er_tittel_linje(linje: list) -> bool:
    """Sjekk om linjen er tittelbanneret eller annen ikke-data-tekst."""
    samlet = " ".join(_rens(w["text"]) for w in linje).lower()
    return any(ord_ in samlet for ord_ in _IKKE_DATA)


def _er_header_linje(linje: list) -> bool:
    """Sjekk om linjen er tabellens header."""
    tekster = [_rens(w["text"]).lower() for w in linje]
    har_post = any(_HDR_ORD_POST in t for t in tekster)
    har_model = any(
        _HDR_ORD_MODEL in t and _HDR_ORD_POST not in t for t in tekster
    )
    return har_post and har_model


def _kommentar(post: str) -> str:
    """Merknad for mengdelisteposten.

    Poster skal starte paa '1' (f.eks. 1302/1303). Starter den paa noe annet
    er det foerste sifferet trolig falt bort under PDF-uttrekket. Leddet etter
    '-' kan ogsaa ha faatt en 0 for mye (030 i stedet for 03).
    """
    merknader = []
    if post and not post.startswith("1"):
        merknader.append(_KOMM_MANGLER_1)
    treff = _FEIL_LEDD.search(post or "")
    if treff:
        merknader.append(_KOMM_FEIL_LEDD.format(
            feil=treff.group(1) + "0", rett=treff.group(1)))
    return _KOMM_SKILLE.join(merknader)


def _post_kandidater(post: str) -> list:
    """Posten slik den staar, fulgt av rettede varianter (uten duplikater).

    Rettinger: '1' foran hvis den mangler, og 030 -> 03 etter '-'.
    Rekkefoelge: original, med 1, rettet ledd, begge deler.
    """
    med_1 = post if post.startswith("1") else "1" + post
    kandidater = [post, med_1,
                  _FEIL_LEDD.sub(r"-\1", post), _FEIL_LEDD.sub(r"-\1", med_1)]
    return list(dict.fromkeys(k for k in kandidater if k))


def _del_linje(linje: list, grense: float) -> tuple:
    """Del linjens ord i (venstre_tekst, hoyre_tekst) ved x-grensen."""
    venstre = [w["text"] for w in linje if w["x0"] < grense]
    hoyre = [w["text"] for w in linje if w["x0"] >= grense]
    return _rens(" ".join(venstre)), _rens(" ".join(hoyre))


def les_pdf(pdf_sti: str, fast_grense: float = None, debug: bool = False) -> list:
    """Les alle sider i PDF-en og returner liste med rad-dicts.

    fast_grense: overstyr automatisk funnet kolonnegrense (x-koordinat).
    debug:       skriv ut hva som blir funnet per side.
    """
    rader = []
    filnavn = os.path.basename(pdf_sti)
    grense = fast_grense   # huskes paa tvers av sider

    with pdfplumber.open(pdf_sti) as pdf:
        for sidenr, side in enumerate(pdf.pages, start=1):
            ord_liste = side.extract_words(
                use_text_flow=False, keep_blank_chars=False
            )
            linjer = _grupper_linjer(ord_liste)

            if debug:
                print(f"    [DEBUG] side {sidenr}: {len(ord_liste)} ord, "
                      f"{len(linjer)} linjer, bredde={side.width:.0f}")

            for linje in linjer:
                # Header: oppdater kolonnegrensen (med mindre den er laast)
                if _er_header_linje(linje):
                    if fast_grense is None:
                        funnet = _finn_kolonnegrense(linje)
                        if funnet is not None:
                            grense = funnet
                            if debug:
                                print(f"    [DEBUG] side {sidenr}: "
                                      f"kolonnegrense x={grense:.1f}")
                    continue

                if _er_tittel_linje(linje):
                    continue

                if grense is None:
                    # Header ikke sett enda -- kan ikke dele linjen
                    continue

                post, model = _del_linje(linje, grense)

                if not post and not model:
                    continue

                # Fortsettelse av forrige rad: tom postkolonne
                if not post and rader:
                    rader[-1][KOL_MODEL] = (
                        f"{rader[-1][KOL_MODEL]} {model}".strip()
                    )
                    continue

                if debug:
                    print(f"    [DEBUG]   '{post}' | '{model}'")

                rader.append({
                    KOL_POST: post,
                    KOL_MODEL: model,
                    KOL_KOMMENTAR: _kommentar(post),
                    "Kildefil": filnavn,
                    "Side": sidenr,
                })

    if grense is None:
        print(f"    [WARNING] {filnavn}: fant aldri header-linjen "
              f"('{KOL_POST}' + '{KOL_MODEL}'). Ingen rader lest.")
        print(f"              Kjor med --debug for aa se hva PDF-en inneholder, "
              f"eller sett grensen manuelt med --grense <x>.")

    return rader


def _finn_kolonne(df: pd.DataFrame, suffix: str):
    """Finn kolonnen hvis navn slutter paa suffix (etter '/' eller '.').

    'FOB_Info/FOB_Status' og 'FOB_Info.FOB_Status' treffer begge paa
    suffix='fob_status'. Returnerer kolonnenavnet, eller None.
    """
    for kol in df.columns:
        siste_ledd = str(kol).replace("/", ".").split(".")[-1].strip().lower()
        if siste_ledd == suffix:
            return kol
    # Fallback: substring-match et sted i navnet
    for kol in df.columns:
        if suffix in str(kol).strip().lower():
            return kol
    return None


def _les_csv(csv_sti: str) -> pd.DataFrame:
    """Les CSV med auto-deteksjon av skilletegn og tegnsett."""
    siste_feil = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(
                csv_sti, sep=None, engine="python", encoding=enc, dtype=str
            )
        except (UnicodeDecodeError, UnicodeError) as feil:
            siste_feil = feil
            continue
    raise RuntimeError(f"Klarte ikke lese {os.path.basename(csv_sti)}: {siste_feil}")


def bygg_status_oppslag(csv_stier: list, debug: bool = False) -> dict:
    """Bygg {mengdelistepost -> status} fra en eller flere CSV-filer.

    - Celler med flere poster (skilt med ';') splittes til en post hver.
    - Har samme post flere ulike statuser, slaas de sammen: 'S4; S5'.
    """
    per_post: dict[str, set] = {}

    for csv_sti in csv_stier:
        navn = os.path.basename(csv_sti)
        df = _les_csv(csv_sti)

        kol_post = _finn_kolonne(df, _CSV_SUFFIX_POST)
        kol_status = _finn_kolonne(df, _CSV_SUFFIX_STATUS)

        if kol_post is None or kol_status is None:
            print(f"    [WARNING] {navn}: fant ikke kolonnene "
                  f"'{_CSV_SUFFIX_POST}' / '{_CSV_SUFFIX_STATUS}'. "
                  f"Kolonner i fila: {list(df.columns)}")
            continue

        if debug:
            print(f"    [DEBUG] {navn}: post='{kol_post}', status='{kol_status}', "
                  f"{len(df)} rader")

        n_par = 0
        for post_celle, status_celle in zip(df[kol_post], df[kol_status]):
            status = _rens(status_celle)
            if not status or status.lower() == "nan":
                continue
            for bit in str(post_celle or "").split(_CSV_POST_SKILLE):
                post = _rens(bit)
                if not post or post.lower() == "nan":
                    continue
                per_post.setdefault(post, set()).add(status)
                n_par += 1

        print(f"    {navn}: {n_par} post/status-par")

    # Slaa sammen flere statuser per post til en sortert streng
    return {post: "; ".join(sorted(statuser))
            for post, statuser in per_post.items()}


def _normaliser_kolonnenavn(navn) -> str:
    """Normaliser kolonnenavn: lowercase uten mellomrom. 'Post nummer' -> 'postnummer'."""
    return "".join(str(navn).split()).lower()


def _formater_rb(verdi) -> str:
    """Formater RB-verdien som tekst. 2.0 -> '2', 2.5 -> '2.5'.

    Merk: 0 er en gyldig verdi og maa ikke behandles som tom.
    """
    tekst = _rens(verdi)
    if not tekst or tekst.lower() == "nan":
        return ""
    try:
        tall = float(tekst.replace(",", "."))
    except ValueError:
        return tekst
    return str(int(tall)) if tall == int(tall) else tekst


def _finn_indeks(navn: list, eksakt: tuple = (), prefiks: tuple = ()):
    """Indeksen til foerste normaliserte kolonnenavn som treffer, ellers None."""
    for i, n in enumerate(navn):
        if n in eksakt or any(n.startswith(p) for p in prefiks):
            return i
    return None


def _les_kontrakt_ark(xlsx_sti: str, debug: bool = False) -> list:
    """Les ALLE ark i en kontrakt-Excel og finn header-raden i hvert ark.

    Leter etter header i de oeverste _MAKS_HEADER_RAD radene, slik at ark med
    tittel-rader over tabellen ogsaa leses riktig. Baade NS3420-ark
    ('Post nummer' + 'RB') og prosessark ('ProsessID: ...' + 'Revisjonsindeks')
    gjenkjennes. Andre ark hoppes over.
    Returnerer liste med (arknavn, postkolonne, rb-kolonne) som Series.
    """
    alle_ark = pd.read_excel(xlsx_sti, sheet_name=None, header=None, dtype=object)

    funnet = []
    for arknavn, raa in alle_ark.items():
        for header_rad in range(min(_MAKS_HEADER_RAD, len(raa))):
            navn = [_normaliser_kolonnenavn(v) for v in raa.iloc[header_rad]]
            i_post = _finn_indeks(navn, eksakt=_RB_KOL_POST)
            i_rb = _finn_indeks(navn, eksakt=_RB_KOL_VERDI_EKSAKT,
                                prefiks=_RB_KOL_VERDI_PREFIKS)
            if i_post is not None and i_rb is not None:
                data = raa.iloc[header_rad + 1:]
                if debug:
                    print(f"    [DEBUG] {os.path.basename(xlsx_sti)} [{arknavn}]: "
                          f"header paa rad {header_rad + 1}, "
                          f"post='{raa.iloc[header_rad, i_post]}', "
                          f"rb='{raa.iloc[header_rad, i_rb]}'")
                funnet.append((arknavn, data.iloc[:, i_post], data.iloc[:, i_rb]))
                break
        else:
            if debug:
                print(f"    [DEBUG] {os.path.basename(xlsx_sti)} [{arknavn}]: "
                      f"ingen post-/RB-kolonner, hoppet over")

    return funnet


def bygg_rb_oppslag(xlsx_stier: list, debug: bool = False) -> tuple:
    """Les en eller flere kontrakt-Excel-filer (alle ark).

    Returnerer (rb_oppslag, alle_poster):
      rb_oppslag  -- {mengdelistepost -> nederste RB-verdi}
      alle_poster -- alle poster i kontrakten, ogsaa de uten RB-verdi

    En rad med verdi i postkolonnen starter en ny blokk. Radene under, med tom
    postkolonne, hoerer til samme post. Siste ikke-tomme RB-verdi i blokka
    vinner -- det er den nederste. Finnes samme post i flere ark/filer, vinner
    den som leses sist.
    """
    per_post: dict[str, str] = {}
    alle_poster: set = set()

    for xlsx_sti in xlsx_stier:
        navn = os.path.basename(xlsx_sti)
        try:
            ark_liste = _les_kontrakt_ark(xlsx_sti, debug=debug)
        except Exception as feil:
            print(f"    [WARNING] {navn}: klarte ikke lese fila: {feil}")
            continue

        if not ark_liste:
            print(f"    [WARNING] {navn}: fant ikke post- og RB-kolonner i noen ark "
                  f"('Post nummer'/'RB' eller 'ProsessID: ...'/'Revisjonsindeks').")
            continue

        for arknavn, post_kol, rb_kol in ark_liste:
            if debug:
                print(f"    [DEBUG] {navn} [{arknavn}]: {len(post_kol)} rader")

            naavaerende_post = None
            n_treff = 0

            for post_celle, rb_celle in zip(post_kol, rb_kol):
                post = _rens(post_celle)
                if post and post.lower() != "nan":
                    naavaerende_post = post
                    alle_poster.add(post)

                if naavaerende_post is None:
                    continue

                rb = _formater_rb(rb_celle)
                if rb != "":                      # 0 er gyldig, "" er ikke
                    # Overskriver bevisst: siste verdi i blokka er den nederste
                    if naavaerende_post not in per_post:
                        n_treff += 1
                    per_post[naavaerende_post] = rb

            print(f"    {navn} [{arknavn}]: {n_treff} poster med RB-verdi")

    return per_post, alle_poster


def _slaa_sammen_kommentar(*deler) -> str:
    """Slaa sammen ikke-tomme kommentardeler med semikolon."""
    return _KOMM_SKILLE.join(d for d in deler if d)


def _finn_pdf_filer(sti: str) -> list:
    """Returner liste med PDF-stier fra en fil- eller mappesti."""
    if os.path.isdir(sti):
        return sorted(
            os.path.join(sti, f)
            for f in os.listdir(sti)
            if f.lower().endswith(".pdf")
        )
    return [sti]


def _formater_ark(ark, df: pd.DataFrame) -> None:
    """Header-farge, kolonnebredder, annenhver-rad, frys og filter."""
    kant = Side(style="thin", color="CCCCCC")
    border = Border(left=kant, right=kant, top=kant, bottom=kant)

    # Header
    for kol_nr, kolonnenavn in enumerate(df.columns, start=1):
        celle = ark.cell(row=1, column=kol_nr)
        celle.fill = _HDR_FILL
        celle.font = _HDR_FONT
        celle.alignment = _HDR_ALIGN
        celle.border = border
        bokstav = get_column_letter(kol_nr)
        ark.column_dimensions[bokstav].width = BREDDER.get(kolonnenavn, 20)

    # Datarader
    for rad_nr in range(2, len(df) + 2):
        annenhver = (rad_nr % 2 == 0)
        for kol_nr in range(1, len(df.columns) + 1):
            celle = ark.cell(row=rad_nr, column=kol_nr)
            celle.font = _FONT_DATA
            celle.alignment = _VTOPP
            celle.border = border
            if annenhver:
                celle.fill = _FILL_ANNENHVER

    ark.freeze_panes = "A2"
    ark.auto_filter.ref = (
        f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    )


def skriv_excel(ark_data: dict, ut_sti: str) -> None:
    """Skriv {arknavn -> DataFrame} til formaterte Excel-ark, i rekkefoelge."""
    with pd.ExcelWriter(ut_sti, engine="openpyxl") as writer:
        for arknavn, df in ark_data.items():
            df.to_excel(writer, sheet_name=arknavn, index=False)
            _formater_ark(writer.sheets[arknavn], df)


def _lag_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Les 'manglende poster'-PDF og eksporter til Excel."
    )
    parser.add_argument("pdf", help="PDF-fil eller mappe med PDF-filer")
    parser.add_argument(
        "-o", "--output", default=None,
        help="Excel-fil ut (standard: <pdf-navn>.xlsx)",
    )
    parser.add_argument(
        "--csv", nargs="+", default=None, metavar="FIL",
        help=("En eller flere CSV-filer med kolonnene "
              "FOB_Info/FOB_Mengdelistepost og FOB_Info/FOB_Status. "
              f"Legger til kolonnen '{KOL_STATUS}'."),
    )
    parser.add_argument(
        "--rb", nargs="+", default=None, metavar="FIL",
        help=("En eller flere kontrakt-Excel-filer. NS3420-ark ('Post nummer' + "
              "'RB') og prosessark ('ProsessID: prosess, sted og element' + "
              "'Revisjonsindeks') gjenkjennes. Alle ark i hver fil leses. "
              "Nederste RB-verdi per post legges i kommentaren."),
    )
    parser.add_argument(
        "--grense", type=float, default=None,
        help="Overstyr kolonnegrensen manuelt (x-koordinat i punkt)",
    )
    parser.add_argument(
        "--debug", action="store_true",
        help="Skriv ut hva som leses per side",
    )
    return parser


def kjor(args: argparse.Namespace) -> int:
    """Kjoer hele uttrekket. Brukes av baade kommandolinja og vinduet.

    args maa ha feltene: pdf, output, csv, rb, grense, debug.
    """
    if not os.path.exists(args.pdf):
        print(f"[FEIL] Finner ikke: {args.pdf}")
        return 1

    pdf_filer = _finn_pdf_filer(args.pdf)
    if not pdf_filer:
        print(f"[FEIL] Ingen PDF-filer i {args.pdf}")
        return 1

    alle_rader = []
    for pdf_sti in pdf_filer:
        print(f"  Leser {os.path.basename(pdf_sti)} ...")
        try:
            rader = les_pdf(pdf_sti, fast_grense=args.grense, debug=args.debug)
        except Exception as feil:
            print(f"    [WARNING] {os.path.basename(pdf_sti)}: {feil}")
            continue
        print(f"    {len(rader)} rader")
        alle_rader.extend(rader)

    if not alle_rader:
        print("[FEIL] Ingen data funnet. Sjekk at PDF-en har en tabell med "
              "kolonnene 'Mengdelistepost' og 'Model'.")
        return 1

    df = pd.DataFrame(alle_rader, columns=KOLONNER)
    kandidater = [_post_kandidater(p) for p in df[KOL_POST]]

    # --- Sjekk mot kontrakten: finnes posten, og RB-verdi i kommentaren ---
    # Poster med kjente uttrekksfeil (mangler '1' i start, 030 i stedet for
    # 03) sjekkes ogsaa i rettet form.
    alle_poster = None
    if args.rb:
        manglende_rb = [s for s in args.rb if not os.path.exists(s)]
        if manglende_rb:
            print(f"[FEIL] Finner ikke: {', '.join(manglende_rb)}")
            return 1

        print(f"\n  Leser kontrakt fra {len(args.rb)} fil(er) ...")
        rb_oppslag, alle_poster = bygg_rb_oppslag(args.rb, debug=args.debug)
        print(f"    {len(alle_poster)} unike poster i kontrakten, "
              f"{len(rb_oppslag)} med RB-verdi")

        kontraktpost, kommentar = [], []
        n_rettet = n_ikke_rettet = 0
        for post, komm, kand in zip(df[KOL_POST], df[KOL_KOMMENTAR], kandidater):
            treff = next((k for k in kand if k in alle_poster), "")
            tillegg = ""
            if treff and treff != post:
                tillegg = _KOMM_FINNES_RETTET.format(post=treff)
                n_rettet += 1
            elif not treff and len(kand) > 1:
                tillegg = _KOMM_IKKE_RETTET
                n_ikke_rettet += 1

            rb = f"{_RB_PREFIKS} {rb_oppslag[treff]}" if treff in rb_oppslag else ""
            kontraktpost.append(treff)
            kommentar.append(_slaa_sammen_kommentar(komm, tillegg, rb))

        df[KOL_KONTRAKTPOST] = kontraktpost
        df[KOL_KOMMENTAR] = kommentar

        n_treff = sum(1 for k in kontraktpost if k)
        print(f"    [INFO] {n_treff} av {len(df)} rader finnes i kontrakten")
        if n_rettet or n_ikke_rettet:
            print(f"    [INFO] Poster med uttrekksfeil: {n_rettet} finnes i "
                  f"rettet form, {n_ikke_rettet} finnes ikke")

    # --- Slaa opp FOB_Status fra CSV ---
    status_oppslag = None
    if args.csv:
        manglende_csv = [s for s in args.csv if not os.path.exists(s)]
        if manglende_csv:
            print(f"[FEIL] Finner ikke: {', '.join(manglende_csv)}")
            return 1

        print(f"\n  Leser status fra {len(args.csv)} CSV-fil(er) ...")
        status_oppslag = bygg_status_oppslag(args.csv, debug=args.debug)
        print(f"    {len(status_oppslag)} unike mengdelisteposter i CSV")

        # Proev kontraktposten foerst (evt. rettet), saa posten slik den
        # staar, og til slutt de rettede variantene.
        kpost = (df[KOL_KONTRAKTPOST] if KOL_KONTRAKTPOST in df.columns
                 else [""] * len(df))
        statuser = []
        for kp, kand in zip(kpost, kandidater):
            noekler = (kp, *kand)
            statuser.append(next(
                (status_oppslag[n] for n in noekler if n and n in status_oppslag),
                "",
            ))
        df[KOL_STATUS] = statuser

        uten_status = int((df[KOL_STATUS] == "").sum())
        if uten_status:
            print(f"    [INFO] {uten_status} av {len(df)} rader fikk ingen "
                  f"status (posten finnes ikke i CSV-ene)")

    # Kolonnerekkefoelge: Kontraktpost ved siden av posten, status etter Model
    rekkefoelge = [KOL_POST, KOL_KONTRAKTPOST, KOL_MODEL, KOL_STATUS,
                   KOL_KOMMENTAR, "Kildefil", "Side"]
    df = df[[k for k in rekkefoelge if k in df.columns]]

    # Advar hvis Model-kolonnen er tom -- da traff nok kolonnegrensen feil
    tomme_model = int((df[KOL_MODEL] == "").sum())
    if tomme_model:
        print(f"  [WARNING] {tomme_model} rader har tom '{KOL_MODEL}'. "
              f"Kolonnegrensen kan vaere feil -- prov --debug eller --grense.")

    dupl = df.duplicated(subset=[KOL_POST, KOL_MODEL]).sum()
    if dupl:
        print(f"  [INFO] {dupl} duplikate rader (post+modell) beholdt i utdata")

    if args.output:
        ut_sti = args.output
    else:
        base = (os.path.basename(os.path.normpath(args.pdf))
                if os.path.isdir(args.pdf)
                else os.path.splitext(os.path.basename(args.pdf))[0])
        ut_sti = os.path.join(
            os.path.dirname(os.path.abspath(args.pdf)), f"{base}.xlsx"
        )

    # --- Ark: alle, med svar (finnes i kontrakt), S5 som ikke finnes ---
    ark_data = {_ARK_ALLE: df}
    if alle_poster is not None:
        i_kontrakt = df[KOL_KONTRAKTPOST] != ""
        ark_data[_ARK_SVAR] = df[i_kontrakt]
        if status_oppslag is not None:
            # \b: 'S4; S5' treffer, men ikke f.eks. 'S50'
            er_s5 = df[KOL_STATUS].str.contains(rf"\b{_STATUS_S5}\b", regex=True)
            ark_data[_ARK_S5] = df[er_s5 & ~i_kontrakt]
        else:
            print(f"  [INFO] Arket '{_ARK_S5}' lages ikke: krever CSV med status.")
    else:
        print(f"  [INFO] Arkene '{_ARK_SVAR}' og '{_ARK_S5}' lages ikke: "
              f"krever kontrakt-Excel.")

    skriv_excel(ark_data, ut_sti)
    print(f"\nFerdig: {len(df)} rader -> {ut_sti}")
    for arknavn, ark_df in ark_data.items():
        print(f"  Ark '{arknavn}': {len(ark_df)} rader")
    print(f"  Unike mengdelisteposter: {df[KOL_POST].nunique()}")
    # contains, ikke ==: kommentaren kan ogsaa inneholde f.eks. '; RB 2'
    n_uten_1 = int(df[KOL_KOMMENTAR].str.contains(_KOMM_MANGLER_1, regex=False).sum())
    if n_uten_1:
        print(f"  Poster som ikke starter paa '1': {n_uten_1}")
    if status_oppslag is not None:
        fordeling = df[KOL_STATUS].replace("", "(ingen)").value_counts()
        print(f"  Statusfordeling:")
        for status, antall in fordeling.items():
            print(f"    {status}: {antall}")
    return 0


# ===== Vindu (tkinter) =====

# Sist brukte filer i vinduet lagres her mellom kjoeringer
_INNSTILLINGER_FIL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "pdf_manglende_poster_innstillinger.json",
)


def _les_innstillinger() -> dict:
    try:
        with open(_INNSTILLINGER_FIL, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _lagre_innstillinger(data: dict) -> None:
    try:
        with open(_INNSTILLINGER_FIL, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # ikke kritisk om lagring feiler

class _KoeStroem:
    """Fil-lignende objekt som sender print-utskrift til en koe.

    tkinter-widgets kan bare roeres fra hovedtraaden, saa arbeidertraaden
    legger teksten i en koe som vinduet henter fra med after().
    """

    def __init__(self, koe):
        self.koe = koe

    def write(self, tekst):
        if tekst:
            self.koe.put(tekst)

    def flush(self):
        pass


def start_gui() -> None:
    """Aapne et enkelt vindu for aa velge filer og kjoere uttrekket."""
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox

    rot = tk.Tk()
    rot.title("Manglende poster: PDF til Excel")
    rot.geometry("780x640")
    rot.minsize(600, 480)

    # Sist brukte filer. Filer som er slettet/flyttet siden sist tas ikke med.
    # Utfila gjenopprettes ikke -- den foreslaas fra PDF-en, saa en gammel
    # utfil ikke blir overskrevet ved et uhell.
    lagret = _les_innstillinger()
    lagret_pdf = lagret.get("pdf", "")
    if not os.path.exists(lagret_pdf):
        lagret_pdf = ""

    pdf_var = tk.StringVar(value=lagret_pdf)
    ut_var = tk.StringVar(
        value=os.path.splitext(lagret_pdf)[0] + ".xlsx" if lagret_pdf else "")
    debug_var = tk.BooleanVar(value=bool(lagret.get("debug", False)))
    logg_koe = queue.Queue()

    ramme = tk.Frame(rot, padx=12, pady=12)
    ramme.pack(fill="both", expand=True)
    ramme.columnconfigure(1, weight=1)
    ramme.rowconfigure(6, weight=1)

    # --- Rad 0: PDF ---
    def velg_pdf():
        sti = filedialog.askopenfilename(
            title="Velg PDF-rapport",
            filetypes=[("PDF", "*.pdf"), ("Alle filer", "*.*")],
        )
        if sti:
            pdf_var.set(sti)
            # Foreslaa utfil ved siden av PDF-en hvis ingen er valgt
            if not ut_var.get():
                ut_var.set(os.path.splitext(sti)[0] + ".xlsx")

    tk.Label(ramme, text="PDF-rapport:").grid(row=0, column=0, sticky="w", pady=4)
    tk.Entry(ramme, textvariable=pdf_var).grid(row=0, column=1, sticky="ew", pady=4)
    tk.Button(ramme, text="Velg...", width=11, command=velg_pdf).grid(
        row=0, column=2, padx=(6, 0), pady=4)

    # --- Rad 1-2: fillister (CSV og RB) ---
    def lag_filliste(rad, tekst, tittel, filtyper):
        tk.Label(ramme, text=tekst, justify="left").grid(
            row=rad, column=0, sticky="nw", pady=4)
        liste = tk.Listbox(ramme, height=3, selectmode="extended")
        liste.grid(row=rad, column=1, sticky="ew", pady=4)
        knapper = tk.Frame(ramme)
        knapper.grid(row=rad, column=2, sticky="n", padx=(6, 0), pady=4)

        def legg_til():
            for sti in filedialog.askopenfilenames(title=tittel, filetypes=filtyper):
                if sti not in liste.get(0, "end"):
                    liste.insert("end", sti)

        def fjern():
            for idx in reversed(liste.curselection()):
                liste.delete(idx)

        tk.Button(knapper, text="Legg til...", width=11, command=legg_til).pack(fill="x")
        tk.Button(knapper, text="Fjern valgt", width=11, command=fjern).pack(
            fill="x", pady=(4, 0))
        return liste

    csv_liste = lag_filliste(
        1, "CSV-filer\n(FOB_Status):", "Velg CSV-filer",
        [("CSV", "*.csv"), ("Alle filer", "*.*")],
    )
    rb_liste = lag_filliste(
        2, "Kontrakt-Excel\n(RB, alle ark):", "Velg kontrakt-Excel (flere tillatt)",
        [("Excel", "*.xlsx *.xlsm *.xls"), ("Alle filer", "*.*")],
    )

    for liste, noekkel in ((csv_liste, "csv"), (rb_liste, "kontrakt")):
        for sti in lagret.get(noekkel, []):
            if isinstance(sti, str) and os.path.exists(sti):
                liste.insert("end", sti)

    def lagre_valg():
        _lagre_innstillinger({
            "pdf": pdf_var.get().strip(),
            "csv": list(csv_liste.get(0, "end")),
            "kontrakt": list(rb_liste.get(0, "end")),
            "debug": bool(debug_var.get()),
        })

    def ved_lukking():
        lagre_valg()
        rot.destroy()

    rot.protocol("WM_DELETE_WINDOW", ved_lukking)

    # --- Rad 3: utfil ---
    def velg_ut():
        naa = ut_var.get()
        sti = filedialog.asksaveasfilename(
            title="Lagre Excel som",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=os.path.dirname(naa) if naa else None,
            initialfile=os.path.basename(naa) if naa else "",
        )
        if sti:
            ut_var.set(sti)

    tk.Label(ramme, text="Lagre som:").grid(row=3, column=0, sticky="w", pady=4)
    tk.Entry(ramme, textvariable=ut_var).grid(row=3, column=1, sticky="ew", pady=4)
    tk.Button(ramme, text="Velg...", width=11, command=velg_ut).grid(
        row=3, column=2, padx=(6, 0), pady=4)

    # --- Rad 4: valg + kjoer ---
    valg = tk.Frame(ramme)
    valg.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 4))
    tk.Checkbutton(valg, text="Debug (vis hver rad som leses)",
                   variable=debug_var).pack(side="left")
    kjor_knapp = tk.Button(valg, text="Kjør", width=14, bg="#1B4AA3", fg="white",
                           activebackground="#155B7A", activeforeground="white")
    kjor_knapp.pack(side="right")

    # --- Rad 5-6: logg ---
    tk.Label(ramme, text="Logg:").grid(row=5, column=0, sticky="w", pady=(8, 2))
    logg_ramme = tk.Frame(ramme)
    logg_ramme.grid(row=6, column=0, columnspan=3, sticky="nsew")
    logg = tk.Text(logg_ramme, wrap="word", font=("Consolas", 9), state="disabled")
    rull = tk.Scrollbar(logg_ramme, command=logg.yview)
    logg.configure(yscrollcommand=rull.set)
    rull.pack(side="right", fill="y")
    logg.pack(side="left", fill="both", expand=True)

    def skriv_logg(tekst):
        logg.configure(state="normal")
        logg.insert("end", tekst)
        logg.see("end")
        logg.configure(state="disabled")

    def hent_fra_koe():
        try:
            while True:
                element = logg_koe.get_nowait()
                if isinstance(element, tuple):          # ("ferdig", kode, ut_sti)
                    _, kode, ut_sti = element
                    kjor_knapp.configure(state="normal", text="Kjør")
                    if kode == 0:
                        messagebox.showinfo("Ferdig", f"Excel lagret:\n{ut_sti}")
                    else:
                        messagebox.showerror("Feil", "Kjøringen feilet. Se loggen.")
                else:
                    skriv_logg(element)
        except queue.Empty:
            pass
        rot.after(100, hent_fra_koe)

    def start_kjoring():
        pdf = pdf_var.get().strip()
        if not pdf:
            messagebox.showwarning("Mangler PDF", "Velg en PDF-rapport først.")
            return

        lagre_valg()

        args = argparse.Namespace(
            pdf=pdf,
            output=ut_var.get().strip() or None,
            csv=list(csv_liste.get(0, "end")) or None,
            rb=list(rb_liste.get(0, "end")) or None,
            grense=None,
            debug=debug_var.get(),
        )

        logg.configure(state="normal")
        logg.delete("1.0", "end")
        logg.configure(state="disabled")
        kjor_knapp.configure(state="disabled", text="Kjører...")

        def arbeider():
            # Kun en kjoering om gangen (knappen er deaktivert), saa global
            # omdirigering av stdout er trygt her.
            with contextlib.redirect_stdout(_KoeStroem(logg_koe)):
                try:
                    kode = kjor(args)
                except Exception as feil:
                    print(f"\n[FEIL] {feil}")
                    kode = 1
            logg_koe.put(("ferdig", kode, args.output or "ved siden av PDF-en"))

        threading.Thread(target=arbeider, daemon=True).start()

    kjor_knapp.configure(command=start_kjoring)
    rot.after(100, hent_fra_koe)
    rot.mainloop()


def main() -> int:
    # Uten argumenter (f.eks. dobbeltklikk paa fila): aapne vinduet
    if len(sys.argv) == 1:
        start_gui()
        return 0
    return kjor(_lag_parser().parse_args())


if __name__ == "__main__":
    sys.exit(main())
