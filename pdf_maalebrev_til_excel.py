# -*- coding: utf-8 -*-
"""
Les alle APM-PDF-filer i en mappe og samle tabelldata i ett Excel-dokument.

Ark 1 - Forklaring og Sammendrag:
  Oversikt over ark, kolonner, farger og symboler, pluss nøkkeltall.

Ark 2 - APM sammenstilling:
  PNS | Ekstra inndeling | APM | Post | Mengde | Enhet | EO |
  Kommentar | Utgatt | Ny_APM

Ark 3 - Post sammenstilling:
  Aggregert pr. Post (kun aktive APMer). Én rad per post per APM,
  pluss [TOTAL]-linje for poster med flere APMer.

Ark 4 - Armering totalmengde:
  Aktive rader med enhet kg (uten stål-referanse).

Ark 5 - Armeringsmengder:
  APM | Post | Formkode | kg_i_modell | Faktor_omfar | Vekt_inkl_omfar

Regler:
  - Identiske filer (MD5) telles kun én gang.
  - Filer med '(N)'-suffiks markeres som utgått dersom grunnfilen finnes.
  - Hoveddatatabellen kan gå over flere PDF-sider.
  - Formkode-tabeller brukes til armeringsarket.
"""

import contextlib
import hashlib
import os
import re
import sys

import datetime

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# -- Kolonnenavn og output --
KOL_OMEGA = "Fra OMEGA: rapportert kummulativ mengde"
KOL_KONTRAKT = "Kontraktsmengde"
KOL_ENHETSPRIS = "Enhetspris"
KOL_PROSENT = "prosent: APM VS. Omega"
KOL_MANGLER = "Ikke i Omega"   # '!' naar posten ikke finnes i Omega i det hele tatt
KOL_OMEGA_ENHET = "Omega Enhet"  # Enhetskode fra Omega (til sammenligning med APM-enhet)
KOL_FORBRUK     = "Forbruk %"    # Fargekodet forbruksprosent (skjult tekst, kun farge)
KOL_DELMAAL     = "Sist oppdatert delmaalebrev"  # Maned+aar for siste oppdaterte delmalebrev (fra Omega)
KOL_SRES        = "SRES"           # Sann/Usann fra Omega SRES-kolonne
KOL_BOQA_KOMM   = "Kommentar fra forrige m\u00e5lebrevspakke"  # BOQA-kommentar fra Omega

KOLONNER_HOVED = [
    KOL_MANGLER, "Flere APMer", "PNS", "Ekstra inndeling", "APM", "Post",
    "Mengde fra APM", "Enhet fra APM", "EO", "Kommentar", "Utgatt", "Ny_APM", "Vedlegg",
    KOL_OMEGA, KOL_KONTRAKT, KOL_ENHETSPRIS, KOL_PROSENT, KOL_DELMAAL, KOL_SRES,
]
KOLONNER_ARM = [
    "APM", "Post", "Formkode", "kg_i_modell", "Faktor_omfar", "Vekt_inkl_omfar",
]
KOLONNER_SAMM = [
    "Post", "Ekstra inndeling", "APM", "Mengde fra APM", "Enhet fra APM", KOL_OMEGA_ENHET, "Kommentar",
    KOL_BOQA_KOMM,
    "Flere APMer", KOL_OMEGA, KOL_KONTRAKT, KOL_ENHETSPRIS, KOL_PROSENT, KOL_FORBRUK, "!", KOL_DELMAAL, KOL_SRES,
]
OUTPUT_FILNAVN = "Maalebrev_samlet.xlsx"

# -- Armerings-header-aliaser (PDF-kolonner -> kanoniske navn) --
_ARM_ALIAS = {
    "post": "Post", "formkode": "Formkode",
    "kg i modell": "kg_i_modell", "kg_i_modell": "kg_i_modell",
    "faktor for omfar": "Faktor_omfar", "faktor_omfar": "Faktor_omfar",
    "faktor omfar": "Faktor_omfar",
    "vekt inkl omfar": "Vekt_inkl_omfar", "vekt_inkl_omfar": "Vekt_inkl_omfar",
    "vekt inkl. omfar": "Vekt_inkl_omfar",
}

# -- Regex (ferdigkompilert) --
_RE_ERSTATTER = re.compile(r"ERSTATTER\s+(APM\s*\d+)", re.IGNORECASE)
_RE_APM_NR = re.compile(r"APM(\d+)", re.IGNORECASE)
_RE_VERSJON = re.compile(r"^(.+?) \((\d+)\)$")
_RE_WHITESPACE = re.compile(r"\s")
_RE_MELLOMROM_I_TALL = re.compile(r"(?<=\d)\s+(?=\d)")
_RE_PROSESSKODE_GYLDIG = re.compile(r"^\d+(?:[.\-]\d+)+$")
_RE_PROSESSKODE_SOEK = re.compile(r"\d{3,4}(?:[.\-]\d+){2,}")
# Gyldige starters for foerste segment i en prosesskode
_GYLDIGE_PROSESSKODE_START = ("1302", "1303", "1407", "01")

# -- Excel-stiler --
_HDR_FILL = PatternFill("solid", fgColor="1B4AA3")
_HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
_VTOPP = Alignment(vertical="top")
_FILL_RAD = PatternFill("solid", fgColor="EEF2FA")
_FILL_UTGAATT = PatternFill("solid", fgColor="FFE0E0")
_FILL_OMEGA_ADV = PatternFill("solid", fgColor="FF9999")
_FILL_ARM = PatternFill("solid", fgColor="E8F5E9")
_FILL_ARM_HDR = PatternFill("solid", fgColor="2E7D32")
_FILL_TOT_HDR = PatternFill("solid", fgColor="5D4037")
_FILL_TOT_RAD = PatternFill("solid", fgColor="EFEBE9")
_FILL_MANGLER_OMEGA = PatternFill("solid", fgColor="FFF3CD")  # gul: post ikke funnet i Omega
_FILL_MANGL_HDR     = PatternFill("solid", fgColor="7B341E")  # mørk rødoransje: header for mangler-ark
_FILL_SAMM_HDR  = PatternFill("solid", fgColor="155B7A")       # moerkebla header for Post sammenstilling
_FILL_SAMM_RAD  = PatternFill("solid", fgColor="FAFBFC")       # varm off-white: alle normale rader
_FILL_SAMM_OVER = PatternFill("solid", fgColor="FF9999")       # rod: Omega > totalmengde (prosent > 100 %)
_FILL_SAMM_TOT  = PatternFill("solid", fgColor="2E6B8A")       # moerk staalblaa: total-rad, hvit tekst
_FILL_SAMM_FLERE = PatternFill("solid", fgColor="C8DCF0")      # mellomblaa: individuelle rader paa poster med flere APMer
_FONT_SAMM_FLERE = Font(italic=True, color="505050", name="Calibri", size=10)  # kursiv moerk graa
_FILL_SAMM_ENHET_FEIL = PatternFill("solid", fgColor="E1BEE7")  # lys lilla: enhet i PDF stemmer ikke med Omega-enhet
_FILL_SAMM_100        = PatternFill("solid", fgColor="92D050")  # lys grønn: prosent er nøyaktig 100 %
_FILL_PROSENT_GROENN  = PatternFill("solid", fgColor="C6EFCE")  # lys grønn: forbruk < 80 %
_FILL_PROSENT_GUL     = PatternFill("solid", fgColor="FFEB9C")  # lys gul:    forbruk 80-100 %
_FILL_PROSENT_ROD     = PatternFill("solid", fgColor="FFC7CE")  # lys rød:    forbruk > 100 %

# Standarfonter for dataceller (alle Calibri 10 for konsistens)
_FONT_DATA     = Font(name="Calibri", size=10)
_FONT_TOT      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)   # hvit bold - total-rader med mørk bakgrunn
_FONT_TOT_ARM  = Font(bold=True, name="Calibri", size=10)                   # mørk bold - total-rader Armering totalmengde
_FONT_100      = Font(color="14451A", name="Calibri", size=10)               # mørk grønn - nøyaktig 100 %
_FONT_100_BOLD = Font(bold=True, color="14451A", name="Calibri", size=10)    # mørk grønn bold - total-rad 100 %
_KANT_TYKK     = Side(style="medium", color="155B7A")                        # tykk blå skillelinje

# Header-kommentarer (tooltips) for kolonnene
_HDR_KOMMENTARER = {
    KOL_MANGLER:     "Utropstegn (!) hvis posten ikke finnes i Omega overhodet.",
    "Flere APMer":   "Checkmark (\u2713) hvis samme prosesskode finnes i mer enn én APM.",
    KOL_OMEGA:       "Rapportert kumulativ mengde fra Omega BOQNT-arket.",
    KOL_KONTRAKT:    "Kontraktsmengde fra Omega BOQNT-arket.",
    KOL_ENHETSPRIS:  "Enhetspris fra Omega BOQNT-arket.",
    KOL_PROSENT:     "Prosent: Omega rapportert mengde / APM-mengde * 100.",
    KOL_FORBRUK:     "Fargekodet forbruksprosent: grønn < 80 %, gul 80-100 %, rød > 100 %.",
    KOL_OMEGA_ENHET: "Enhetskode fra Omega. Lilla bakgrunn = avvik fra APM-enhet.",
    "Vekt_inkl_omfar": "Beregnet vekt inkl. omfar-faktor (kg). Beregnes automatisk dersom PDF-kolonnen er tom.",
}

# Roterende fargepalett for APM-rader i Malebrev-arket.
# Unngaar roed/rosa (brukes for utgaatt/Omega-advarsel).
_APM_FARGER = [
    PatternFill("solid", fgColor="D6E4F7"),  # lys blaa
    PatternFill("solid", fgColor="D5F0DC"),  # lys groenn
    PatternFill("solid", fgColor="FFF9C4"),  # lys gul
    PatternFill("solid", fgColor="FDEBD0"),  # lys oransje
    PatternFill("solid", fgColor="D0F0F0"),  # lys turkis
    PatternFill("solid", fgColor="E8DAEF"),  # lys lilla
    PatternFill("solid", fgColor="F5E6C8"),  # lys brun/tan
    PatternFill("solid", fgColor="DCF8E8"),  # lys mynte
]

BREDDER_HOVED = {
    KOL_MANGLER: 14, "Flere APMer": 12, "PNS": 18, "Ekstra inndeling": 20, "APM": 40, "Post": 26,
    "Mengde fra APM": 12, "Enhet fra APM": 10, "EO": 8, "Kommentar": 45,
    "Utgatt": 10, "Ny_APM": 40, "Vedlegg": 30,
    KOL_OMEGA: 22, KOL_KONTRAKT: 16, KOL_ENHETSPRIS: 14, KOL_PROSENT: 20,
    KOL_DELMAAL: 20, KOL_SRES: 10,
}
BREDDER_ARM = {
    "APM": 40, "Post": 26, "Formkode": 14,
    "kg_i_modell": 14, "Faktor_omfar": 14, "Vekt_inkl_omfar": 18,
}
BREDDER_SAMM = {
    "Post": 26, "Ekstra inndeling": 20, "APM": 40,
    "Mengde fra APM": 14, "Enhet fra APM": 10, KOL_OMEGA_ENHET: 12, "Kommentar": 45,
    KOL_BOQA_KOMM: 45,
    "Flere APMer": 12, KOL_OMEGA: 22, KOL_KONTRAKT: 16, KOL_ENHETSPRIS: 14,
    KOL_PROSENT: 20, KOL_FORBRUK: 14, "!": 4, KOL_DELMAAL: 20, KOL_SRES: 10,
}
KOLONNER_OMEGA_UTEN_APM = [
    "Prosesskode", KOL_OMEGA, KOL_KONTRAKT, KOL_ENHETSPRIS, "Enhet", KOL_DELMAAL, KOL_SRES,
]
BREDDER_OMEGA_UTEN_APM = {
    "Prosesskode": 26, KOL_OMEGA: 22, KOL_KONTRAKT: 16, KOL_ENHETSPRIS: 14,
    "Enhet": 10, KOL_DELMAAL: 20, KOL_SRES: 10,
}
KOLONNER_MANGLER_SAMM = [
    "Prosesskode", KOL_OMEGA, KOL_KONTRAKT, KOL_ENHETSPRIS, "Enhet", KOL_DELMAAL, KOL_SRES, "Merknad", KOL_BOQA_KOMM,
]
BREDDER_MANGLER_SAMM = {
    "Prosesskode": 26, KOL_OMEGA: 22, KOL_KONTRAKT: 16, KOL_ENHETSPRIS: 14,
    "Enhet": 10, KOL_DELMAAL: 20, KOL_SRES: 10, "Merknad": 34, KOL_BOQA_KOMM: 45,
}


# ===== Hjelpefunksjoner =====

def _rens(verdi) -> str:
    """Rens celleverdi: fjern whitespace og linjeskift."""
    return str(verdi).strip().replace("\n", " ") if verdi is not None else ""


def _normaliser_enhet(enhet: str) -> str:
    """Normaliser enhetskode for sammenligning.

    Haandterer vanlige varianter som dukker opp i PDF vs. Omega:
      m2 <-> m²,  m3 <-> m³,  lm <-> LM  osv.
    Returnerer lowercase normalisert streng.
    """
    e = str(enhet or "").strip().lower()
    e = e.replace("\u00b2", "2")   # m² -> m2
    e = e.replace("\u00b3", "3")   # m³ -> m3
    e = e.replace("\u00b9", "1")   # superscript 1
    if e == "st":                  # st -> stk
        e = "stk"
    return e


def _normaliser_prosesskode(tekst: str, advar_fn=None) -> str:
    """Rens og valider prosesskode.

    Noen PDF-er gir forurensede verdier som '_1A30P3M.0811.06.241.1.1' i stedet
    for '1303.01.06.241.1.1' pga. fontkodings-problemer i pdfplumber.
    Andre ganger faar foerste segment et ekstra ledende siffer, f.eks.
    '11302-03.81.241' i stedet for '1302-03.81.241'.
    Bindestrek er kun tillatt som foerste skilletegn etter de 4 foerste sifferne;
    alle andre bindestreker erstattes med punktum.
    Returnerer tom streng hvis ingenting gyldig funnes.
    """
    if not tekst:
        return ""
    renset = tekst.strip()
    # Allerede gyldig format: kun siffer, punktum og bindestrek
    if _RE_PROSESSKODE_GYLDIG.match(renset):
        kode = _trim_prosesskode_start(renset, advar_fn)
        return _fix_separatorer(kode, advar_fn)
    # Forsoek aa trekke ut gyldig moenster inni den forurensede teksten
    treff = _RE_PROSESSKODE_SOEK.search(renset)
    if treff:
        funnet = _fix_separatorer(_trim_prosesskode_start(treff.group(), advar_fn), advar_fn)
        if advar_fn:
            advar_fn(f"[WARNING] Forurensed prosesskode '{renset}' -> bruker '{funnet}'")
        return funnet
    # Siste utvei: PDF-tekst med tegn fra APM-navn fletta inn i prosesskoden
    # (skjer naar APM-kolonnens lange tekst overlapper Post-kolonnens x-posisjon i PDF).
    # Strip alle tegn som ikke er siffer, punktum eller bindestrek, og valider resultatet.
    stripped = re.sub(r"[^\d.\-]", "", renset)
    stripped = re.sub(r"\.{2,}", ".", stripped).strip(".")
    if _RE_PROSESSKODE_GYLDIG.match(stripped):
        kode = _trim_prosesskode_start(stripped, advar_fn)
        kode = _fix_separatorer(kode, advar_fn)
        if kode and any(kode.startswith(p) for p in _GYLDIGE_PROSESSKODE_START) and len(re.split(r"[.\-]", kode)) >= 4:
            if advar_fn:
                advar_fn(f"[WARNING] Tegnfletting i prosesskode '{renset}' -> '{kode}'")
            return kode
    return ""


def _apm_prefix_segmenter(apm_navn: str) -> list:
    """Trekk ut ledende numeriske segmenter fra APM-navn.

    Eksempel: '1303_01_LYA' -> ['1303', '01']
              'APM65 (1)'   -> []
    Splitter paa understrek, bindestrek og mellomrom; stopper ved foerste
    ikke-rent-numerisk segment.
    """
    deler = re.split(r"[_\-\s]+", apm_navn.strip())
    prefix = []
    for del_ in deler:
        if re.match(r"^\d+$", del_):
            prefix.append(del_)
        else:
            break
    return prefix


def _korriger_post_med_apm_prefix(post: str, apm_navn: str, omega_poster: set) -> str:
    """Forsok aa rekonstruere prosesskode ved aa bruke numerisk prefix fra APM-navn.

    Eksempel:
      post='0811.06.241.1.1', apm='1303_01_LYA'
      Prover: '1303.01.06.241.1.1' -> om funnet i omega_poster returneres dette.

    Returnerer ny prosesskode hvis funnet i omega_poster, ellers tom streng.
    """
    apm_prefix = _apm_prefix_segmenter(apm_navn)
    if not apm_prefix:
        return ""

    post_segmenter = re.split(r"[.\-]", post)
    # Prover aa erstatte de foerste 1..N segmentene med apm_prefix
    for n_replace in range(1, len(post_segmenter)):
        rester = post_segmenter[n_replace:]
        if not rester:
            continue
        ny_post = _normaliser_prosesskode(".".join(apm_prefix + rester))
        if ny_post and ny_post in omega_poster:
            return ny_post
    return ""


def _dekorrupt_via_omega(
    prosesskode: str,
    pns: str,
    enhet: str,
    omega_data: dict,
    advar_fn=None,
) -> str:
    """Finn riktig prosesskode ved aa trunkere et tegnfletting-resultat mot Omega.

    Bakgrunn: naar APM-kolonnens lange tekst fletter seg inn i Post-kolonnen,
    gir tegnfletting-rensingen en kode med KORREKTE tidlige segmenter men
    EKSTRA segmenter paa slutten fra APM-filnavnets siffer.
    Eks: tegnfletting gir '1303.02.05.251.250.15' i stedet for '1303.02.05.251.2.1'.
    De foerste fire segmentene stemmer; '250.15' er forurensning fra 'APM205'-sifre.

    Strategi (segment-trunkering + subsequence):
    1. Fjern ett segment om gangen fra slutten av tegnfletting-koden.
    2. For hvert prefix, finn alle Omega-koder som starter med prefikset.
    3. Begrens til PNS-prefix (foerste 2 segmenter fra PNS-kolonnen).
    4. Filtrer paa enhetskode.
    5. Subsequence-test: Omega-kodens siste segment-siffer maa vaere en ordnet
       delmengde av sifrene i de fjernede tegnfletting-segmentene.
       (f.eks. drop '250.15' siffer='25015', kand. '2.1' siffer='21' -> subseq. av '25015' -> JA)
    6. Returnerer koden hvis entydig, ellers prover kortere prefix.
    """
    if not omega_data or not prosesskode:
        return ""

    # PNS-prefix som ekstra guard (f.eks. '1303.02')
    pns_norm = re.sub(r"[^\d.]", "", (pns or "").strip()).strip(".")
    pns_seg = [s for s in pns_norm.split(".") if s]
    pns_prefix = ".".join(pns_seg[:2]) if len(pns_seg) >= 2 else ""

    enhet_norm = _normaliser_enhet(enhet) if enhet else ""
    segmenter = re.split(r"[.\-]", prosesskode)

    def er_sub(sub_str: str, full_str: str) -> bool:
        it = iter(full_str)
        return all(c in it for c in sub_str)

    # Proev aa trunkere fra slutten (behold minst 3 segmenter)
    for n_drop in range(1, len(segmenter) - 2):
        prefix_segs = segmenter[:len(segmenter) - n_drop]
        fjernede_segs = segmenter[len(segmenter) - n_drop:]
        soeke_prefix = ".".join(prefix_segs) + "."
        fjernede_siffer = re.sub(r"[^\d]", "", "".join(fjernede_segs))

        kandidater = [k for k in omega_data if k.startswith(soeke_prefix)]
        if not kandidater:
            continue

        # Begrens til PNS-prefix
        if pns_prefix:
            kandidater = [k for k in kandidater if k.startswith(pns_prefix + ".")]
        if not kandidater:
            continue

        # Filtrer paa enhet
        if enhet_norm:
            med_enhet = [
                k for k in kandidater
                if _normaliser_enhet(omega_data[k].get("enhet") or "") == enhet_norm
            ]
            if med_enhet:
                kandidater = med_enhet

        if len(kandidater) == 1:
            kode = kandidater[0]
            if advar_fn:
                advar_fn(
                    f"[WARNING] Omega-dekorrupsjon (trunkering -{n_drop}, entydig): "
                    f"'{prosesskode}' -> '{kode}'"
                )
            return kode

        # Subsequence-test naar flere kandidater gjenstaar:
        # Omega-kodens segmenter etter prefikset maa vaere en ordnet delmengde
        # av sifrene i de fjernede tegnfletting-segmentene.
        if fjernede_siffer:
            n_prefix = len(prefix_segs)
            matchende = []
            for k in kandidater:
                k_segs = re.split(r"[.\-]", k)
                k_siste = k_segs[n_prefix:]
                k_siste_siffer = re.sub(r"[^\d]", "", "".join(k_siste))
                if k_siste_siffer and er_sub(k_siste_siffer, fjernede_siffer):
                    matchende.append(k)
            if len(matchende) == 1:
                kode = matchende[0]
                if advar_fn:
                    advar_fn(
                        f"[WARNING] Omega-dekorrupsjon (trunkering -{n_drop}, subseq): "
                        f"'{prosesskode}' -> '{kode}'"
                    )
                return kode

    return ""


def _fix_gruppe_konsistens(rader: list, omega_data: dict, omega_poster: set, advar_fn=None) -> None:
    """Bruk gruppe-konsistens for aa fikse gjenvaerende korrupte prosesskoder.

    Rader som deler samme base (foerste 4 segmenter av _raa_post) og samme
    korrupte 5. segment tilhorer trolig samme prosesskode-gruppe.

    Strategi (i prioritert rekkefølge per gruppe):
    1. Soesken-modus: ett gruppemedlem er allerede loest via _dekorrupt_via_omega.
       Utled APM-bidragssifferene i siste segment fra soeskenets korrupte vs. korrekte kode.
       Bruk dem til direkte rekonstruksjon av Y for alle uloeste rader.
    2. Global-APM-modus: ingen loest soesken, men APM-monsteret er etablert globalt fra
       andre loeste rader i samme PDF.  Beregn X ved aa fjerne globalt APM-bidrag fra
       korrupt 5. segment, og Y fra korrupt 6. segment. Verifiser mot Omega + enhet.
    3. Omega-modus: ingen loest soesken og ingen global APM-kunnskap.
       Prøv alle X-verdier fra Omega.  Velg X hvis noyaktig ett treffer alle rader
       (enhet + subsequence av korrupte siste segmenter).
    """
    from collections import defaultdict

    def er_sub(s: str, f: str) -> bool:
        """True hvis s er en ordnet delmengde av f."""
        it = iter(f)
        return all(c in it for c in s)

    def fjern_subseq_tegn(corrupt: str, apm: str) -> str:
        """Fjern apm-tegnene (som ordnet delmengde) fra corrupt; returner resten (reelle siffer)."""
        resultat = list(corrupt)
        pos = 0
        for c in apm:
            while pos < len(resultat) and resultat[pos] != c:
                pos += 1
            if pos < len(resultat):
                resultat[pos] = None
                pos += 1
        return "".join(ch for ch in resultat if ch is not None)

    # Steg A: Beregn globalt APM-moenster fra alle loeste (korrupt -> korrekt) rader.
    # Brukes i gruppe-modus 2 for grupper uten in-gruppe soesken.
    global_apm_5_sett = set()
    global_apm_6_sett = set()
    for rad in rader:
        raa = rad.get("_raa_post", "")
        post = rad.get("Post", "")
        if not raa or raa == post:
            continue
        raa_s  = re.split(r"[.\-]", raa)
        post_s = re.split(r"[.\-]", post)
        if len(raa_s) == 6 and len(post_s) == 6:
            global_apm_5_sett.add(fjern_subseq_tegn(raa_s[4], post_s[4]))
            global_apm_6_sett.add(fjern_subseq_tegn(raa_s[5], post_s[5]))
    global_apm_5 = list(global_apm_5_sett)[0] if len(global_apm_5_sett) == 1 else None
    global_apm_6 = list(global_apm_6_sett)[0] if len(global_apm_6_sett) == 1 else None

    # Steg B: Grupper rader paa (base4, korrupt_5te_segment)
    grupper = defaultdict(list)
    for i, rad in enumerate(rader):
        raa = rad.get("_raa_post", "")
        if not raa:
            continue
        segs = re.split(r"[.\-]", raa)
        if len(segs) == 6:
            key = (".".join(segs[:4]), segs[4])
            grupper[key].append(i)

    for key, idx_liste in grupper.items():
        if len(idx_liste) < 2:
            continue

        base_4 = key[0]
        corrupt_5 = key[1]

        # --- Modus 1: soesken-modus ---
        soesken_map = {}
        for i in idx_liste:
            post = rader[i]["Post"]
            if post in omega_poster:
                post_segs = re.split(r"[.\-]", post)
                if len(post_segs) >= 5:
                    soesken_map[post_segs[4]] = i

        if len(soesken_map) == 1:
            x_val = list(soesken_map.keys())[0]
            sib_i = soesken_map[x_val]
            sib_raa_segs  = re.split(r"[.\-]", rader[sib_i]["_raa_post"])
            sib_post_segs = re.split(r"[.\-]", rader[sib_i]["Post"])
            apm_i_siste = fjern_subseq_tegn(sib_raa_segs[-1], sib_post_segs[-1])

            assignments = {}
            ok = True
            for i in idx_liste:
                if rader[i]["Post"] in omega_poster:
                    continue
                raa_segs = re.split(r"[.\-]", rader[i]["_raa_post"])
                enhet_norm = _normaliser_enhet(rader[i].get("Enhet fra APM", ""))
                real_y = fjern_subseq_tegn(raa_segs[-1], apm_i_siste) if apm_i_siste else raa_segs[-1]
                if not real_y:
                    ok = False; break
                kode_kand = f"{base_4}.{x_val}.{real_y}"
                if kode_kand not in omega_data:
                    ok = False; break
                kode_enhet = _normaliser_enhet(omega_data[kode_kand].get("enhet") or "")
                if enhet_norm and kode_enhet != enhet_norm:
                    ok = False; break
                assignments[i] = kode_kand

            if ok and assignments:
                for i, kode in assignments.items():
                    if advar_fn:
                        advar_fn(
                            f"[WARNING] Omega-gruppe-dekorrupsjon (soesken, X={x_val}): "
                            f"'{rader[i]['Post']}' -> '{kode}'"
                        )
                    rader[i]["Post"] = kode
            continue

        # --- Modus 2: global-APM-modus ---
        if global_apm_5 is not None and global_apm_6 is not None:
            real_x = fjern_subseq_tegn(corrupt_5, global_apm_5)
            if real_x:
                assignments = {}
                ok = True
                for i in idx_liste:
                    if rader[i]["Post"] in omega_poster:
                        continue
                    raa_segs = re.split(r"[.\-]", rader[i]["_raa_post"])
                    enhet_norm = _normaliser_enhet(rader[i].get("Enhet fra APM", ""))
                    real_y = fjern_subseq_tegn(raa_segs[-1], global_apm_6) if global_apm_6 else raa_segs[-1]
                    if not real_y:
                        ok = False; break
                    kode_kand = f"{base_4}.{real_x}.{real_y}"
                    if kode_kand not in omega_data:
                        ok = False; break
                    kode_enhet = _normaliser_enhet(omega_data[kode_kand].get("enhet") or "")
                    if enhet_norm and kode_enhet != enhet_norm:
                        ok = False; break
                    assignments[i] = kode_kand

                if ok and assignments:
                    for i, kode in assignments.items():
                        if advar_fn:
                            advar_fn(
                                f"[WARNING] Omega-gruppe-dekorrupsjon (global-APM, X={real_x}): "
                                f"'{rader[i]['Post']}' -> '{kode}'"
                            )
                        rader[i]["Post"] = kode
                    continue

        # --- Modus 3: omega-modus ---
        soek = base_4 + "."
        x_kand_liste = list({
            re.split(r"[.\-]", k)[4]
            for k in omega_data
            if k.startswith(soek) and len(re.split(r"[.\-]", k)) >= 5
        })
        n_prefix = len(re.split(r"\.", base_4)) + 1

        def test_x_omega(x_v):
            res = {}
            for i in idx_liste:
                if rader[i]["Post"] in omega_poster:
                    continue
                raa_segs = re.split(r"[.\-]", rader[i]["_raa_post"])
                fjernede_siffer = re.sub(r"[^\d]", "", "".join(raa_segs[4:]))
                enhet_norm = _normaliser_enhet(rader[i].get("Enhet fra APM", ""))
                soek_x = f"{base_4}.{x_v}."
                kand = [k for k in omega_data if k.startswith(soek_x)]
                if enhet_norm:
                    m_e = [k for k in kand
                           if _normaliser_enhet(omega_data[k].get("enhet") or "") == enhet_norm]
                    if m_e:
                        kand = m_e
                matchende = []
                for k in kand:
                    k_segs = re.split(r"[.\-]", k)
                    y_siffer = re.sub(r"[^\d]", "", "".join(k_segs[n_prefix:]))
                    if y_siffer and er_sub(y_siffer, fjernede_siffer):
                        matchende.append(k)
                if len(matchende) == 1:
                    res[i] = matchende[0]
                else:
                    return None
            return res if res else None

        plausible = []
        for x in x_kand_liste:
            r = test_x_omega(x)
            if r is not None:
                plausible.append((x, r))

        if len(plausible) == 1:
            x_val, assignments = plausible[0]
            for i, kode in assignments.items():
                if advar_fn:
                    advar_fn(
                        f"[WARNING] Omega-gruppe-dekorrupsjon (omega, X={x_val}): "
                        f"'{rader[i]['Post']}' -> '{kode}'"
                    )
                rader[i]["Post"] = kode


def _trim_prosesskode_start(kode: str, advar_fn=None) -> str:
    """Fjern ekstra ledende siffer fra foerste segment hvis det slutter paa en kjent gyldig start.

    Eksempel: '11302-03.81.241' -> '1302-03.81.241'
    """
    # Finn foerste skilletegn for aa isolere foerste segment
    foerste_sep = next((i for i, c in enumerate(kode) if c in ".-"), len(kode))
    segment = kode[:foerste_sep]
    rest = kode[foerste_sep:]
    for prefix in _GYLDIGE_PROSESSKODE_START:
        if segment == prefix:
            return kode  # allerede korrekt
        if len(segment) > len(prefix) and segment.endswith(prefix):
            ny = prefix + rest
            if advar_fn:
                advar_fn(f"[WARNING] Feil prosesskode-start '{kode}' -> '{ny}'")
            return ny
    return kode


def _normaliser_andre_segment(kode: str, omega_poster: set, advar_fn=None) -> str:
    """Fjern trailing null fra andre segment naar det gir treff i Omega.

    Eksempel: '1302-030.83.7612' -> '1302-03.83.7612'
    (PDF-en printer av og til '030' der Omega bruker '03'.)
    """
    sep_idx = next((i for i, c in enumerate(kode) if c in ".-"), -1)
    if sep_idx == -1:
        return kode
    etter = kode[sep_idx + 1:]
    neste_sep = next((i for i, c in enumerate(etter) if c in ".-"), len(etter))
    andre_seg = etter[:neste_sep]
    resten = etter[neste_sep:]
    if andre_seg.endswith("0") and len(andre_seg) >= 2:
        kandidat = kode[:sep_idx] + kode[sep_idx] + andre_seg[:-1] + resten
        if kandidat in omega_poster:
            if advar_fn:
                advar_fn(f"[WARNING] Andre-segment normalisert '{kode}' -> '{kandidat}'")
            return kandidat
    return kode


def _normaliser_foerste_separator(kode: str, omega_poster: set, advar_fn=None) -> str:
    """Erstatt foerste '.' med '-' naar det gir treff i Omega.

    Eksempel: '1303.03.82.191' -> '1303-03.82.191'
    (Noen PDF-er bruker punktum der Omega bruker bindestrek som foerste skilletegn.)
    Gjelder kun naar koden begynner med 4 siffer etterfulgt av '.'.
    """
    import re as _re
    if _re.match(r'^\d{4}\.', kode):
        kandidat = kode[:4] + '-' + kode[5:]
        if kandidat in omega_poster:
            if advar_fn:
                advar_fn(f"[WARNING] Foerste separator normalisert '{kode}' -> '{kandidat}'")
            return kandidat
    return kode


def _fix_separatorer(kode: str, advar_fn=None) -> str:
    """Sikre at bindestrek kun forekommer som foerste skilletegn etter de 4 foerste sifferne.

    Eksempel: '1303-10-43.121' -> '1303-10.43.121'
    Eksempel: '1302.01-83.7613' -> '1302.01.83.7613'
    """
    # Finn posisjon og type foerste skilletegn
    foerste_sep_idx = next((i for i, c in enumerate(kode) if c in ".-"), -1)
    if foerste_sep_idx == -1 or "-" not in kode:
        return kode  # ingen bindestreker -- ingenting aa fikse

    if kode[foerste_sep_idx] == "-" and foerste_sep_idx == 4:
        # Bindestrek etter 4 siffer er lovlig -- erstatt alle ETTERFØLGENDE bindestreker med punktum
        prefix = kode[:foerste_sep_idx + 1]        # f.eks. '1303-'
        resten = kode[foerste_sep_idx + 1:].replace("-", ".")
        ny = prefix + resten
    else:
        # Bindestrek paa feil posisjon -- erstatt alle bindestreker med punktum
        ny = kode.replace("-", ".")

    if ny != kode:
        if advar_fn:
            advar_fn(f"[WARNING] Feil separatorer i prosesskode '{kode}' -> '{ny}'")
    return ny


def _norsk_til_float(tekst: str):
    """Konverter norsk tallstreng til float.

    - Komma = desimaltegn (1.234,56 -> 1234.56)
    - Flere punktum = tusenskilletegn (1.234.567 -> 1234567)
    - Enkelt punktum = desimalpunkt (689.552 -> 689.552)
    """
    tekst = _RE_WHITESPACE.sub("", tekst).replace("'", "").replace('"', "")
    if not tekst:
        return None
    if "," in tekst:
        tekst = tekst.replace(".", "").replace(",", ".")
    elif tekst.count(".") > 1:
        tekst = tekst.replace(".", "")
    try:
        return float(tekst)
    except ValueError:
        return None


def _rens_mengde(verdi):
    """Returner float hvis mulig, ellers renset streng, eller None."""
    renset = _RE_MELLOMROM_I_TALL.sub("", _rens(verdi))
    if not renset:
        return None
    return _norsk_til_float(renset) or renset


def _parse_tall(verdi):
    """Parse celleverdi til float, eller None."""
    renset = _rens(verdi)
    return _norsk_til_float(renset) if renset else None


def _md5_fil(filsti: str) -> str:
    """Beregn MD5-hash for en fil."""
    hasher = hashlib.md5()
    with open(filsti, "rb") as fil:
        for blokk in iter(lambda: fil.read(65536), b""):
            hasher.update(blokk)
    return hasher.hexdigest()


def _ny_border():
    """Lag fersk Border (deling av Border-objekter gir hash-rekursjon i openpyxl)."""
    kant = Side(style="thin", color="CCCCCC")
    return Border(left=kant, right=kant, top=kant, bottom=kant)


def _ny_border_samm():
    """Border for Post sammenstilling: tynne vertikale kanter, tydelig bunnlinje som skiller radene."""
    side_v = Side(style="thin", color="D8E4EE")   # dempeblaa vertikal linje
    bunn   = Side(style="thin", color="8BAEC7")   # litt tydeligere bunnlinje
    ingen  = Side(style=None)
    return Border(left=side_v, right=side_v, top=ingen, bottom=bunn)


def _hent_tekst_fra_celle_via_chars(side, x0: float, top: float, x1: float, bottom: float) -> str:
    """Rekonstruer celletekst ved aa filtrere sidens tegn-liste til en gitt bbox.

    Brukes som fallback for Post-kolonnen naar pdfplumber's extract_tables()
    fletter tegn fra nabokolonner (f.eks. lang APM-kolonne-tekst) inn i cellen.
    Tegn sorteres venstre-til-hoyre (x0) innenfor samme linje.
    """
    tegn = sorted(
        [c for c in (side.chars or [])
         if x0 <= c["x0"] < x1 and top <= c["top"] < bottom],
        key=lambda c: c["x0"],
    )
    return "".join(c["text"] for c in tegn).strip()


# ===== Fil-analyse (duplikater, versjoner, vedlegg) =====

def _analyser_filer(mappe: str) -> tuple:
    """Returner (hopp_over, utgaatt_map, alle_stier).

    - hopp_over:   set med filnavn som er identiske kopier.
    - utgaatt_map: dict {gammelt_navn -> nytt_navn} for eldre versjoner.
    - alle_stier:  dict {navn_uten_ext -> full_sti}.
    """
    pdf_filer = sorted(fil for fil in os.listdir(mappe) if fil.lower().endswith(".pdf"))
    alle_stier = {os.path.splitext(fil)[0]: os.path.join(mappe, fil) for fil in pdf_filer}

    # Grupper filer etter MD5-hash for aa finne identiske kopier
    hash_grupper: dict[str, list] = {}
    for filnavn in pdf_filer:
        filhash = _md5_fil(os.path.join(mappe, filnavn))
        hash_grupper.setdefault(filhash, []).append(filnavn)

    hopp_over: set[str] = set()
    for duplikater in hash_grupper.values():
        if len(duplikater) <= 1:
            continue
        duplikater.sort(key=lambda navn: (
            (1, int(treff.group(1))) if (treff := re.search(r"\((\d+)\)", navn)) else (0, 0)
        ))
        for kopi in duplikater[1:]:
            hopp_over.add(kopi)
            print(f"    [IDENTISK] Hopper over {kopi} (lik {duplikater[0]})")

    # Finn utgaatte versjoner: "APM65 (1)" er utgaatt hvis "APM65" finnes
    aktive_navn = {os.path.splitext(fil)[0] for fil in pdf_filer if fil not in hopp_over}
    utgaatt_map: dict[str, str] = {}
    for navn in aktive_navn:
        match = _RE_VERSJON.match(navn)
        if not match:
            continue
        base = match.group(1)
        if base in aktive_navn:
            utgaatt_map[navn] = base
        else:
            moenster = re.compile(r"^" + re.escape(base) + r" \((\d+)\)$")
            varianter = sorted(
                [n for n in aktive_navn if moenster.match(n)],
                key=lambda x: int(moenster.match(x).group(1)),
            )
            for eldre in varianter[:-1]:
                utgaatt_map[eldre] = varianter[-1]

    return hopp_over, utgaatt_map, alle_stier


def _finn_vedlegg(mappe: str) -> dict:
    """Returner {basename -> kommaseparert liste} for ikke-PDF-filer med samme base som en PDF."""
    alle_filer = os.listdir(mappe)
    pdf_baser = {os.path.splitext(fil)[0] for fil in alle_filer if fil.lower().endswith(".pdf")}
    vedlegg: dict[str, list] = {}
    for filnavn in sorted(alle_filer):
        base, ext = os.path.splitext(filnavn)
        if ext.lower() != ".pdf" and base in pdf_baser:
            vedlegg.setdefault(base, []).append(filnavn)
    return {base: ", ".join(filer) for base, filer in vedlegg.items()}


# ===== PDF-lesing =====

def _er_hoved_header(rad: list) -> bool:
    """Sjekk om raden er hovedtabellens header (inneholder 'pns' og 'mengde')."""
    tekster = [str(celle).strip().lower() for celle in rad if celle]
    har_pns = any("pns" in t for t in tekster)
    har_mengde = any("mengde" in t for t in tekster)
    return har_pns and har_mengde


def _er_hoved_header_delt_venstre(rad: list) -> bool:
    """Delt-tabell-format: venstre del har 'pns' og 'postid' men ikke 'mengde'."""
    tekster = [str(celle).strip().lower() for celle in rad if celle]
    har_pns = any("pns" in t for t in tekster)
    har_postid = any("postid" in t for t in tekster)
    har_mengde = any("mengde" in t for t in tekster)
    return har_pns and har_postid and not har_mengde


def _er_hoved_header_delt_hoyre(rad: list) -> bool:
    """Delt-tabell-format: høyre del har 'mengde' og 'enhet' men ikke 'pns'."""
    tekster = [str(celle).strip().lower() for celle in rad if celle]
    har_pns = any("pns" in t for t in tekster)
    har_mengde = any("mengde" in t for t in tekster)
    har_enhet = any("enhet" in t for t in tekster)
    return har_mengde and har_enhet and not har_pns


def _er_armering_header(rad: list) -> bool:
    """Sjekk om raden er armeringstabellens header (Formkode, kg i modell, osv.)."""
    armeringsord = {"formkode", "kg i modell", "kg_i_modell",
                    "faktor for omfar", "vekt inkl omfar", "vekt inkl. omfar"}
    tekster = {str(celle).strip().lower() for celle in rad if celle}
    return bool(tekster & armeringsord)


def _er_datarad(rad: list) -> bool:
    """Sjekk om raden ser ut som en datarad (foerste kolonne har PNS-format)."""
    if not rad or len(rad) < 4:
        return False
    foerste = str(rad[0] or "").strip()
    return "." in foerste and len(foerste) > 4


def _tom_rad(rad: list) -> bool:
    """Sjekk om hele raden er tom."""
    return not rad or all(celle is None or str(celle).strip() == "" for celle in rad)


def _er_streng_prosesskode(verdi: str) -> bool:
    """Streng sjekk for gyldige prosesskoder brukt som vakt mot fremmede tabeller.

    Krav:
    - Starter med en kjent 4-sifret prefix (1302, 1303) ELLER '01'.
    - Har minst 4 punkt/bindestrek-separerte segmenter.
    Dette avviser enkle desimaltall (24.29, 0.90) og korte koder.
    """
    if not verdi:
        return False
    if not verdi.startswith(_GYLDIGE_PROSESSKODE_START):
        return False
    segmenter = re.split(r'[.\-]', verdi)
    return len(segmenter) >= 3


def _er_hoved_fortsettelse(tabell: list, forventet_kol_antall: int, apm_navn: str = "", offset: int = 0) -> bool:
    """Sjekk om en tabell uten kjent header ser ut som en fortsettelse av hoveddatatabellen.

    offset: ekstra kolonner foerst (0 = normal, 1 = pdfplumber la til tom ledende kol).
    Krav:
    - Minst en ikke-tom rad har en gyldig prosesskode-lignende verdi i kolonne 3+offset.
    - Kolonneantallet er innenfor +/- 1 av (forventet + offset).
    - Kolonne 2+offset (APM) er enten tom eller inneholder apm_navn.
    """
    if forventet_kol_antall and tabell:
        for rad in tabell:
            if not _tom_rad(rad):
                if abs(len(rad) - forventet_kol_antall - offset) > 1:
                    return False
                break
    kol_min = 4 + offset
    for rad in tabell:
        if _tom_rad(rad) or len(rad) < kol_min:
            continue
        verdi = str(rad[3 + offset] or "").strip()
        apm_verdi = str(rad[2 + offset] or "").strip()
        if apm_verdi:
            if apm_navn and apm_navn not in apm_verdi and apm_verdi not in apm_navn:
                continue
            if not (_RE_PROSESSKODE_GYLDIG.match(verdi) or _RE_PROSESSKODE_SOEK.search(verdi)):
                continue
        else:
            if not _er_streng_prosesskode(verdi):
                continue
        return True
    return False


def _finn_kolskift(tabell: list, forventet_kol_antall: int, apm_navn: str = "") -> int:
    """Finn kolonneforskyvning for en potensiell fortsettelse-tabell.

    Noen PDFer (f.eks. ved cellesammenslaaing) gir en ekstra tom kolonne foerst.
    Proever offset 0 og 1 og returnerer foerste gyldige offset.
    Returnerer -1 hvis tabellen ikke er en gyldig fortsettelse.
    """
    for offset in (0, 1):
        if _er_hoved_fortsettelse(tabell, forventet_kol_antall, apm_navn, offset):
            return offset
    return -1


def les_pdf_hoved(pdf_sti: str, omega_data: dict = None, pdf=None) -> list:
    """Les hoveddatatabellen fra en APM-PDF. Returnerer liste med rad-dicts.

    pdf: valgfritt allerede-åpnet pdfplumber.PDF-objekt (då åpnes ikkje fila på nytt).
    omega_data: valgfritt Omega-ordbok ({prosesskode: {mengde, enhet, ...}}).
    Naar oppgitt:
    - x-koordinat-basert tegn-ekstraksjon for Post-kolonnen (mot tegn-fletting)
    - PNS+enhet+subsequence-basert korreksjon mot Omega for koder som ikke finnes direkte.
    """
    rader = []
    apm_navn = os.path.splitext(os.path.basename(pdf_sti))[0]
    omega_poster = set(omega_data.keys()) if omega_data else set()
    inne_i_hovedtabell = False
    hoved_kol_antall   = 0  # sporet fra header-raden
    kolskift           = 0  # ekstra kolonneoffset (0 = normal, 1 = skjevet tabell)
    post_col_x0 = post_col_x1 = None  # x-range for Post-kolonnen (fra header-celle-bbox)

    try:
        ctx = contextlib.nullcontext(pdf) if pdf is not None else pdfplumber.open(pdf_sti)
        with ctx as pdf:
            for side in pdf.pages:
                for tabell_obj in (side.find_tables() or []):
                    tabell = tabell_obj.extract()
                    if not tabell:
                        continue

                    # Skann de foerste radene etter kjente headere
                    hoved_idx = None
                    arm_idx   = None
                    for rad_idx, rad in enumerate(tabell[:4]):
                        if rad and _er_hoved_header(rad) and hoved_idx is None:
                            hoved_idx = rad_idx
                        if rad and _er_armering_header(rad) and arm_idx is None:
                            arm_idx = rad_idx

                    if hoved_idx is not None:
                        # Ny hoveddatatabell-header funnet
                        inne_i_hovedtabell = True
                        hoved_kol_antall   = len(tabell[hoved_idx])
                        data_start = hoved_idx + 1
                        kolskift   = 0
                        # Hent Post-kolonnens x-range fra header-cellens bbox (kolonne 3)
                        try:
                            hdr_celler = tabell_obj.rows[hoved_idx].cells
                            if len(hdr_celler) > 3 and hdr_celler[3]:
                                post_col_x0 = hdr_celler[3][0]   # x0
                                post_col_x1 = hdr_celler[3][2]   # x1
                        except Exception:
                            pass
                    elif arm_idx is not None:
                        # Armeringstabellen begynner -- avslutt hoveddatatabellen
                        inne_i_hovedtabell = False
                        continue
                    elif inne_i_hovedtabell:
                        # Ingen kjent header: finn kolonneoffset (0 eller 1) og verifiser.
                        skift = _finn_kolskift(tabell, hoved_kol_antall, apm_navn)
                        if skift < 0:
                            # Ukjent tabell -- hopp over, men behold inne_i_hovedtabell
                            # slik at vi fortsatt fanger neste side av hoveddatatabellen.
                            continue
                        data_start = 0
                        kolskift   = skift
                    else:
                        continue

                    for rad_nr, rad in enumerate(tabell[data_start:]):
                        if _tom_rad(rad):
                            continue
                        o = kolskift
                        celler = list(rad) + [None] * (8 + o + 1)
                        # Ekstra vakt: kolonne 2+o (APM) skal vaere tom eller matche filnavnet.
                        # Naar kol 2+o er tom bruker vi streng prosesskode-sjekk for aa
                        # hindre at desimaltall fra fremmede tabeller leses inn.
                        apm_celle_raa = str(celler[2 + o] or "").strip()
                        # Ta bare foerste linje: pdfplumber kan inkludere linjeskift fra tekst-
                        # ombroytning i cellen, som ellers bryter substreng-sammenligningen.
                        apm_celle = apm_celle_raa.split("\n")[0].strip()
                        if apm_celle:
                            if apm_navn not in apm_celle and apm_celle not in apm_navn:
                                # APM-celle samsvarer ikke direkte -- kan vaere PDF-tekstkorrupsjon.
                                # Fall back: godta raden hvis prosesskoden er strengt gyldig.
                                if not _er_streng_prosesskode(_rens(celler[3 + o])):
                                    continue
                        else:
                            # Tom APM-kolonne: krev streng prosesskode i kol 3+o
                            rens_val = _rens(celler[3 + o])
                            if not _er_streng_prosesskode(rens_val):
                                # Prøv normalisering (f.eks. '11302-...' -> '1302-...')
                                norm_val = _normaliser_prosesskode(rens_val)
                                if not _er_streng_prosesskode(norm_val or ""):
                                    continue

                        # Hent Post-tekst: forsoek x-koordinat-basert ekstraksjon foerst
                        # (beskytter mot tegn-fletting fra lang APM-kolonne-tekst).
                        post_raa = _rens(celler[3 + o])
                        if post_col_x0 is not None:
                            try:
                                rad_idx_tab = data_start + rad_nr
                                if rad_idx_tab < len(tabell_obj.rows):
                                    rad_celler = tabell_obj.rows[rad_idx_tab].cells
                                    # Post-kolonnens fysiske x-posisjon er uavhengig av kolskift.
                                    # Finn cellen som starter naer post_col_x0 (innenfor 20 pt).
                                    post_celle = None
                                    for c in rad_celler:
                                        if c and abs(c[0] - post_col_x0) < 20:
                                            post_celle = c
                                            break
                                    if post_celle is None and len(rad_celler) > 3 + o:
                                        post_celle = rad_celler[3 + o]
                                    if post_celle:
                                        cx0, cy0, cx1, cy1 = post_celle
                                        chars_tekst = _hent_tekst_fra_celle_via_chars(
                                            side, cx0, cy0, cx1, cy1
                                        )
                                        if chars_tekst:
                                            post_raa = chars_tekst
                            except Exception:
                                pass

                        prosesskode = _normaliser_prosesskode(post_raa, advar_fn=print)
                        if not prosesskode:
                            continue

                        # Lagre original kode foer potensiell dekorrupsjon (brukes av gruppe-konsistens)
                        raa_prosesskode = prosesskode

                        # Omega-kontroll: korriger koder som ikke finnes direkte i Omega.
                        # Bruker segment-trunkering av tegnfletting-resultatet:
                        # tidlige segmenter er korrekte, ekstra sluttegmenter er forurensning
                        # fra APM-filnavnets siffer (f.eks. '205' fra 'APM205').
                        if omega_data and prosesskode not in omega_poster:
                            # Prøv trailing-null i andre segment (f.eks. '1302-030' -> '1302-03')
                            korrigert = _normaliser_andre_segment(prosesskode, omega_poster, advar_fn=print)
                            if korrigert != prosesskode:
                                prosesskode = korrigert
                        if omega_data and prosesskode not in omega_poster:
                            # Prøv å bytte første '.' til '-' (f.eks. '1303.03.82.191' -> '1303-03.82.191')
                            korrigert = _normaliser_foerste_separator(prosesskode, omega_poster, advar_fn=print)
                            if korrigert != prosesskode:
                                prosesskode = korrigert
                        if omega_data and prosesskode not in omega_poster:
                            pns_val   = _rens(celler[0 + o])
                            enhet_val = _rens(celler[5 + o])
                            ny = _dekorrupt_via_omega(
                                prosesskode, pns_val, enhet_val, omega_data, advar_fn=print
                            )
                            if ny:
                                prosesskode = ny

                        rader.append({
                            KOL_MANGLER: "", "!": "", "PNS": _rens(celler[0 + o]),
                            "Ekstra inndeling": _rens(celler[1 + o]),
                            "APM": apm_navn, "Post": prosesskode,
                            "Mengde fra APM": _rens_mengde(celler[4 + o]),
                            "Enhet fra APM": _rens(celler[5 + o]),
                            "EO": _rens(celler[6 + o]),
                            "Kommentar": _rens(celler[7 + o]),
                            "Utgatt": "", "Ny_APM": "", "Vedlegg": "",
                            KOL_OMEGA: None, KOL_KONTRAKT: None,
                            KOL_ENHETSPRIS: None, KOL_PROSENT: None,
                            "_raa_post": raa_prosesskode,
                            "_pdf_post": post_raa,
                        })
    except Exception as feil:
        print(f"    [WARNING] Hoved: {os.path.basename(pdf_sti)}: {feil}")

    # Andre runde: fikse gjenvaerende korrupte koder ved gruppe-konsistens
    if omega_data and rader:
        _fix_gruppe_konsistens(rader, omega_data, omega_poster, advar_fn=print)

    # Tredje runde: rapporter rader som fortsatt ikke finnes i Omega etter alle fiksforsok.
    # Disse kan ikke rettes automatisk og krever manuell gjennomgang av PDF-filen.
    if omega_data and rader:
        ukjente = [
            rad for rad in rader
            if rad["Post"] not in omega_poster
        ]
        if ukjente:
            print(
                f"    [FEIL] {apm_navn}: {len(ukjente)} rad(er) med ukjent prosesskode "
                f"etter alle rettingsforsok -- PDF bor sjekkes manuelt:"
            )
            for rad in ukjente:
                print(f"        - Post='{rad['Post']}' | Enhet={rad['Enhet fra APM']} | "
                      f"Kommentar={rad['Kommentar']}")

    return rader


def les_pdf_erstatter(pdf_sti: str, pdf=None) -> list:
    """Finn 'ERSTATTER APMxx'-referanser i fritekst. Returnerer liste med APM-IDer."""
    funnet = []
    try:
        ctx = contextlib.nullcontext(pdf) if pdf is not None else pdfplumber.open(pdf_sti)
        with ctx as pdf:
            for side in pdf.pages:
                tekst = side.extract_text() or ""
                for match in _RE_ERSTATTER.finditer(tekst):
                    apm_id = re.sub(r"\s+", "", match.group(1)).upper()
                    if apm_id not in funnet:
                        funnet.append(apm_id)
    except Exception:
        pass
    return funnet


def les_pdf_armering(pdf_sti: str, pdf=None) -> list:
    """Les armeringsdata (Formkode-tabeller) fra en APM-PDF."""
    rader = []
    apm_navn = os.path.splitext(os.path.basename(pdf_sti))[0]

    try:
        ctx = contextlib.nullcontext(pdf) if pdf is not None else pdfplumber.open(pdf_sti)
        with ctx as pdf:
            for side in pdf.pages:
                for tabell in (side.extract_tables() or []):
                    if not tabell:
                        continue
                    # Finn armerings-header (kan vaere paa rad 0, 1 eller 2)
                    header_rad, data_start = None, None
                    for rad_idx in range(min(3, len(tabell))):
                        if _er_armering_header(tabell[rad_idx]):
                            header_rad = tabell[rad_idx]
                            data_start = rad_idx + 1
                            break
                    if data_start is None:
                        continue

                    # Map kolonnenavn til indekser
                    kol_idx = {
                        _ARM_ALIAS[str(celle or "").strip().lower()]: pos
                        for pos, celle in enumerate(header_rad)
                        if str(celle or "").strip().lower() in _ARM_ALIAS
                    }
                    if "Post" not in kol_idx:
                        kol_idx = {"Post": 0, "Formkode": 1, "kg_i_modell": 2,
                                   "Faktor_omfar": 3, "Vekt_inkl_omfar": 4}

                    for rad in tabell[data_start:]:
                        if _tom_rad(rad):
                            continue
                        celler = list(rad) + [None] * 10
                        hent = lambda noekkkel: celler[kol_idx[noekkkel]] if noekkkel in kol_idx else None
                        prosesskode = _normaliser_prosesskode(_rens(hent("Post")), advar_fn=print)
                        if not prosesskode:
                            continue
                        _kg   = _parse_tall(hent("kg_i_modell"))
                        _fak  = _parse_tall(hent("Faktor_omfar"))
                        _vekt = _parse_tall(hent("Vekt_inkl_omfar"))
                        if _vekt is None and _kg is not None:
                            _vekt = round(_kg * (_fak if _fak is not None else 1.0), 4)
                        rader.append({
                            "APM": apm_navn, "Post": prosesskode,
                            "Formkode": _rens(hent("Formkode")),
                            "kg_i_modell": _kg,
                            "Faktor_omfar": _fak,
                            "Vekt_inkl_omfar": _vekt,
                        })
    except Exception as feil:
        print(f"    [WARNING] Armering: {os.path.basename(pdf_sti)}: {feil}")

    return rader


# ===== Omega-kobling =====

_NOR_MAANED = [
    "Januar", "Februar", "Mars", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Desember",
]


def _siste_soendag_i_maaned(aar: int, maaned: int) -> datetime.date:
    """Returnerer datoen for siste sondag i angitt aar og maaned."""
    import calendar
    siste_dag = calendar.monthrange(aar, maaned)[1]
    d = datetime.date(aar, maaned, siste_dag)
    while d.weekday() != 6:  # 6 = Sunday
        d -= datetime.timedelta(days=1)
    return d


def _dato_til_maaned_aar(dato) -> str:
    """Konverter send-dato til maaledssyklus-maaned.

    Syklus (fast grense: 20. i maaneden):
    - Sendt ETTER den 20.: tilhoerer syklus M (denne maaneden).
    - Sendt PAA eller FOER den 20.: tilhoerer syklus M-1 (forrige maaned).

    Eks: 26.03.2026 (etter 20.03.2026) -> 'Mars 2026'
         15.03.2026 (foer 20.03.2026)  -> 'Februar 2026'
    """
    if dato is None or dato == "":
        return ""
    if isinstance(dato, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                dato = datetime.datetime.strptime(dato.strip(), fmt)
                break
            except ValueError:
                pass
        else:
            return str(dato)
    if not hasattr(dato, "month"):
        return str(dato)
    dato_d = dato.date() if isinstance(dato, datetime.datetime) else dato
    if dato_d.day > 20:
        # Sendt etter den 20.: denne mnd er syklus-maaneden
        return f"{_NOR_MAANED[dato_d.month - 1]} {dato_d.year}"
    else:
        # Sendt paa eller foer den 20.: forrige mnd er syklus-maaneden
        if dato_d.month == 1:
            return f"Desember {dato_d.year - 1}"
        return f"{_NOR_MAANED[dato_d.month - 2]} {dato_d.year}"


_RE_KOMMENTAR_DATO = re.compile(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})")


def _til_dato(verdi):
    """Normaliser ein verdi (datetime/streng) til datetime.date, eller None."""
    if verdi is None or verdi == "":
        return None
    if isinstance(verdi, datetime.datetime):
        return verdi.date()
    if isinstance(verdi, datetime.date):
        return verdi
    if isinstance(verdi, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(verdi.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _ekstraher_kommentardato(tekst):
    """Hent fyrste dato (d.m.yyyy) frå ein kommentar-streng, eller None."""
    if not tekst:
        return None
    m = _RE_KOMMENTAR_DATO.search(str(tekst))
    if not m:
        return None
    d, mnd, aar = (int(x) for x in m.groups())
    try:
        return datetime.date(aar, mnd, d)
    except ValueError:
        return None


def _maalebrev_maaned(delmaal, boqa) -> str:
    """Syklus-månad for ein post, justert mot kommentar-dato.

    Regel (frå brukar):
      - Dersom datoen i 'Kommentar fra forrige delmålebrev' er lik datoen i
        'Sist oppdatert', vart kommentaren logga samstundes med siste
        oppdatering → bruk månaden FØR Sist-oppdatert sin eigen månad.
        (Den 20.-regelen gjeld IKKJE her: både 07.10 og 21.10 -> September.)
      - Elles brukast 'Sist oppdatert' (siste dato) som grunnlag, med
        den vanlege 20.-regelen.
    """
    ddato = _til_dato(delmaal)
    if ddato is None:
        return _dato_til_maaned_aar(delmaal)   # tom -> "", uparsbar streng -> str

    kdato = _ekstraher_kommentardato(boqa)
    if kdato is not None and kdato == ddato:
        # Månaden før Sist-oppdatert sin kalendermånad
        if ddato.month == 1:
            return f"Desember {ddato.year - 1}"
        return f"{_NOR_MAANED[ddato.month - 2]} {ddato.year}"

    return _dato_til_maaned_aar(delmaal)


def _sres_til_tekst(verdi) -> str:
    """Konverter SRES-verdi (bool/tekst) til 'Sann' / 'Usann' / ''."""
    if verdi is None or verdi == "":
        return ""
    if isinstance(verdi, bool):
        return "Sann" if verdi else "Usann"
    s = str(verdi).strip().lower()
    if s in ("true", "sann", "1", "yes", "ja"):
        return "Sann"
    if s in ("false", "usann", "0", "no", "nei"):
        return "Usann"
    return str(verdi)


_PDF_OMEGA_CACHE: dict = {}


def les_omega(sti: str) -> dict:
    """Les Omega BOQNT-arket.

    Returnerer {prosesskode: {"mengde", "kontrakt", "pris", "enhet", "delmaal", "sres"}}
    for ALLE poster som finnes i arket, uavhengig av om de har kumulativ mengde.
    Cachar resultatet på (path, mtime).
    """
    # Cache: (path, mtime) — invalideres om fila endrast
    try:
        cache_key = (sti, os.path.getmtime(sti))
        if cache_key in _PDF_OMEGA_CACHE:
            kakra = _PDF_OMEGA_CACHE[cache_key]
            print(f"  Omega lest: (cached, {len(kakra)} prosesskoder)")
            return kakra
    except OSError:
        cache_key = None

    omega_data: dict[str, dict] = {}
    try:
        wb = load_workbook(sti, read_only=True, data_only=True)
        ark = wb["BOQNT"] if "BOQNT" in wb.sheetnames else wb.active
        post_kol, mengde_kol, kontrakt_kol, pris_kol, enhet_kol, delmaal_kol, sres_kol, boqa_kol = None, None, None, None, None, None, None, None
        for rad in ark.iter_rows(values_only=True):
            # Finn header-raden foerst
            if post_kol is None:
                celler = [str(celle or "").strip() for celle in rad]
                if "Navn (Prosesskode)" not in celler:
                    continue
                post_kol = celler.index("Navn (Prosesskode)")
                for idx, hdr in enumerate(celler):
                    hdr_l = hdr.lower()
                    if mengde_kol is None and "rapportert" in hdr_l and "kumulativ" in hdr_l:
                        mengde_kol = idx
                    elif mengde_kol is None and hdr_l.strip() == "rapportert":          # April-format
                        mengde_kol = idx
                    elif kontrakt_kol is None and "kontraktsmengde" in hdr_l:
                        kontrakt_kol = idx
                    elif kontrakt_kol is None and hdr_l.strip() == "opprinnelig":    # April-format
                        kontrakt_kol = idx
                    elif pris_kol is None and "enhetspris" in hdr_l:
                        pris_kol = idx
                    elif enhet_kol is None and "enhet" in hdr_l:
                        enhet_kol = idx
                    elif delmaal_kol is None and "sist" in hdr_l and ("delm" in hdr_l or "brev" in hdr_l):
                        delmaal_kol = idx
                    elif delmaal_kol is None and "sist" in hdr_l and "oppdatert" in hdr_l:  # April-format
                        delmaal_kol = idx
                    elif sres_kol is None and hdr_l.strip() == "sres":
                        sres_kol = idx
                    elif boqa_kol is None and "boqa" in hdr_l and "kommentar" in hdr_l:
                        boqa_kol = idx
                    elif boqa_kol is None and "kommentar" in hdr_l and ("forrige" in hdr_l or "delm" in hdr_l):  # April-format
                        boqa_kol = idx
                continue
            if post_kol is None:
                continue
            prosesskode = str(rad[post_kol] or "").strip()
            if not prosesskode:
                continue

            # Hent verdier -- None hvis kolonnen ikke finnes eller cellen er tom/tekst
            def _hent_tall(kol_idx):
                if kol_idx is None:
                    return None
                v = rad[kol_idx]
                return v if isinstance(v, (int, float)) else None

            def _hent_tekst(kol_idx):
                if kol_idx is None:
                    return None
                v = rad[kol_idx]
                return _rens(str(v)) if v else None

            mengde_verdi   = _hent_tall(mengde_kol)
            kontrakt_verdi = _hent_tall(kontrakt_kol)
            pris_verdi     = _hent_tall(pris_kol)
            enhet_verdi    = _hent_tekst(enhet_kol)

            # Dato og SRES -- lagres raast; konverteres til tekst ved utskrift
            delmaal_verdi = rad[delmaal_kol] if delmaal_kol is not None else None
            sres_verdi    = rad[sres_kol]    if sres_kol    is not None else None
            boqa_verdi    = _hent_tekst(boqa_kol)

            # Legg til posten uansett -- selv om kumulativ mengde er tom
            omega_data[prosesskode] = {
                "mengde":   mengde_verdi,
                "kontrakt": kontrakt_verdi,
                "pris":     pris_verdi,
                "enhet":    enhet_verdi,
                "delmaal":  delmaal_verdi,
                "sres":     sres_verdi,
                "boqa":     boqa_verdi,
            }
        wb.close()
        har_kumul    = mengde_kol is not None
        har_kontrakt = kontrakt_kol is not None
        har_pris     = pris_kol is not None
        har_enhet    = enhet_kol is not None
        har_delmaal  = delmaal_kol is not None
        har_sres     = sres_kol is not None
        print(f"  Omega lest: {len(omega_data)} prosesskoder"
              f" | kumulativ: {'ja' if har_kumul else 'ikke funnet'}"
              f" | kontraktsmengde: {'ja' if har_kontrakt else 'ikke funnet'}"
              f" | enhetspris: {'ja' if har_pris else 'ikke funnet'}"
              f" | enhet: {'ja' if har_enhet else 'ikke funnet'}"
              f" | delmaalebrev: {'ja' if har_delmaal else 'ikke funnet'}"
              f" | sres: {'ja' if har_sres else 'ikke funnet'}")
    except Exception as feil:
        print(f"[WARNING] Omega: {feil}")
    if cache_key and omega_data:
        _PDF_OMEGA_CACHE[cache_key] = omega_data
    return omega_data


# ===== APM betong/forskaling konsistenssjekk =====

_APM_TYKKELSE_RE = re.compile(r'\b(\d{2,4})\s*(?:mm|$|\s+B\d)', re.IGNORECASE)
_APM_SKIP_TYPER  = {"søyle", "soyle", "bjelke", "ramme"}


def _apm_forskaling_faktor(ktype: str) -> float:
    """Antal sider med forskaling: vegg/parapet = 2, plate/dekke = 1."""
    if any(s in ktype for s in ("vegg", "parapet")):
        return 2.0
    return 1.0

KOLONNER_APM_BF = [
    ("Gruppe (prefiks)",     22, False),
    ("Beskrivelse",          30, False),
    ("Tykkelse (mm)",        14, True),
    ("Betong-post",          22, False),
    ("APM",                  40, False),
    ("APM betong (m³)",      16, True),
    ("Forv. forskal. (m²)",  22, True),
    ("Forskaling-post",      22, False),
    ("APM forskal. (m²)",    20, True),
    ("Avvik (%)",            12, True),
    ("Kontrakt betong (m³)", 20, True),
    ("Rapp. betong (m³)",    20, True),
]


def bygg_apm_betong_forskaling(
    df_hoved: pd.DataFrame,
    omega_data: dict,
    isy_map: dict,
) -> list[dict]:
    """Sjekk APM betong (m³) mot tilhøyrande forskalingspost (m²) via ISY-tykkelse."""
    aktiv = df_hoved[df_hoved["Utgatt"].str.lower() != "ja"].copy()
    aktiv["_enhet"] = aktiv["Enhet fra APM"].astype(str).str.strip().str.lower()
    aktiv["_mengde"] = pd.to_numeric(
        aktiv["Mengde fra APM"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    aktiv["_post"]   = aktiv["Post"].astype(str).str.strip()
    aktiv["_gruppe"] = aktiv["_post"].apply(
        lambda k: k.rsplit(".", 1)[0] if "." in k else ""
    )

    betong_rader = aktiv[aktiv["_enhet"].isin(["m3", "m³"])]
    resultat: list[dict] = []
    sette: set = set()

    for _, row in betong_rader.iterrows():
        kode       = row["_post"]
        apm        = str(row["APM"])
        gruppe_pnr = row["_gruppe"]
        if not gruppe_pnr:
            continue
        key = (apm, kode)
        if key in sette:
            continue
        sette.add(key)

        gruppe_besk  = isy_map.get(gruppe_pnr, {}).get("beskrivelse", "")
        gruppe_tittel = gruppe_besk.split("\n")[0].strip()

        ktype = gruppe_tittel.split()[0].lower() if gruppe_tittel.strip() else ""
        if any(s in ktype for s in _APM_SKIP_TYPER):
            continue

        m = _APM_TYKKELSE_RE.search(gruppe_tittel)
        if not m:
            continue
        tykkelse_mm = float(m.group(1))
        tykkelse_m  = tykkelse_mm / 1000.0
        fk_faktor   = _apm_forskaling_faktor(ktype)  # 2 for vegg/parapet

        apm_betong = row["_mengde"]
        if not pd.notna(apm_betong):
            continue

        # Finn forskalingspost: same APM, same gruppe, m2, FORSKALING i ISY-tittel
        gruppe_prefix = gruppe_pnr + "."
        mask_fk = (
            (aktiv["APM"] == apm)
            & aktiv["_post"].str.startswith(gruppe_prefix)
            & aktiv["_enhet"].isin(["m2", "m²"])
        )
        forskaling_kode = ""
        apm_forskaling  = None
        for _, fk in aktiv[mask_fk].iterrows():
            isy_t = isy_map.get(fk["_post"], {}).get("beskrivelse", "").split("\n")[0].upper()
            if "FORSKALING" in isy_t:
                forskaling_kode = fk["_post"]
                apm_forskaling  = fk["_mengde"]
                break

        forventa = round(float(apm_betong) / tykkelse_m * fk_faktor, 2) if tykkelse_m else None
        if forventa and apm_forskaling is not None and pd.notna(apm_forskaling) and forventa != 0:
            avvik_pst = round(float(apm_forskaling) / forventa * 100 - 100, 1)
        else:
            avvik_pst = None

        om = omega_data.get(kode, {})
        resultat.append({
            "gruppe_pnr":       gruppe_pnr,
            "gruppe_tittel":    gruppe_tittel,
            "tykkelse_mm":      tykkelse_mm,
            "betong_kode":      kode,
            "apm_namn":         apm,
            "apm_betong":       round(float(apm_betong), 3),
            "forventa_forsk":   forventa,
            "forskaling_kode":  forskaling_kode,
            "apm_forsk":        round(float(apm_forskaling), 2) if apm_forskaling is not None and pd.notna(apm_forskaling) else None,
            "avvik_pst":        avvik_pst,
            "kontr_betong":     om.get("kontrakt"),
            "rapp_betong":      om.get("mengde"),
        })

    resultat.sort(key=lambda r: (r["betong_kode"], r["apm_namn"]))
    return resultat


_APM_BF_NØKLAR = [
    "gruppe_pnr", "gruppe_tittel", "tykkelse_mm", "betong_kode", "apm_namn",
    "apm_betong", "forventa_forsk", "forskaling_kode", "apm_forsk",
    "avvik_pst", "kontr_betong", "rapp_betong",
]


# ===== Parallell PDF-prosessering =====

def _les_en_pdf_worker(args):
    """Worker: les ein PDF (hoved + armering + erstatter) og returner alt.

    Fangar stdout/stderr slik at GUI-loggen får meldingane i samla bolkar.
    Må vere på modul-nivå for å vere picklable av ProcessPoolExecutor.
    """
    import io, contextlib
    mappe, filnavn, omega_data, er_utgaatt, ny_apm, vedlegg = args
    full_sti = os.path.join(mappe, filnavn)

    buf = io.StringIO()
    feil = None
    hoved_rader, arm_rader, erstatter_liste = [], [], []
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            # Opne PDF berre éin gong og send same objekt til alle tre lesarane
            with pdfplumber.open(full_sti) as pdf:
                hoved_rader     = les_pdf_hoved(full_sti, omega_data=omega_data, pdf=pdf)
                arm_rader       = les_pdf_armering(full_sti, pdf=pdf)
                erstatter_liste = les_pdf_erstatter(full_sti, pdf=pdf)
    except Exception as e:
        feil = f"{type(e).__name__}: {e}"

    if er_utgaatt:
        for rad in hoved_rader:
            rad["Utgatt"] = "Ja"
            rad["Ny_APM"] = ny_apm
    if vedlegg:
        for rad in hoved_rader:
            rad["Vedlegg"] = vedlegg

    return filnavn, hoved_rader, arm_rader, erstatter_liste, feil, buf.getvalue()


# ===== Hovedprosessering =====

ARK_RETTEDE = "Rettede poster"
KOLONNER_RETTEDE = ["Feiltype", "APM", "Post i PDF", "Rettet til", "Detalj"]


def _registrer_feil_poster(df_hoved, pdf_poster: list, omega_data: dict) -> list[dict]:
    """Meld APM-poster som er retta av koden, eller har feil enhet, i feillista.

    Returnerer radene òg, så dei kan lagrast i arket ARK_RETTEDE — då kjem dei
    med i feillista sjølv når APM-steget vert hoppa over seinare.
    """
    import feil_poster

    rader: list[dict] = []
    n_rettet = n_enhet = 0
    for i, idx in enumerate(df_hoved.index):
        post = str(df_hoved.at[idx, "Post"] or "").strip()
        apm = df_hoved.at[idx, "APM"]
        fra_pdf = str(pdf_poster[i] if i < len(pdf_poster) else "").strip()

        # Mellomrom og linjeskift i PDF-teksten er ikkje ei retting
        if fra_pdf and "".join(fra_pdf.split()) != post:
            rader.append({"Feiltype": feil_poster.RETTET_AUTOMATISK, "APM": apm,
                          "Post i PDF": fra_pdf, "Rettet til": post,
                          "Detalj": "Posten i PDF-en er rettet av koden"})
            n_rettet += 1

        enhet_apm = _normaliser_enhet(str(df_hoved.at[idx, "Enhet fra APM"] or ""))
        enhet_omega = _normaliser_enhet((omega_data.get(post) or {}).get("enhet") or "")
        if enhet_apm and enhet_omega and enhet_apm != enhet_omega:
            rader.append({"Feiltype": feil_poster.FEIL_ENHET, "APM": apm,
                          "Post i PDF": post, "Rettet til": "",
                          "Detalj": f"Enhet i APM: {enhet_apm}, i Omega: {enhet_omega}"})
            n_enhet += 1

    registrer_rettede(rader)
    if n_rettet or n_enhet:
        print(f"  Feilliste: {n_rettet} poster rettet automatisk, "
              f"{n_enhet} med annen enhet enn Omega")
    return rader


def registrer_rettede(rader) -> None:
    """Meld rader frå ARK_RETTEDE (eller _registrer_feil_poster) i feillista."""
    import feil_poster
    for r in rader:
        feil_poster.registrer(
            r.get("Feiltype", ""), kilde="APM", post=r.get("Post i PDF"),
            rettet_til=r.get("Rettet til"), element=r.get("APM"),
            detalj=r.get("Detalj"))


def _skriv_rettede_ark(output_path: str, rader: list[dict]) -> None:
    """Legg arket ARK_RETTEDE til i den ferdige Maalebrev-fila."""
    from lagring import vent_til_skrivbar
    vent_til_skrivbar(output_path)
    wb = load_workbook(output_path)
    if ARK_RETTEDE in wb.sheetnames:
        del wb[ARK_RETTEDE]
    ws = wb.create_sheet(ARK_RETTEDE)
    ws.append(KOLONNER_RETTEDE)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rader:
        ws.append([r.get(k, "") for k in KOLONNER_RETTEDE])
    for kol, breidd in zip("ABCDE", (18, 40, 26, 26, 50)):
        ws.column_dimensions[kol].width = breidd
    ws.freeze_panes = "A2"
    wb.save(output_path)


def extract_folder(mappe: str, output_path: str = None, omega_path: str = None,
                   isy_path=None,
                   test_modus: bool = False,
                   max_workers=None) -> str:
    """Les alle APM-PDFer i mappen og skriv samlet Excel-fil.

    isy_path:   ISY Linker-fil(er) for betong/forskaling-sjekk (valfri).
    test_modus: velg tilfeldig 15 % av filene for rask feilsoking.
    """
    import random
    if output_path is None:
        output_path = os.path.join(os.path.dirname(os.path.normpath(mappe)), OUTPUT_FILNAVN)

    pdf_filer = sorted(fil for fil in os.listdir(mappe) if fil.lower().endswith(".pdf"))
    if not pdf_filer:
        print("[WARNING] Ingen PDF-filer funnet.")
        return ""

    if test_modus:
        totalt_antall = len(pdf_filer)
        antall_test = max(1, round(totalt_antall * 0.15))
        pdf_filer = sorted(random.sample(pdf_filer, antall_test))
        print(f"[TEST] Kjorer med {antall_test} av {totalt_antall} PDF-filer (30 % tilfeldig utvalg).")

    # Les Omega foerst slik at kjente prosesskoder kan brukes som kontroll
    # under PDF-parsing (x-koordinat-korreksjon og Omega-nearest-match).
    omega_data: dict = {}
    omega_poster: set = set()
    if omega_path and os.path.isfile(omega_path):
        print(f"Leser Omega: {os.path.basename(omega_path)}")
        omega_data = les_omega(omega_path)
        omega_poster = set(omega_data.keys())
        print()

    print("Analyserer fil-duplikater (MD5-hash)...")
    hopp_over, utgaatt_map, alle_stier = _analyser_filer(mappe)
    vedlegg_map = _finn_vedlegg(mappe)
    antall = len(pdf_filer)

    alle_hoved, alle_arm = [], []
    erstatter_map: dict[str, str] = {}  # {"APM65": "filnavn_som_erstatter"}

    # Bygg oppgaveliste: hopp over identiske kopiar med ein gong
    oppgaver: list[tuple] = []
    for filnavn in pdf_filer:
        if filnavn in hopp_over:
            continue
        apm_navn   = os.path.splitext(filnavn)[0]
        er_utgaatt = apm_navn in utgaatt_map
        ny_apm     = utgaatt_map.get(apm_navn, "")
        vedlegg    = vedlegg_map.get(apm_navn, "")
        oppgaver.append((mappe, filnavn, omega_data, er_utgaatt, ny_apm, vedlegg))

    if max_workers is None:
        max_workers = max(1, min(len(oppgaver), (os.cpu_count() or 4) - 1))

    # Skriv ut "hoppet over"-meldingar opp front
    for filnavn in pdf_filer:
        if filnavn in hopp_over:
            print(f"  [---/{antall}] {filnavn} -> [IDENTISK KOPI - hoppet over]")

    print(f"Parser {len(oppgaver)} PDF-filer i {max_workers} parallelle prosessar...")

    resultat_map: dict[str, tuple] = {}

    if max_workers > 1 and len(oppgaver) > 1:
        from concurrent.futures import ProcessPoolExecutor, as_completed
        with ProcessPoolExecutor(max_workers=max_workers) as ex:
            futs = {ex.submit(_les_en_pdf_worker, opg): opg[1] for opg in oppgaver}
            for ferdig, fut in enumerate(as_completed(futs), 1):
                filnavn, hoved, arm, ersta, feil, log = fut.result()
                if log:
                    print(log, end="")
                apm_navn = os.path.splitext(filnavn)[0]
                er_utg = apm_navn in utgaatt_map
                status = f" [UTGAATT -> {utgaatt_map.get(apm_navn) or '?'}]" if er_utg else ""
                if feil:
                    print(f"  [{ferdig:>3}/{len(oppgaver)}] {filnavn} -> [FEIL: {feil}]")
                else:
                    print(f"  [{ferdig:>3}/{len(oppgaver)}] {filnavn} -> "
                          f"{len(hoved)} rader, {len(arm)} arm{status}")
                resultat_map[filnavn] = (hoved, arm, ersta)
    else:
        # Sekvensiell fallback (få filer eller 1 kjerne)
        for nr, opg in enumerate(oppgaver, 1):
            filnavn, hoved, arm, ersta, feil, log = _les_en_pdf_worker(opg)
            if log:
                print(log, end="")
            apm_navn = os.path.splitext(filnavn)[0]
            er_utg = apm_navn in utgaatt_map
            status = f" [UTGAATT -> {utgaatt_map.get(apm_navn) or '?'}]" if er_utg else ""
            if feil:
                print(f"  [{nr:>3}/{len(oppgaver)}] {filnavn} -> [FEIL: {feil}]")
            else:
                print(f"  [{nr:>3}/{len(oppgaver)}] {filnavn} -> "
                      f"{len(hoved)} rader, {len(arm)} arm{status}")
            resultat_map[filnavn] = (hoved, arm, ersta)

    # Samle i deterministisk rekkefølge (filnavn-sortert) og bygg erstatter-map
    for filnavn in pdf_filer:
        if filnavn not in resultat_map:
            continue
        hoved, arm, ersta = resultat_map[filnavn]
        alle_hoved.extend(hoved)
        alle_arm.extend(arm)
        apm_navn = os.path.splitext(filnavn)[0]
        for erstattet_apm in ersta:
            erstatter_map[erstattet_apm] = apm_navn
            print(f"    [ERSTATTER] {apm_navn} erstatter {erstattet_apm}")

    # Marker rader som utgaatt basert paa 'ERSTATTER'-tekst i PDF-ene
    if erstatter_map:
        for rad in alle_hoved:
            match = _RE_APM_NR.search(rad["APM"])
            if not match:
                continue
            apm_nokkel = f"APM{match.group(1)}"
            if apm_nokkel in erstatter_map and rad.get("Utgatt") != "Ja":
                rad["Utgatt"] = "Ja"
                rad["Ny_APM"] = erstatter_map[apm_nokkel]
                print(f"    [UTGAATT] {rad['APM']} markert utgaatt, erstattet av {rad['Ny_APM']}")

    if not alle_hoved and not alle_arm:
        print("[ERROR] Ingen data funnet.")
        return ""

    df_hoved = pd.DataFrame(alle_hoved, columns=KOLONNER_HOVED)
    df_arm = pd.DataFrame(alle_arm, columns=KOLONNER_ARM)
    # Posten slik han står i PDF-en — same rekkjefølgje som df_hoved
    pdf_poster = [rad.get("_pdf_post", "") for rad in alle_hoved]
    rettede_rader: list[dict] = []

    # Omega-kobling: slaa opp prosesskode mot rapportert kumulativ mengde
    # (omega_data og omega_poster er allerede lest inn ovenfor)
    if omega_data:
        # Map hvert felt fra Omega-ordboken (None hvis posten ikke finnes)
        df_hoved[KOL_OMEGA]      = df_hoved["Post"].map(lambda p: (omega_data.get(p) or {}).get("mengde"))
        df_hoved[KOL_KONTRAKT]   = df_hoved["Post"].map(lambda p: (omega_data.get(p) or {}).get("kontrakt"))
        df_hoved[KOL_ENHETSPRIS] = df_hoved["Post"].map(lambda p: (omega_data.get(p) or {}).get("pris"))
        df_hoved[KOL_DELMAAL]    = df_hoved["Post"].map(lambda p: _maalebrev_maaned((omega_data.get(p) or {}).get("delmaal"), (omega_data.get(p) or {}).get("boqa")))
        df_hoved[KOL_SRES]       = df_hoved["Post"].map(lambda p: _sres_til_tekst((omega_data.get(p) or {}).get("sres")))

        # Ikke-i-Omega-flagg: post finnes i PDF men eksisterer ikke i Omega overhodet
        # (skilles fra poster som finnes men har tom kumulativ mengde)
        omega_poster = set(omega_data.keys())
        ikke_i_omega = ~df_hoved["Post"].isin(omega_poster)
        df_hoved.loc[ikke_i_omega, KOL_MANGLER] = "!"

        # APM-prefix-korreksjon: for poster som mangler i Omega, forsok aa rekonstruere
        # prosesskoden ved aa bytte ledende segmenter med numerisk prefix fra APM-navn.
        # Eksempel: '0811.06.241.1.1' + APM '1303_01_LYA' -> '1303.01.06.241.1.1'
        mangler_mask = df_hoved[KOL_MANGLER] == "!"
        if mangler_mask.any():
            antall_korrigert = 0
            for idx in df_hoved.index[mangler_mask]:
                post_orig = str(df_hoved.at[idx, "Post"] or "")
                apm_navn  = str(df_hoved.at[idx, "APM"] or "")
                ny_post = _korriger_post_med_apm_prefix(post_orig, apm_navn, omega_poster)
                if ny_post:
                    df_hoved.at[idx, "Post"]         = ny_post
                    omega_info = omega_data.get(ny_post) or {}
                    df_hoved.at[idx, KOL_OMEGA]      = omega_info.get("mengde")
                    df_hoved.at[idx, KOL_KONTRAKT]   = omega_info.get("kontrakt")
                    df_hoved.at[idx, KOL_ENHETSPRIS] = omega_info.get("pris")
                    df_hoved.at[idx, KOL_DELMAAL]    = _maalebrev_maaned(omega_info.get("delmaal"), omega_info.get("boqa"))
                    df_hoved.at[idx, KOL_SRES]       = _sres_til_tekst(omega_info.get("sres"))
                    df_hoved.at[idx, KOL_MANGLER]    = ""
                    antall_korrigert += 1
                    print(f"    [KORRIGERT] '{post_orig}' -> '{ny_post}' (APM: {apm_navn})")
            if antall_korrigert:
                print(f"  APM-prefix-korreksjon: {antall_korrigert} poster korrigert")
                # Oppdater ikke_i_omega-masken etter korreksjon
                ikke_i_omega = df_hoved[KOL_MANGLER] == "!"

        try:
            rettede_rader = _registrer_feil_poster(df_hoved, pdf_poster, omega_data)
        except Exception as exc:      # feillista skal aldri velte APM-uttrekket
            print(f"  [ADVARSEL] Feillista: {exc}")

        # Vektorisert prosent-beregning (raskere enn apply rad-for-rad)
        mengde_serie = pd.to_numeric(df_hoved["Mengde fra APM"], errors="coerce")
        omega_serie  = pd.to_numeric(df_hoved[KOL_OMEGA], errors="coerce")
        gyldig = mengde_serie.notna() & omega_serie.notna() & (mengde_serie != 0) & (omega_serie != 0)
        df_hoved[KOL_PROSENT] = pd.Series(dtype=float)
        df_hoved.loc[gyldig, KOL_PROSENT] = (omega_serie[gyldig] / mengde_serie[gyldig] * 100).round(0)

        # Advarsel-kolonne: Omega-mengde overskrider APM-mengde
        omega_over_apm = gyldig & (omega_serie > mengde_serie)

        antall_advarsler = omega_over_apm.sum()
        print(f"  Omega koblinger: {df_hoved[KOL_OMEGA].notna().sum()} treff, "
              f"{ikke_i_omega.sum()} mangler i Omega, "
              f"{antall_advarsler} advarsler (Omega > APM)")

    # Armering totalmengde: aktive kg-rader uten staal-referanse
    er_aktiv_kg = (
        (df_hoved["Utgatt"].str.lower() != "ja")
        & (df_hoved["Enhet fra APM"].str.strip().str.lower() == "kg")
        & ~df_hoved["Ekstra inndeling"].str.contains("staal|st\u00e5l", case=False, na=False, regex=True)
        & ~df_hoved["Kommentar"].str.contains("staal|st\u00e5l", case=False, na=False, regex=True)
    )
    df_arm_total = df_hoved[er_aktiv_kg].reset_index(drop=True)

    # Marker rader der samme prosesskode finnes i mer enn en APM (checkmark-kolonne)
    post_apm_teller = df_hoved.groupby("Post")["APM"].transform("nunique")
    df_hoved["Flere APMer"] = post_apm_teller.map(lambda n: "\u2713" if n > 1 else "")

    df_post_samm     = _bygg_post_sammenstilling(df_hoved, omega_data)
    df_arm_total_samm = _bygg_post_sammenstilling(df_arm_total, omega_data)

    # Omega uten APM: Omega-poster med delmaalebrev = naeste maaned uten tilhorende APM
    df_omega_uten_apm = pd.DataFrame()
    df_mangler_samm   = pd.DataFrame()
    if omega_data:
        maal_maaned = _dato_til_maaned_aar(datetime.datetime.today())
        apm_poster = set(df_hoved["Post"].dropna().unique())
        rader_omega_mangler = []
        for prosesskode, info in sorted(omega_data.items()):
            if prosesskode in apm_poster:
                continue
            if _maalebrev_maaned(info.get("delmaal"), info.get("boqa")) != maal_maaned:
                continue
            rader_omega_mangler.append({
                "Prosesskode":   prosesskode,
                KOL_OMEGA:       info.get("mengde"),
                KOL_KONTRAKT:    info.get("kontrakt"),
                KOL_ENHETSPRIS:  info.get("pris"),
                "Enhet":         info.get("enhet"),
                KOL_DELMAAL:     maal_maaned,
                KOL_SRES:        _sres_til_tekst(info.get("sres")),
            })
        if rader_omega_mangler:
            df_omega_uten_apm = pd.DataFrame(rader_omega_mangler, columns=KOLONNER_OMEGA_UTEN_APM)
        print(f"Omega uten APM ({maal_maaned}): {len(rader_omega_mangler)} poster uten tilhorende APM")

        # Mangler i Post sammenstilling: alle Omega-poster for gjeldende maaned (unntatt 01.*)
        # som ikke finnes i Post sammenstilling (superset av Omega uten APM)
        post_samm_poster = set(df_post_samm["Post"].dropna().unique()) if not df_post_samm.empty else set()
        rader_mangler_samm = []
        for prosesskode, info in sorted(omega_data.items()):
            if prosesskode.startswith("01.") or prosesskode == "01":
                continue
            if _maalebrev_maaned(info.get("delmaal"), info.get("boqa")) != maal_maaned:
                continue
            if prosesskode in post_samm_poster:
                continue
            if prosesskode in apm_poster:
                merknad = "APM utgatt - ikke i erstatter-APM"
            else:
                merknad = "Ingen APM levert"
            rader_mangler_samm.append({
                "Prosesskode":  prosesskode,
                KOL_OMEGA:      info.get("mengde"),
                KOL_KONTRAKT:   info.get("kontrakt"),
                KOL_ENHETSPRIS: info.get("pris"),
                "Enhet":        info.get("enhet"),
                KOL_DELMAAL:    maal_maaned,
                KOL_SRES:       _sres_til_tekst(info.get("sres")),
                "Merknad":      merknad,
                KOL_BOQA_KOMM:  info.get("boqa") or "",
            })
        if rader_mangler_samm:
            df_mangler_samm = pd.DataFrame(rader_mangler_samm, columns=KOLONNER_MANGLER_SAMM)
        print(f"Mangler i Post samm. ({maal_maaned}): {len(rader_mangler_samm)} poster ikke i Post sammenstilling")

    print(f"\nHoveddatatabell:      {len(df_hoved)} rader")
    print(f"Armeringsdata:        {len(df_arm)} rader")
    print(f"Armering totalmengde: {len(df_arm_total_samm)} rader (sammenstilt)")
    print(f"Post sammenstilling:  {len(df_post_samm)} rader")

    # APM betong/forskaling-sjekk (krev ISY Linker)
    apm_bf_rader: list[dict] = []
    if isy_path:
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from mengde_oversikt import les_isy
            print("Bygger APM betong/forskaling-sjekk …")
            isy_map = les_isy(isy_path)
            apm_bf_rader = bygg_apm_betong_forskaling(df_hoved, omega_data, isy_map)
            print(f"APM Rapp kontroll:    {len(apm_bf_rader)} vegger")
        except Exception as _e:
            print(f"[WARNING] APM betong/forskaling-sjekk feilet: {_e}")

    subfolder_namn = os.path.basename(os.path.normpath(mappe))
    _skriv_excel(df_hoved, df_arm, df_arm_total_samm, df_post_samm, df_omega_uten_apm,
                 df_mangler_samm, output_path, alle_stier, omega_path, subfolder_namn,
                 apm_bf_rader=apm_bf_rader)
    try:
        _skriv_rettede_ark(output_path, rettede_rader)
    except Exception as exc:
        print(f"  [ADVARSEL] Klarte ikkje skrive arket '{ARK_RETTEDE}': {exc}")
    print(f"[SUCCESS] Lagret: {output_path}")
    return output_path


# ===== Excel-skriving =====

def _bygg_post_sammenstilling(df_hoved: pd.DataFrame, omega_data: dict = None) -> pd.DataFrame:
    """Bygg Post-sammenstilling fra aktive (ikke-utgaatte) Malebrev-rader.

    Samme kolonner som Malebrev, sortert paa Post -> APM.
    - Enkelt APM per post: Omega/prosent vises direkte paa den raden.
    - Flere APMer per post: Omega/prosent vises paa en [TOTAL]-rad etter gruppen.

    KOL_OMEGA_ENHET: viser enhetskoden fra Omega slik at brukeren kan sammenligne direkte.
    """
    if omega_data is None:
        omega_data = {}

    df_aktiv = df_hoved[df_hoved["Utgatt"].str.lower() != "ja"].copy()
    if df_aktiv.empty:
        return pd.DataFrame()

    df_aktiv["_mengde_num"] = pd.to_numeric(df_aktiv["Mengde fra APM"], errors="coerce")
    df_aktiv = df_aktiv.sort_values(["Post", "APM"])

    rader = []
    for post, gruppe in df_aktiv.groupby("Post", sort=True):
        # Poster som mangler i Omega vises uten Omega-data (ikke hoppet over)

        # Beregn felles Omega-info for hele Post-gruppen
        omega_val = None
        omg_ser = gruppe[KOL_OMEGA].dropna()
        if not omg_ser.empty:
            omega_val = omg_ser.iloc[0]

        ikke_i_omega = "!" if gruppe[KOL_MANGLER].eq("!").any() else ""

        total = gruppe["_mengde_num"].sum(min_count=1)
        total = float(total) if pd.notna(total) else None

        prosent = None
        if total and omega_val is not None and total != 0:
            try:
                prosent = round(float(omega_val) / total * 100, 0)
            except (TypeError, ValueError):
                pass

        utrop = "!" if (prosent is not None and prosent > 100.5) else ""
        enhet = next((str(v).strip() for v in gruppe["Enhet fra APM"] if v and str(v).strip()), "")
        har_flere_apm = gruppe["APM"].nunique() > 1

        # Omega-enhet og kontraktsmengde for posten (lik for alle APMer paa same post)
        omega_enhet_raa  = (omega_data.get(post) or {}).get("enhet") or ""
        omega_enhet_vis  = _normaliser_enhet(omega_enhet_raa)   # viser ASCII-variant (m3, m2 osv.)
        omega_kontrakt   = (omega_data.get(post) or {}).get("kontrakt")
        omega_enhetspris = (omega_data.get(post) or {}).get("pris")
        omega_boqa       = (omega_data.get(post) or {}).get("boqa") or ""
        omega_delmaal    = _maalebrev_maaned((omega_data.get(post) or {}).get("delmaal"), omega_boqa)
        omega_sres       = _sres_til_tekst((omega_data.get(post) or {}).get("sres"))

        # Samle alle kommentarer fra APM-rader (brukes paa [TOTAL]-raden)
        alle_kommentarer = "; ".join(
            str(row["Kommentar"]).strip()
            for _, row in gruppe.iterrows()
            if str(row["Kommentar"]).strip()
        )

        # Samle unike Ekstra inndeling-verdier til [TOTAL]-raden
        alle_ekstra = "; ".join(
            dict.fromkeys(  # bevarer rekkefolgje, fjerner duplikater
                v for v in (str(row["Ekstra inndeling"]).strip() for _, row in gruppe.iterrows())
                if v
            )
        )

        # Enhetsavvik for hele posten
        enhet_feil_post = bool(
            omega_enhet_vis
            and _normaliser_enhet(enhet)
            and _normaliser_enhet(enhet) != omega_enhet_vis
        )

        # -- Individuelle APM-rader --
        noen_enhet_feil = False   # blir True dersom minst én rad i gruppen har enhetsavvik
        for _, row in gruppe.iterrows():
            rad_enhet_feil = bool(
                omega_enhet_vis
                and _normaliser_enhet(str(row["Enhet fra APM"]))
                and _normaliser_enhet(str(row["Enhet fra APM"])) != omega_enhet_vis
            )
            # Ved kun en APM: legg Omega/prosent direkte paa raden
            if not har_flere_apm:
                rad_omega    = omega_val
                rad_prosent  = prosent
                rad_mangler  = ikke_i_omega
                if prosent == 100.0 and not rad_enhet_feil:
                    rad_utrop = "\u2713"        # checkmark: mål nådd
                elif utrop == "!" or rad_enhet_feil:
                    rad_utrop = "!"
                else:
                    rad_utrop = ""
            else:
                rad_omega    = None
                rad_prosent  = None
                rad_mangler  = ""
                rad_utrop    = "!" if rad_enhet_feil else ""

            if rad_enhet_feil:
                noen_enhet_feil = True

            rader.append({
                "Post": row["Post"],
                "Ekstra inndeling": row["Ekstra inndeling"],
                "APM": row["APM"],
                "Mengde fra APM": row["_mengde_num"],
                "Enhet fra APM": row["Enhet fra APM"],
                KOL_OMEGA_ENHET: omega_enhet_vis,
                "Kommentar": row["Kommentar"],
                KOL_BOQA_KOMM: omega_boqa,
                "Flere APMer": "X" if har_flere_apm else "",
                KOL_OMEGA: rad_omega,
                KOL_KONTRAKT: omega_kontrakt,
                KOL_ENHETSPRIS: omega_enhetspris,
                KOL_PROSENT: rad_prosent,
                KOL_FORBRUK: None if rad_prosent == 100.0 else rad_prosent,
                "!": rad_utrop,
                KOL_DELMAAL: omega_delmaal,
                KOL_SRES: omega_sres,
            })

        # -- TOTAL-rad: kun naar posten har mer enn en APM --
        if har_flere_apm:
            rader.append({
                "Post": post,
                "Ekstra inndeling": alle_ekstra,
                "APM": "[TOTAL]",
                "Mengde fra APM": total,
                "Enhet fra APM": enhet,
                KOL_OMEGA_ENHET: omega_enhet_vis,
                "Kommentar": alle_kommentarer,
                KOL_BOQA_KOMM: omega_boqa,
                "Flere APMer": "[TOTAL]",
                KOL_OMEGA: omega_val,
                KOL_KONTRAKT: omega_kontrakt,
                KOL_ENHETSPRIS: omega_enhetspris,
                KOL_PROSENT: prosent,
                KOL_FORBRUK: None if prosent == 100.0 else prosent,
                "!": "\u2713" if (prosent == 100.0 and not noen_enhet_feil) else ("!" if (utrop == "!" or noen_enhet_feil) else ""),
                KOL_DELMAAL: omega_delmaal,
                KOL_SRES: omega_sres,
            })

    return pd.DataFrame(rader, columns=KOLONNER_SAMM)


def _legg_til_forbruk_cf(ark, forbruk_kol_nr: int, maks_rad: int):
    """Legg til ColorScaleRule (grønn-gul-rød) på forbruk-kolonnen, med grønn override ved nøyaktig 100 %."""
    if not forbruk_kol_nr or maks_rad < 2:
        return
    fra     = f"{get_column_letter(forbruk_kol_nr)}2"
    til     = f"{get_column_letter(forbruk_kol_nr)}{maks_rad}"
    omraade = f"{fra}:{til}"
    # Grønn override for nøyaktig 100 % (prioritet 1 = høyest)
    regel_100 = CellIsRule(operator="equal", formula=["100"], fill=_FILL_SAMM_100)
    regel_100.priority = 1
    ark.conditional_formatting.add(omraade, regel_100)
    # Fargegradering: grønn (0) -> gul (80) -> rød (100)
    regel_skala = ColorScaleRule(
        start_type="num", start_value=0,   start_color="63BE7B",
        mid_type="num",   mid_value=80,    mid_color="FFEB84",
        end_type="num",   end_value=100,   end_color="F8696B",
    )
    regel_skala.priority = 2
    ark.conditional_formatting.add(omraade, regel_skala)


def _legg_til_header_kommentarer(ark, kommentar_dict: dict):
    """Legg til kommentarer (tooltips) på header-celler."""
    for celle in ark[1]:
        tekst = kommentar_dict.get(celle.value)
        if tekst:
            celle.comment = Comment(tekst, "SSA")


def _skriv_oversikt_ark(wb, df_hoved: pd.DataFrame, df_post_samm: pd.DataFrame, omega_path: str):
    """Skriv Forklaring og Sammendrag-ark til wb (settes inn som første ark)."""
    ark = wb.create_sheet("Forklaring og Sammendrag", 0)

    _T      = Font(bold=True, color="FFFFFF", name="Calibri", size=12)
    _H      = Font(bold=True, name="Calibri", size=10)
    _N      = Font(name="Calibri", size=10)
    _FT     = PatternFill("solid", fgColor="1B4AA3")   # blå seksjonstittel
    _FT_GRN = PatternFill("solid", fgColor="2E7D32")   # mørk grønn - sammendrag-header
    _FS     = PatternFill("solid", fgColor="EEF2FA")   # lys blå - annenhver forklaringsrad
    _FS_GRN = PatternFill("solid", fgColor="E8F5E9")   # lys grønn - alle sammendrag-rader

    def seksjonstitel(r, tekst, header_fyll=None):
        c = ark.cell(r, 1, tekst)
        c.fill = header_fyll if header_fyll is not None else _FT
        c.font = _T
        c.alignment = Alignment(horizontal="left", vertical="center")
        ark.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ark.row_dimensions[r].height = 22

    def verdirad(r, kol1, kol2="", rad_fyll=None):
        c1 = ark.cell(r, 1, kol1)
        c1.font = _H
        c1.alignment = Alignment(vertical="top", wrap_text=True)
        c2 = ark.cell(r, 2, kol2)
        c2.font = _N
        c2.alignment = Alignment(vertical="top", wrap_text=True)
        if rad_fyll is not None:
            c1.fill = rad_fyll
            c2.fill = rad_fyll
        elif r % 2 == 0:
            c1.fill = _FS
            c2.fill = _FS

    r = 1

    # -- Sammendrag (grønn seksjon) --
    seksjonstitel(r, "SAMMENDRAG", _FT_GRN); r += 1
    antall_apm     = df_hoved["APM"].nunique() if "APM" in df_hoved.columns else 0
    antall_rader   = len(df_hoved)
    antall_mangler = int((df_hoved[KOL_MANGLER] == "!").sum()) if KOL_MANGLER in df_hoved.columns else 0
    antall_adv     = int((df_post_samm["!"] == "!").sum()) if not df_post_samm.empty and "!" in df_post_samm.columns else 0
    omega_filnavn  = os.path.basename(omega_path) if omega_path else "(ikke valgt)"
    dato           = datetime.date.today().isoformat()
    for kol1, kol2 in [
        ("Generert dato",     dato),
        ("Antall APMer",      str(antall_apm)),
        ("Antall posterader", str(antall_rader)),
        ("Mangler i Omega",   str(antall_mangler)),
        ("Omega-fil",         omega_filnavn),
    ]:
        verdirad(r, kol1, kol2, _FS_GRN); r += 1
    # Advarsler-rad med hyperlenke til Post sammenstilling
    verdirad(r, "Advarsler (klikk for å åpne)", str(antall_adv), _FS_GRN)
    c_adv = ark.cell(r, 2)
    c_adv.hyperlink = "#'Post sammenstilling'!A1"
    c_adv.font = Font(
        bold=True,
        color="C00000" if antall_adv > 0 else "1A6B1A",
        underline="single", name="Calibri", size=10,
    )
    r += 1

    r += 1
    # -- Ark-oversikt --
    seksjonstitel(r, "ARK-OVERSIKT"); r += 1
    for namn, beskrivelse in [
        ("Forklaring og Sammendrag", "Dette arket. Oversikt over ark, kolonner, farger og symboler, pluss nøkkeltall."),
        ("APM sammenstilling",       "Alle postlinjer fra alle APM-PDFer. Én rad per post per APM."),
        ("Post sammenstilling",      "Aggregert per post: individuelle APM-rader + [TOTAL]-linje for poster med flere APMer."),
        ("Armering totalmengde",     "Aktive kg-poster (uten stål-referanse), aggregert per post og APM."),
        ("Armeringsmengder",         "Formkode-tabeller med modellvekt og omfar-faktor fra PDFene."),
        ("Mangler i Post samm.",     "Omega-poster for gjeldende maaned som mangler i Post sammenstilling (unntatt 01.*). Gul = ingen APM levert, rosa = APM utgatt men erstatter dekker ikke posten."),
    ]:
        verdirad(r, namn, beskrivelse); r += 1

    r += 1
    # -- Kolonneforklaring --
    seksjonstitel(r, "KOLONNEFORKLARING"); r += 1
    for kol_namn, forklaring in [
        ("PNS",              "Prosjektnedbrytningsstruktur - hierarkisk kode fra APM-dokumentet."),
        ("Ekstra inndeling", "Valgfri ekstra kategorisering fra APM (f.eks. fagområde eller del-område)."),
        ("APM",              "Navn på APM-PDF-filen (uten .pdf-endelse). Klikk for å åpne filen (hyperlenke)."),
        ("Post",             "Prosesskode (normalisert format, f.eks. 1303.01.06.241.1.1)."),
        ("Mengde fra APM",  "Målt mengde fra APM-dokumentet."),
        ("Enhet fra APM",   "Enhetskode fra APM (m, m2, m3, stk, kg, lm osv.)."),
        ("EO",               "Endringsordre-nummer fra APM-dokumentet."),
        ("Kommentar",        "Fritekst-kommentar fra APM-dokumentet."),
        ("Utgatt",           "'Ja' hvis APM-filen er erstattet av en nyere versjon."),
        ("Ny_APM",           "Navn på APM-filen som erstatter denne versjonen (hyperlenke)."),
        ("Vedlegg",          "Andre filer i mappen med samme base-navn som APM-filen."),
        (KOL_MANGLER,        "Utropstegn (!) hvis posten ikke finnes i Omega overhodet."),
        ("Flere APMer",       "Checkmark (\u2713) i APM sammenstilling: samme prosesskode finnes i mer enn én APM. 'X' / '[TOTAL]' i Post sammenstilling: individuell / samlelinje for poster med flere APMer."),
        ("!",                "Utropstegn (!) i Post sammenstilling / Armering totalmengde: Omega overskrider APM-mengde (> 100 %). Checkmark (\u2713) = nøyaktig 100 %."),
        (KOL_OMEGA,          "Rapportert kumulativ mengde fra Omega BOQNT-arket."),
        (KOL_KONTRAKT,       "Kontraktsmengde fra Omega BOQNT-arket."),
        (KOL_ENHETSPRIS,     "Enhetspris fra Omega BOQNT-arket."),
        (KOL_PROSENT,        "Prosentandel: Omega rapportert mengde / APM-mengde * 100."),
        (KOL_FORBRUK,        "Fargekodet visualisering av forbruksprosent: grønn < 80 %, gul 80-100 %, rød > 100 %."),
        (KOL_OMEGA_ENHET,    "Enhetskode fra Omega for sammenligning med APM-enhet. Lilla bakgrunn = avvik."),
        ("Formkode",         "Formkode for armeringsjern fra 3D-modellen."),
        ("kg_i_modell",      "Modellvekt fra 3D-modellen (kg), uten omfar-faktor."),
        ("Faktor_omfar",     "Faktor for omfar (påslagsfaktor for armering, typisk 1.0-1.15)."),
        ("Vekt_inkl_omfar",  "Beregnet vekt inkl. omfar-faktor (kg). Beregnes automatisk dersom PDF-kolonnen er tom."),
        ("Flere APMer",      "'X' hvis det er flere APMer på posten; '[TOTAL]' på samlelinjen."),
    ]:
        verdirad(r, kol_namn, forklaring); r += 1

    r += 1
    # -- Fargeforklaring --
    seksjonstitel(r, "FARGEFORKLARING"); r += 1
    for fyll, farge_namn, beskrivelse in [
        (_HDR_FILL,             "Mørk blå header",              "Kolonneoverskrifter i alle ark."),
        (_FILL_UTGAATT,         "Lys rød",                      "Utgått APM (erstattet av nyere versjon)."),
        (_FILL_MANGLER_OMEGA,   "Lys gul",                      "Posten mangler i Omega (ikke registrert overhodet)."),
        (_FILL_OMEGA_ADV,       "Rød",                          "Omega rapportert mengde overskrider APM-mengde."),
        (_FILL_SAMM_TOT,        "Mørk stålblå (hvit tekst)",    "[TOTAL]-rad: sum av alle APMer på posten."),
        (_FILL_SAMM_FLERE,      "Mellomblå (kursiv grå tekst)", "Individuelle APM-rader på poster med flere APMer."),
        (_FILL_SAMM_ENHET_FEIL, "Lys lilla",                    "Enhet i APM stemmer ikke med Omega-enhet."),
        (_FILL_SAMM_100,        "Lys grønn (hel rad)",           "Prosent er nøyaktig 100 % - målet er nådd."),
        (_FILL_PROSENT_GROENN,  "Lys grønn",                    "Forbruk < 80 % - godt margin mot kontraktsmengde."),
        (_FILL_PROSENT_GUL,     "Lys gul",                      "Forbruk 80-100 % - nærmer seg kontraktsmengde."),
        (_FILL_PROSENT_ROD,     "Rød-rosa",                     "Forbruk > 100 % - overskrider kontraktsmengde."),
        (_FILL_ARM,             "Lys grønn",                    "Annenhver rad i Armeringsmengder-arket."),
    ]:
        c_farge = ark.cell(r, 1, farge_namn)
        c_farge.fill = fyll
        c_farge.font = Font(name="Calibri", size=10,
                            color="FFFFFF" if fyll in (_HDR_FILL, _FILL_SAMM_TOT) else "000000")
        c_farge.alignment = Alignment(vertical="top")
        c_besk = ark.cell(r, 2, beskrivelse)
        c_besk.font = _N
        c_besk.alignment = Alignment(vertical="top", wrap_text=True)
        r += 1

    r += 1
    # -- Symbolforklaring --
    seksjonstitel(r, "SYMBOLFORKLARING"); r += 1
    for symbol, forklaring in [
        ("! i kolonne 'Ikke i Omega'", "Posten (prosesskoden) finnes ikke i Omega overhodet."),
        ("! i kolonne '!'",            "Omega rapportert mengde er større enn APM-mengde (forbruk > 100 %)."),
        ("[TOTAL]",                    "Samlelinjen for alle APMer på en post (Post sammenstilling / Armering totalmengde)."),
        ("[IDENTISK KOPI]",            "PDF-filen er en eksakt kopi (MD5-hash) av en annen fil og ble hoppet over."),
        ("ERSTATTER APMxx",            "Tekst i PDF-en som angir at denne APM-en erstatter en eldre APM."),
        ("Utgatt = Ja",                "APM-en er erstattet; fremdeles inkludert i dataene, men markert med lys rød bakgrunn."),
    ]:
        verdirad(r, symbol, forklaring); r += 1

    ark.column_dimensions["A"].width = 46
    ark.column_dimensions["B"].width = 72
    ark.sheet_view.showGridLines = False


def _sett_skillelinje(ark, kol_venstre: int, kol_hoyre: int):
    """Sett tykk vertikal skillelinje mellom to kolonner for alle rader inkl. header."""
    if not kol_venstre or not kol_hoyre:
        return
    for rad_nr in range(1, ark.max_row + 1):
        c_v = ark.cell(rad_nr, kol_venstre)
        c_h = ark.cell(rad_nr, kol_hoyre)
        b_v = c_v.border
        b_h = c_h.border
        c_v.border = Border(left=b_v.left, right=_KANT_TYKK, top=b_v.top, bottom=b_v.bottom)
        c_h.border = Border(left=_KANT_TYKK, right=b_h.right, top=b_h.top, bottom=b_h.bottom)


def _sett_hyperlenke(celle, pdf_stier: dict, subfolder_namn: str = ""):
    """Legg til fil-hyperlenke paa en celle hvis APM-navnet finnes i pdf_stier.
    Bruker relativ sti slik at lenken fungerer uavhengig av hvor filene er
    plassert. Naar Excel-filen ligger ett nivaa over PDF-ene brukes
    'subfoldernavn/filnavn.pdf' som relativ sti."""
    navn = str(celle.value or "")
    if navn and navn in pdf_stier:
        filnavn = os.path.basename(pdf_stier[navn])
        rel_sti = (subfolder_namn + "/" + filnavn) if subfolder_namn else filnavn
        celle.hyperlink = rel_sti
        celle.font = Font(color="1B4AA3", underline="single", name="Calibri", size=10)


def _escape_formler(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix formel-tegn med apostrof saa Excel ikke tolker dem som formler.
    Erstatter ogsaa tomme strenger med None slik at Excel-celler blir virkelig tomme."""
    def esc(verdi):
        if verdi == "":
            return None
        return "'" + verdi if isinstance(verdi, str) and verdi[:1] in "=+-@" else verdi
    return df.apply(lambda kol: kol.map(esc))


def _formater_header(ark, header_fill):
    """Sett header-stil paa foerste rad i arket."""
    for celle in ark[1]:
        celle.fill = header_fill
        celle.font = _HDR_FONT
        celle.alignment = _HDR_ALIGN
        celle.border = _ny_border()
    ark.row_dimensions[1].height = 28


def _sett_kolonnebredder(ark, kolonner: list, bredder: dict, med_autofilter: bool = True):
    """Sett kolonnebredder basert paa bredde-mapping.
    med_autofilter=False brukes paa ark som har en Excel-tabell (table),
    da tabellen har sin egen innebygget autofilter og to autofilter paa
    samme ark gjor filen korrupt."""
    for idx, kol_navn in enumerate(kolonner, 1):
        ark.column_dimensions[get_column_letter(idx)].width = bredder.get(kol_navn, 15)
    ark.freeze_panes = "A2"
    if med_autofilter:
        ark.auto_filter.ref = ark.dimensions


def _formater_tallcelle(ark, rad_nr: int, kol_nr: int, fmt: str):
    """Sett tallformat paa en celle hvis verdien er numerisk."""
    if not kol_nr:
        return
    celle = ark.cell(rad_nr, kol_nr)
    if isinstance(celle.value, (int, float)):
        celle.number_format = fmt


def _skriv_excel(df_hoved: pd.DataFrame, df_arm: pd.DataFrame,
                 df_arm_total_samm: pd.DataFrame, df_post_samm: pd.DataFrame,
                 df_omega_uten_apm: pd.DataFrame, df_mangler_samm: pd.DataFrame,
                 sti: str, pdf_stier: dict, omega_path: str = None,
                 subfolder_namn: str = "",
                 apm_bf_rader=None):
    """Skriv alle ark til Excel og formater dem."""

    from lagring import vent_til_skrivbar
    vent_til_skrivbar(sti)

    # Skriv raadata
    with pd.ExcelWriter(sti, engine="openpyxl") as writer:
        _escape_formler(df_hoved).to_excel(writer, index=False, sheet_name="APM sammenstilling")
        _escape_formler(df_arm).to_excel(writer, index=False, sheet_name="Armeringsmengder")
        if not df_arm_total_samm.empty:
            _escape_formler(df_arm_total_samm).to_excel(writer, index=False, sheet_name="Armering totalmengde")
        if not df_post_samm.empty:
            _escape_formler(df_post_samm).to_excel(writer, index=False, sheet_name="Post sammenstilling")
        if not df_omega_uten_apm.empty:
            _escape_formler(df_omega_uten_apm).to_excel(writer, index=False, sheet_name="Omega uten APM")
        if not df_mangler_samm.empty:
            _escape_formler(df_mangler_samm).to_excel(writer, index=False, sheet_name="Mangler i Post samm.")

    wb = load_workbook(sti)

    # -- Ark 1: APM sammenstilling --
    ark_mal = wb["APM sammenstilling"]
    _formater_header(ark_mal, _HDR_FILL)

    kol = {celle.value: celle.column for celle in ark_mal[1]}
    apm_kol        = kol.get("APM")
    ny_apm_kol     = kol.get("Ny_APM")
    utgaatt_kol    = kol.get("Utgatt")
    mengde_kol     = kol.get("Mengde fra APM")
    mangler_kol    = kol.get(KOL_MANGLER)
    omega_kol      = kol.get(KOL_OMEGA)
    kontrakt_kol   = kol.get(KOL_KONTRAKT)
    enhetspris_kol = kol.get(KOL_ENHETSPRIS)
    prosent_kol    = kol.get(KOL_PROSENT)

    # Bygg APM -> unik farge-mapping (roterer gjennom _APM_FARGER etter foerste-forekomst)
    apm_farge_map: dict = {}
    if apm_kol:
        for r in range(2, ark_mal.max_row + 1):
            apm_navn = str(ark_mal.cell(r, apm_kol).value or "").strip()
            if apm_navn and apm_navn not in apm_farge_map:
                apm_farge_map[apm_navn] = _APM_FARGER[len(apm_farge_map) % len(_APM_FARGER)]

    for rad_nr in range(2, ark_mal.max_row + 1):
        er_utgaatt       = str(ark_mal.cell(rad_nr, utgaatt_kol or 1).value or "").lower() == "ja"
        er_ikke_i_omega  = mangler_kol and str(ark_mal.cell(rad_nr, mangler_kol).value or "") == "!"
        apm_verdi_mal    = str(ark_mal.cell(rad_nr, apm_kol).value or "").strip() if apm_kol else ""
        if er_utgaatt:
            fyll = _FILL_UTGAATT
        elif er_ikke_i_omega:
            fyll = _FILL_MANGLER_OMEGA
        else:
            fyll = apm_farge_map.get(apm_verdi_mal)
        for celle in ark_mal[rad_nr]:
            if fyll:
                celle.fill = fyll
            celle.alignment = _VTOPP
            celle.border = _ny_border()
            celle.font = _FONT_DATA
        _formater_tallcelle(ark_mal, rad_nr, mengde_kol,     '#,##0.00')
        _formater_tallcelle(ark_mal, rad_nr, omega_kol,      '#,##0.00')
        _formater_tallcelle(ark_mal, rad_nr, kontrakt_kol,   '#,##0.00')
        _formater_tallcelle(ark_mal, rad_nr, enhetspris_kol, '#,##0.00')
        _formater_tallcelle(ark_mal, rad_nr, prosent_kol,    '0"%"')
        if apm_kol:
            _sett_hyperlenke(ark_mal.cell(rad_nr, apm_kol), pdf_stier, subfolder_namn)
        if ny_apm_kol:
            _sett_hyperlenke(ark_mal.cell(rad_nr, ny_apm_kol), pdf_stier, subfolder_namn)

    _sett_kolonnebredder(ark_mal, KOLONNER_HOVED, BREDDER_HOVED)
    _legg_til_header_kommentarer(ark_mal, _HDR_KOMMENTARER)

    # -- Ark 2: Armeringsmengder --
    ark_arm = wb["Armeringsmengder"]
    _formater_header(ark_arm, _FILL_ARM_HDR)

    kol_arm = {celle.value: celle.column for celle in ark_arm[1]}
    apm_kol_arm = kol_arm.get("APM")
    tall_kolonner = [kol_arm.get(navn) for navn in ("kg_i_modell", "Faktor_omfar", "Vekt_inkl_omfar")
                     if kol_arm.get(navn)]

    for rad_nr in range(2, ark_arm.max_row + 1):
        fyll = _FILL_ARM if rad_nr % 2 == 0 else None
        for celle in ark_arm[rad_nr]:
            if fyll:
                celle.fill = fyll
            celle.alignment = _VTOPP
            celle.border = _ny_border()
            celle.font = _FONT_DATA
        for tall_kol in tall_kolonner:
            _formater_tallcelle(ark_arm, rad_nr, tall_kol, '#,##0.00')
        if apm_kol_arm:
            _sett_hyperlenke(ark_arm.cell(rad_nr, apm_kol_arm), pdf_stier, subfolder_namn)

    _sett_kolonnebredder(ark_arm, KOLONNER_ARM, BREDDER_ARM)

    # -- Ark 3: Armering totalmengde (sammenstilt format) --
    if "Armering totalmengde" in wb.sheetnames:
        ark_tot = wb["Armering totalmengde"]
        _formater_header(ark_tot, _FILL_TOT_HDR)

        kol_tot = {celle.value: celle.column for celle in ark_tot[1]}
        apm_kol_tot      = kol_tot.get("APM")
        mengde_kol_tot   = kol_tot.get("Mengde fra APM")
        omega_kol_tot    = kol_tot.get(KOL_OMEGA)
        prosent_kol_tot  = kol_tot.get(KOL_PROSENT)
        forbruk_kol_tot  = kol_tot.get(KOL_FORBRUK)
        mangler_kol_tot  = kol_tot.get(KOL_MANGLER)
        utrop_kol_tot    = kol_tot.get("!")

        for rad_nr in range(2, ark_tot.max_row + 1):
            apm_verdi   = str(ark_tot.cell(rad_nr, apm_kol_tot).value or "") if apm_kol_tot else ""
            er_total    = apm_verdi == "[TOTAL]"
            er_over_100 = utrop_kol_tot and str(ark_tot.cell(rad_nr, utrop_kol_tot).value or "") == "!"
            er_mangler  = mangler_kol_tot and str(ark_tot.cell(rad_nr, mangler_kol_tot).value or "") == "!"

            if er_total and er_over_100:
                fyll = _FILL_SAMM_OVER
            elif er_total and er_mangler:
                fyll = _FILL_MANGLER_OMEGA
            elif er_total:
                fyll = _FILL_SAMM_TOT
            elif er_over_100:
                fyll = _FILL_SAMM_OVER
            elif er_mangler:
                fyll = _FILL_MANGLER_OMEGA
            else:
                fyll = _FILL_TOT_RAD if rad_nr % 2 == 0 else None

            for celle in ark_tot[rad_nr]:
                if fyll:
                    celle.fill = fyll
                celle.alignment = _VTOPP
                celle.border = _ny_border()
                if er_total:
                    celle.font = _FONT_TOT_ARM
                else:
                    celle.font = _FONT_DATA

            _formater_tallcelle(ark_tot, rad_nr, mengde_kol_tot,  '#,##0.00')
            _formater_tallcelle(ark_tot, rad_nr, omega_kol_tot,   '#,##0.00')
            _formater_tallcelle(ark_tot, rad_nr, prosent_kol_tot, '0"%"')
            if forbruk_kol_tot:
                ark_tot.cell(rad_nr, forbruk_kol_tot).number_format = ';;;'

            if not er_total and apm_kol_tot:
                _sett_hyperlenke(ark_tot.cell(rad_nr, apm_kol_tot), pdf_stier, subfolder_namn)

        _sett_kolonnebredder(ark_tot, KOLONNER_SAMM, BREDDER_SAMM)
        _legg_til_forbruk_cf(ark_tot, forbruk_kol_tot, ark_tot.max_row)
        ark_tot.freeze_panes = "C2"

    # -- Ark 4: Post sammenstilling --
    if "Post sammenstilling" in wb.sheetnames:
        ark_samm = wb["Post sammenstilling"]
        _formater_header(ark_samm, _FILL_SAMM_HDR)

        kol_samm           = {celle.value: celle.column for celle in ark_samm[1]}
        apm_kol_samm       = kol_samm.get("APM")
        mengde_kol_samm    = kol_samm.get("Mengde fra APM")
        enhet_kol_samm     = kol_samm.get("Enhet fra APM")
        omega_enhet_kol_samm = kol_samm.get(KOL_OMEGA_ENHET)
        omega_kol_samm     = kol_samm.get(KOL_OMEGA)
        kontrakt_kol_samm    = kol_samm.get(KOL_KONTRAKT)
        enhetspris_kol_samm  = kol_samm.get(KOL_ENHETSPRIS)
        prosent_kol_samm     = kol_samm.get(KOL_PROSENT)
        forbruk_kol_samm     = kol_samm.get(KOL_FORBRUK)
        utrop_kol_samm     = kol_samm.get("!")
        flere_apm_kol_samm = kol_samm.get("Flere APMer")

        for rad_nr in range(2, ark_samm.max_row + 1):
            apm_verdi    = str(ark_samm.cell(rad_nr, apm_kol_samm).value or "") if apm_kol_samm else ""
            er_total     = apm_verdi == "[TOTAL]"
            har_flere    = flere_apm_kol_samm and str(ark_samm.cell(rad_nr, flere_apm_kol_samm).value or "") == "X"

            # Farge basert paa prosent-verdien direkte
            prosent_verdi = ark_samm.cell(rad_nr, prosent_kol_samm).value if prosent_kol_samm else None
            try:
                _p = float(prosent_verdi) if prosent_verdi is not None else None
                er_over_100    = _p is not None and _p > 100
                er_akkurat_100 = _p is not None and _p == 100.0
            except (TypeError, ValueError):
                er_over_100    = False
                er_akkurat_100 = False

            # Enhetssammenligning (normaliser begge sider slik at m2 == m2, m3 == m3 osv.)
            enhet_pdf   = _normaliser_enhet(ark_samm.cell(rad_nr, enhet_kol_samm).value or "") if enhet_kol_samm else ""
            enhet_omega = _normaliser_enhet(ark_samm.cell(rad_nr, omega_enhet_kol_samm).value or "") if omega_enhet_kol_samm else ""
            er_enhet_feil = bool(enhet_pdf and enhet_omega and enhet_pdf != enhet_omega)

            # [TOTAL]-rader: bruk "!"-kolonnen som signal for enhetsavvik (propagert fra individuelle rader)
            er_total_utrop = (
                er_total
                and utrop_kol_samm is not None
                and str(ark_samm.cell(rad_nr, utrop_kol_samm).value or "") == "!"
            )

            if er_total and er_over_100:
                fyll = _FILL_SAMM_OVER
            elif er_total and er_total_utrop:
                fyll = _FILL_SAMM_ENHET_FEIL
            elif er_total and er_akkurat_100 and not er_enhet_feil:
                fyll = _FILL_SAMM_100
            elif er_total:
                fyll = _FILL_SAMM_TOT
            elif er_over_100:
                fyll = _FILL_SAMM_OVER
            elif er_enhet_feil:
                fyll = _FILL_SAMM_ENHET_FEIL
            elif er_akkurat_100:
                fyll = _FILL_SAMM_100
            elif har_flere:
                fyll = _FILL_SAMM_FLERE
            else:
                fyll = _FILL_SAMM_RAD

            for celle in ark_samm[rad_nr]:
                if fyll:
                    celle.fill = fyll
                celle.alignment = _VTOPP
                celle.border = _ny_border_samm()
                if er_akkurat_100 and not er_enhet_feil and er_total:
                    celle.font = _FONT_100_BOLD
                elif er_akkurat_100 and not er_enhet_feil:
                    celle.font = _FONT_100
                elif er_total and er_total_utrop:
                    celle.font = _FONT_TOT_ARM   # bold svart: lesbart paa lys lilla
                elif er_total:
                    celle.font = _FONT_TOT
                elif har_flere and not er_total:
                    celle.font = _FONT_SAMM_FLERE
                else:
                    celle.font = _FONT_DATA

            _formater_tallcelle(ark_samm, rad_nr, mengde_kol_samm,     '#,##0.00')
            _formater_tallcelle(ark_samm, rad_nr, omega_kol_samm,      '#,##0.00')
            _formater_tallcelle(ark_samm, rad_nr, kontrakt_kol_samm,   '#,##0.00')
            _formater_tallcelle(ark_samm, rad_nr, enhetspris_kol_samm, '#,##0.00')
            _formater_tallcelle(ark_samm, rad_nr, prosent_kol_samm,    '0"%"')
            if forbruk_kol_samm:
                c_forbruk = ark_samm.cell(rad_nr, forbruk_kol_samm)
                c_forbruk.number_format = ';;;'
                if er_akkurat_100 and not er_enhet_feil:
                    c_forbruk.fill = _FILL_SAMM_100

            # Hyperlenke setter bla skrift paa APM-cellen (overskriver kursiv for den cellen)
            if not er_total and apm_kol_samm:
                _sett_hyperlenke(ark_samm.cell(rad_nr, apm_kol_samm), pdf_stier, subfolder_namn)

        _sett_kolonnebredder(ark_samm, KOLONNER_SAMM, BREDDER_SAMM, med_autofilter=False)
        _legg_til_forbruk_cf(ark_samm, forbruk_kol_samm, ark_samm.max_row)
        _sett_skillelinje(ark_samm, flere_apm_kol_samm, omega_kol_samm)
        ark_samm.freeze_panes = "C2"  # frys Post + Ekstra inndeling

        # Kolonnenavnet "!" er ugyldig i Excel-tabeller; gi det et gyldig navn i cellen.
        # utrop_kol_samm er allerede lagret saa dette pavirker ikke logikken ovenfor.
        if utrop_kol_samm:
            ark_samm.cell(1, utrop_kol_samm).value = "Feil enhet"

        # Register som Excel-tabell (gir filter-piler og tabellnavn)
        tab_ref = f"A1:{get_column_letter(ark_samm.max_column)}{ark_samm.max_row}"
        tab = Table(displayName="PostSammenstilling", ref=tab_ref)
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False,
        )
        ark_samm.add_table(tab)

    # -- Ark: Omega uten APM --
    _ARK_OMEGA_UTEN = "Omega uten APM"
    if _ARK_OMEGA_UTEN in wb.sheetnames:
        ark_omega = wb[_ARK_OMEGA_UTEN]
        _formater_header(ark_omega, _HDR_FILL)
        kol_o = {celle.value: celle.column for celle in ark_omega[1]}
        mengde_k_o = kol_o.get(KOL_OMEGA)
        kontr_k_o  = kol_o.get(KOL_KONTRAKT)
        pris_k_o   = kol_o.get(KOL_ENHETSPRIS)
        for rad_nr in range(2, ark_omega.max_row + 1):
            for celle in ark_omega[rad_nr]:
                celle.font = _FONT_DATA
                celle.alignment = _VTOPP
            _formater_tallcelle(ark_omega, rad_nr, mengde_k_o, '#,##0.00')
            _formater_tallcelle(ark_omega, rad_nr, kontr_k_o,  '#,##0.00')
            _formater_tallcelle(ark_omega, rad_nr, pris_k_o,   '#,##0.00')
        _sett_kolonnebredder(ark_omega, KOLONNER_OMEGA_UTEN_APM, BREDDER_OMEGA_UTEN_APM, med_autofilter=False)
        tab_omega_ut = Table(
            displayName="OmegaUtenAPM",
            ref=f"A1:{get_column_letter(ark_omega.max_column)}{ark_omega.max_row}",
        )
        tab_omega_ut.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ark_omega.add_table(tab_omega_ut)

    # -- Ark: Mangler i Post sammenstilling --
    _ARK_MANGLER = "Mangler i Post samm."
    if _ARK_MANGLER in wb.sheetnames:
        ark_mangl = wb[_ARK_MANGLER]
        _formater_header(ark_mangl, _FILL_MANGL_HDR)
        kol_m = {celle.value: celle.column for celle in ark_mangl[1]}
        mengde_k_m  = kol_m.get(KOL_OMEGA)
        kontr_k_m   = kol_m.get(KOL_KONTRAKT)
        pris_k_m    = kol_m.get(KOL_ENHETSPRIS)
        merknad_k_m = kol_m.get("Merknad")
        for rad_nr in range(2, ark_mangl.max_row + 1):
            merknad_verdi = str(ark_mangl.cell(rad_nr, merknad_k_m).value or "") if merknad_k_m else ""
            fyll = _FILL_UTGAATT if "utgatt" in merknad_verdi.lower() else _FILL_MANGLER_OMEGA
            for celle in ark_mangl[rad_nr]:
                celle.fill = fyll
                celle.font = _FONT_DATA
                celle.alignment = _VTOPP
            _formater_tallcelle(ark_mangl, rad_nr, mengde_k_m, '#,##0.00')
            _formater_tallcelle(ark_mangl, rad_nr, kontr_k_m,  '#,##0.00')
            _formater_tallcelle(ark_mangl, rad_nr, pris_k_m,   '#,##0.00')
        _sett_kolonnebredder(ark_mangl, KOLONNER_MANGLER_SAMM, BREDDER_MANGLER_SAMM, med_autofilter=False)
        tab_mangl = Table(
            displayName="ManglerIPostSamm",
            ref=f"A1:{get_column_letter(ark_mangl.max_column)}{ark_mangl.max_row}",
        )
        tab_mangl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False,
        )
        ark_mangl.add_table(tab_mangl)
    _skriv_oversikt_ark(wb, df_hoved, df_post_samm, omega_path)

    # -- APM Rapp kontroll-ark --
    if apm_bf_rader:
        _GRNN = PatternFill("solid", fgColor="C6EFCE")
        _RAUD = PatternFill("solid", fgColor="FFC7CE")
        _HDR_BF  = PatternFill("solid", fgColor="1F4E79")
        _HDR_FNT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        _CAL10   = Font(name="Calibri", size=10)
        _TOP     = Alignment(vertical="top")

        ws_bf = wb.create_sheet("APM Rapp kontroll")
        for ci, (hdr, bredde, _) in enumerate(KOLONNER_APM_BF, 1):
            c = ws_bf.cell(row=1, column=ci, value=hdr)
            c.font = _HDR_FNT
            c.fill = _HDR_BF
            c.alignment = _TOP
            ws_bf.column_dimensions[get_column_letter(ci)].width = bredde
        ws_bf.freeze_panes = "A2"

        for ri, rad in enumerate(apm_bf_rader, 2):
            vals = [rad.get(k) for k in _APM_BF_NØKLAR]
            for ci, (val, (hdr, _, er_tal)) in enumerate(zip(vals, KOLONNER_APM_BF), 1):
                c = ws_bf.cell(row=ri, column=ci, value=val)
                c.font      = _CAL10
                c.alignment = _TOP
                if er_tal and isinstance(val, (int, float)):
                    c.number_format = "#,##0.00"
                if hdr == "Avvik (%)" and isinstance(val, (int, float)):
                    c.fill = _GRNN if abs(val) <= 3 else _RAUD

        if len(apm_bf_rader) > 0:
            tbl = Table(
                displayName="APMRappKontroll",
                ref=f"A1:{get_column_letter(len(KOLONNER_APM_BF))}{len(apm_bf_rader) + 1}",
            )
            tbl.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium7",
                showFirstColumn=False, showLastColumn=False,
                showRowStripes=True, showColumnStripes=False,
            )
            ws_bf.add_table(tbl)

    # -- Ark-rekkefølge --
    _orden = ["Forklaring og Sammendrag", "APM sammenstilling",
              "Post sammenstilling", "APM Rapp kontroll",
              "Omega uten APM", "Mangler i Post samm.",
              "Armering totalmengde", "Armeringsmengder"]
    wb._sheets = ([wb[n] for n in _orden if n in wb.sheetnames]
                  + [s for s in wb._sheets if s.title not in _orden])

    wb.save(sti)


# ===== GUI =====

def _start_gui():
    import threading
    import tkinter as tk
    from tkinter import filedialog, scrolledtext

    root = tk.Tk()
    root.title("Målebrev PDF -> Excel")
    root.resizable(True, True)
    root.minsize(560, 420)

    # Mappe-velger
    frm_mappe = tk.Frame(root, padx=10, pady=8)
    frm_mappe.pack(fill=tk.X)
    tk.Label(frm_mappe, text="Mappe:", width=7, anchor="w").pack(side=tk.LEFT)
    mappe_var = tk.StringVar(value="")
    tk.Entry(frm_mappe, textvariable=mappe_var, width=55).pack(
        side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

    def velg_mappe():
        valgt = filedialog.askdirectory(title="Velg mappe med PDF-filer")
        if valgt:
            mappe_var.set(valgt)

    tk.Button(frm_mappe, text="Bla...", command=velg_mappe, width=6).pack(side=tk.LEFT)

    # Omega-velger
    frm_omega = tk.Frame(root, padx=10)
    frm_omega.pack(fill=tk.X, pady=(0, 4))
    tk.Label(frm_omega, text="Omega:", width=7, anchor="w").pack(side=tk.LEFT)
    omega_var = tk.StringVar(value="")
    tk.Entry(frm_omega, textvariable=omega_var, width=55).pack(
        side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

    def velg_omega():
        valgt = filedialog.askopenfilename(
            title="Velg Omega Excel-fil",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Alle filer", "*.*")],
        )
        if valgt:
            omega_var.set(valgt)

    tk.Button(frm_omega, text="Bla...", command=velg_omega, width=6).pack(side=tk.LEFT)

    # Logg-vindu
    frm_logg = tk.Frame(root, padx=10)
    frm_logg.pack(fill=tk.BOTH, expand=True)
    logg_tekst = scrolledtext.ScrolledText(
        frm_logg, height=18, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD)
    logg_tekst.pack(fill=tk.BOTH, expand=True)

    # Knapper og status
    frm_knapper = tk.Frame(root, padx=10, pady=8)
    frm_knapper.pack(fill=tk.X)
    resultat_sti = [""]
    test_var = tk.BooleanVar(value=False)
    start_knapp = tk.Button(frm_knapper, text="Start", width=10)
    test_cb = tk.Checkbutton(frm_knapper, text="Test (30 %)", variable=test_var,
                             fg="#8B4513", font=("Segoe UI", 9))
    aapne_knapp = tk.Button(frm_knapper, text="Aapne Excel", width=12, state=tk.DISABLED)
    start_knapp.pack(side=tk.LEFT, padx=(0, 6))
    test_cb.pack(side=tk.LEFT, padx=(0, 10))
    aapne_knapp.pack(side=tk.LEFT)
    tk.Label(frm_knapper, text="", width=2).pack(side=tk.LEFT)
    status_var = tk.StringVar(value="Klar")
    tk.Label(frm_knapper, textvariable=status_var, anchor="w", fg="#555").pack(
        side=tk.LEFT, fill=tk.X, expand=True)

    def skriv_logg(melding: str):
        logg_tekst.config(state=tk.NORMAL)
        logg_tekst.insert(tk.END, melding + "\n")
        logg_tekst.see(tk.END)
        logg_tekst.config(state=tk.DISABLED)
        root.update_idletasks()

    def kjor_ekstraksjon():
        mappe = mappe_var.get().strip()
        if not os.path.isdir(mappe):
            skriv_logg(f"[ERROR] Mappen finnes ikke: {mappe}")
            return

        start_knapp.config(state=tk.DISABLED)
        aapne_knapp.config(state=tk.DISABLED)
        status_var.set("Kjorer...")
        logg_tekst.config(state=tk.NORMAL)
        logg_tekst.delete("1.0", tk.END)
        logg_tekst.config(state=tk.DISABLED)
        resultat_sti[0] = ""

        class _StdoutTilLogg:
            def write(self, tekst):
                if tekst.strip():
                    root.after(0, skriv_logg, tekst.rstrip())
            def flush(self):
                pass

        def bakgrunnsjobb():
            orig_stdout = sys.stdout
            sys.stdout = _StdoutTilLogg()
            try:
                omega = omega_var.get().strip()
                resultat = extract_folder(
                    mappe,
                    omega_path=omega if os.path.isfile(omega) else None,
                    test_modus=test_var.get(),
                )
                resultat_sti[0] = resultat
                if resultat:
                    root.after(0, lambda: status_var.set("[SUCCESS] Ferdig!"))
                    root.after(0, lambda: aapne_knapp.config(state=tk.NORMAL))
                else:
                    root.after(0, lambda: status_var.set("[ERROR] Ingen data."))
            except Exception as feil:
                root.after(0, skriv_logg, f"[ERROR] {feil}")
                root.after(0, lambda: status_var.set("Feil - se logg"))
            finally:
                sys.stdout = orig_stdout
                root.after(0, lambda: start_knapp.config(state=tk.NORMAL))

        threading.Thread(target=bakgrunnsjobb, daemon=True).start()

    def aapne_resultat():
        if resultat_sti[0] and os.path.isfile(resultat_sti[0]):
            os.startfile(resultat_sti[0])
            root.destroy()

    start_knapp.config(command=kjor_ekstraksjon)
    aapne_knapp.config(command=aapne_resultat)
    root.mainloop()


# ===== Oppstart =====

if __name__ == "__main__":
    if len(sys.argv) > 1:
        mappe = sys.argv[1]
        if not os.path.isdir(mappe):
            print(f"[ERROR] Mappen finnes ikke: {mappe}")
            sys.exit(1)
        print(f"Leser PDF-filer fra: {mappe}")
        extract_folder(mappe)
    else:
        _start_gui()