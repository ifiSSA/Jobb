# -*- coding: utf-8 -*-
"""
Prognoseark til FOB si manadlege pris- og mengderapport (K5A / K7).

Byggjer vidare på mengde_oversikt.py — denne modulen reknar IKKJE ut mengder
på nytt. Den les den ferdige `mengde_oversikt_*.xlsx` (ark I_modell,
Ikke_i_modell, BIM_utan_Omega), legg til ein prognosekolonne og skriv rapporten
i det formatet FOB har bestilt:

    Forside              Kontraktssum ved signert kontrakt, kontraktsendringar,
                         gjeldande kontraktsverdi, revisjonsbeskrivelsar,
                         forventa omfang av planlagde revisjonsbeskrivelsar,
                         total prognose.
    NS3420               Mengdeoversikt for postar i NS3420-arket.
    Prosesskoden         Mengdeoversikt for postar i prosesskodearket.
    Revisjonsbeskrivelser Nye postar/prosessar og utgåtte element.
    Alle poster          Heile grunnlaget.
    Kontroll             Avstemming og teljingar.

Prognoseregelen (avklart med FOB/Isak):
    Prognose = MAX(mengde frå modell, APM-mengde, kabelliste)  når minst éin har tal
    Prognose = kontraktsmengde                                 når ingen av dei har tal
    Prognose = kontraktsmengde  når prognosen kjem frå kabellista og verdien
               avvik meir enn KABEL_MAKS_AVVIK_KR frå kontraktsverdien
    Prognose = kontraktsmengde  når prognosen kjem frå APM, APM-mengda ikkje er
               stadfesta av modell/kontrakt, og ho avvik meir enn
               APM_MAKS_AVVIK_RAPPORTERT_PST frå det som er fakturert hittil
    Prognose = kontraktsmengde  når prognosen kjem frå APM og APM-mengda er
               meir enn APM_MAKS_UNDER_KONTRAKT_PST (50 %) under kontraktsmengda
    Prognose = fakturert (rapportert) mengd  når målebrevet er ferdig
               (SRES er sann i Omega)
    Prognose = den gamle prognosen i Omega  når prosesskoden startar med eit
               prefiks i OMEGA_BEHOLD_PREFIKS (01.). Manglar posten prognose
               i Omega, vert kontraktsmengda brukt.

Bruk:
    from prognose_ark import lag_prognose
    lag_prognose(oversikt_excel=..., omega_path=..., isy_path=[...])
"""

from __future__ import annotations

import os
import re
import shutil
import warnings
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# =============================================================================
# Konstantar
# =============================================================================

# Ark i mengde_oversikt_*.xlsx som inneheld postar
OVERSIKT_ARK = ("I_modell", "Ikke_i_modell")
OVERSIKT_BIM = "BIM_utan_Omega"

# Kolonneheadar i mengde_oversikt → internt namn
OVERSIKT_KOL = {
    "Mengdelistepost":          "post",
    "Fag":                      "fag",
    "Tittel (Omega)":           "tittel",
    "Mengde fra modell":        "modell",
    "Enhet i modell":           "modellenhet",
    "APM-mengde":               "apm",
    "Enhet":                    "enhet",
    "Enhetspris":               "enhetspris",
    "Kontraktsmengde":          "kontrakt",
    "Rapportert mengde":        "rapportert",
    "APM-kilde":                "apm_kilde",
    "Beskrivelse (ISY Linker)": "beskrivelse",
}

# Post-kolonnar i hovudarka — avgjer om posten høyrer til NS3420 eller Prosesskoden
NS3420_POSTKOL = ("Post nummer",)
PROSESS_POSTKOL = ("ProsessID: prosess, sted og element",)

# Står denne teksten i "All tekst" i ISY-kontrakten, er posten ein opsjon.
# Det vert lagt til i kolonnen "Sikkerhet grunn".
OPSJON_TEKST = "OPSJON"

GRUNNLAG_NS = "NS3420"
GRUNNLAG_PK = "Prosesskoden"
GRUNNLAG_NY = "Ny post (revisjonsbeskrivelse)"

# Omega: gruppeoverskrifter i rad 2 og kolonnane under kvar av dei
OMEGA_GRUPPER = {
    "kontraktsmengde": ["opprinnelig", "endringer", "gjeldende", "prognose"],
    "kontraktsverdier": ["opprinnelig", "endringer", "revidert", "gjeldende", "prognose"],
    "utført": ["faktisk", "rapportert"],
}

# Rund sum-einingar: modellen kan ikkje telje opp fleire runde summar
RUNDSUM = {"rs", "rund sum", "rundsum"}

# Omega-prognosen overstyrer den utrekna prognosen når han er sett bevisst:
#   - prognosen er 0 (posten går ut)
#   - posten har kommentar denne månaden
#   - prognosen avvik frå kontraktsmengda
# Slå av med omega_overstyrer=False når Omega-prognosen ikkje lenger skal
# gå føre den utrekna (dette gjeld berre for september 2026).
OMEGA_OVERSTYRER_STANDARD = True

# Ei overstyring vert ståande sjølv om grunnlaget flyttar seg — mengder endrar
# seg kvar månad, og å la overstyringa falle bort ville tvinge fram manuelt
# arbeid. I staden vert ho varsla når spriket mot det utrekna er stort nok til
# å vere verdt ein ny vurdering.
VARSEL_AVVIK_PST = 10.0        # prosent
VARSEL_AVVIK_KR = 50_000.0     # kroner

# Kabellista: spriker verdien av ein prognose frå kabellista meir enn dette
# frå kontraktsverdien (opp eller ned), vert kontraktsmengda brukt i staden.
# None = slå av regelen.
KABEL_MAKS_AVVIK_KR: float | None = 500_000.0

# APM-mengda er skjør — ho kjem frå målebrev som kan vere ufullstendige eller
# dobbeltførte. To kontrollar dempar det:
#   1. Er APM-mengda lik modellmengda eller kontraktsmengda (innanfor
#      APM_BEKREFTA_PST), er ho stadfesta av ei uavhengig kjelde. Då er
#      prognosen sikrare, og ho står.
#   2. Elles: avvik APM-mengda meir enn APM_MAKS_AVVIK_RAPPORTERT_PST frå det
#      som er fakturert hittil, vert kontraktsmengda brukt i staden.
# Kontroll 2 krev at noko er fakturert. Er rapportert mengd tom eller null,
# finst det ikkje noko å samanlikne med, og APM-mengda står.
# None = slå av kontroll 2.
APM_BEKREFTA_PST = 1.0                                # prosent
APM_MAKS_AVVIK_RAPPORTERT_PST: float | None = 20.0    # prosent

# Er APM-mengda meir enn dette under kontraktsmengda, er målebreva truleg
# ufullstendige, og kontraktsmengda vert brukt i staden. Gjeld òg når
# APM-mengda er stadfesta av modellen. None = slå av regelen.
APM_MAKS_UNDER_KONTRAKT_PST: float | None = 50.0      # prosent

# Sikkerheit på prognosen (kolonnen "Sikkerhet"):
#   Helt sikker  prognosen er sett til 0 i Omega, eller målebrevet er ferdig
#   Sikker       minst to kjelder (modell, APM, kabelliste, kontrakt) er
#                "så og si like" prognosen — innanfor SIKKERHET_LIK_PST.
#                Ein Omega-prognose med kommentar, eller ein som ei kjelde
#                stadfestar, tel òg som sikker.
#   Usikker      APM-/kabelregelen har sett inn kontraktsmengda, Omega-
#                prognosen er merkt "Vurder på nytt", modellen har enhetsavvik,
#                eller kjeldene spriker meir enn SIKKERHET_SPRIK_PST
#   Middels      resten — t.d. berre éi kjelde, eller berre kontraktsmengda
SIKKERHET_LIK_PST = 5.0       # prosent
SIKKERHET_SPRIK_PST = 50.0    # prosent

# Omega-opplasting (BOQNT-arket som skal lastast opp att i Omega).
# Rad 2 i malen har dei tekniske kolonnenamna. Vi finn kolonnane på namn og
# ikkje på posisjon, så arket toler at kolonnar vert flytta eller lagt til.
#   Name        = prosesskoden — koplinga mot mengdelisteposten vår
#   ForecastQty = prognosekolonnen som skal oppdaterast
OMEGA_OPPL_KOL_POST = "name"
OMEGA_OPPL_KOL_PROGNOSE = "forecastqty"
OMEGA_OPPL_DESIMALAR = 4      # rundar bort flyttalsstøy frå utrekninga

# Ferdig målebrev: SRES-kolonnen i Omega er sann. Då er mengda gjort opp, og
# prognosen skal vere lik den fakturerte (rapporterte) mengda.
# Sett False for å slå av regelen.
BRUK_FERDIG_MAALEBREV = True

# Postar som alltid skal behalde prognosen som alt står i Omega, uansett kva
# vi reknar ut. Prefikset vert samanlikna mot starten av prosesskoden.
# Tom tuple = slå av regelen.
OMEGA_BEHOLD_PREFIKS = ("01.",)

# Ein kommentar i Omega overstyrer ikkje i seg sjølv. Berre kommentarar som
# seier at posten er utgått, eller at mengda er regulert som følgje av ei
# endringsordre (EO) eller ein PF, tel. Lista kan utvidast her.
KOMMENTAR_UTGATT = ("utgått", "utgatt", "utgår", "utgaar", "utgaatt", "utgåtte")

# Kommentarar som seier at prognosen er lagt inn manuelt i Omega. Då skal
# Omega-prognosen brukast framfor den utrekna.
# >>> FYLL UT med ordlyden de faktisk brukar i Omega. Lista er med vilje smal:
#     kvart ord her gjer at posten hentar prognosen frå Omega i staden. <<<
KOMMENTAR_EGEN_PROGNOSE = (
    "egen prognose", "eigen prognose", "manuell prognose",
    "manuelt satt", "manuelt sett", "prognose satt", "prognose sett",
)
KOMMENTAR_MENGDEREGULERING = ("mengderegul", "mengde regul", "mengde-regul")
_RE_EO_PF = re.compile(r"\b(eo|pf)[-\s]?\d*\b", re.IGNORECASE)


def er_sann(verdi) -> bool:
    """Tolk ein Sann/Usann-verdi frå Omega (bool eller tekst).

    Same tolking som i pdf_maalebrev_til_excel: true/sann/1/yes/ja er sant,
    alt anna — medrekna tomt og ukjent — er usant.
    """
    if verdi is None or verdi == "":
        return False
    if isinstance(verdi, bool):
        return verdi
    return str(verdi).strip().lower() in ("true", "sann", "1", "yes", "ja")


def kommentar_grunn(tekst) -> str:
    """
    Kva i kommentaren som utløyser overstyring — tom streng når ingenting gjer det.

        "Posten utgår, se EO-078"            → "utgått"
        "Mengderegulert iht. EO-122"         → "mengderegulering (EO)"
        "Avventer avklaring med byggherre"   → ""
    """
    rå = str(tekst or "")
    t = _norm(rå)
    if not t:
        return ""
    if any(ord_ in t for ord_ in KOMMENTAR_UTGATT):
        return "utgått"
    if any(ord_ in t for ord_ in KOMMENTAR_MENGDEREGULERING):
        m = _RE_EO_PF.search(rå)
        if m:
            return f"mengderegulering ({m.group(1).upper()})"
    if any(ord_ in t for ord_ in KOMMENTAR_EGEN_PROGNOSE):
        return "egen prognose"
    return ""


# =============================================================================
# Stil
# =============================================================================

_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
_TIT_FILL = PatternFill("solid", fgColor="DCE6F1")
_GUL_FILL = PatternFill("solid", fgColor="FFF2CC")
_CAL10 = Font(name="Calibri", size=10)
_SIKKERHET_FILL = {
    "Helt sikker": PatternFill("solid", fgColor="A9D08E"),
    "Sikker":      PatternFill("solid", fgColor="E2EFDA"),
    "Middels":     PatternFill("solid", fgColor="FFF2CC"),
    "Usikker":     PatternFill("solid", fgColor="F8CBAD"),
}
_CAL10B = Font(name="Calibri", size=10, bold=True)
_TOP = Alignment(vertical="top")
_WRAP = Alignment(wrap_text=True, vertical="top")
_HOYRE = Alignment(horizontal="right", vertical="top")
_TYNN = Side(style="thin", color="BFBFBF")

_KR = "#,##0"
_TAL = "#,##0.00"
_PST = "#,##0.0"


# =============================================================================
# Hjelparar
# =============================================================================

def _tal(v):
    """Robust tal-parsing: norsk desimalkomma, einingssuffiks, stjernemerking."""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if pd.notna(v) else None
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "-", "--"):
        return None
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s in ("", "-", ".", ","):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


# Einingar som tyder det same. Modellen og kontrakta skriv dei ulikt.
_ENHET_SYNONYM = {
    "m2": "m2", "m²": "m2", "kvm": "m2",
    "m3": "m3", "m³": "m3", "kbm": "m3",
    "stk": "stk", "stk.": "stk", "st": "stk", "styk": "stk",
    "lm": "m", "m": "m", "meter": "m",
    "kg": "kg", "tonn": "tonn", "t": "tonn",
    "rs": "rs", "rundsum": "rs", "rund sum": "rs",
    "time": "time", "timer": "time", "t.": "time",
    "uke": "uke", "stk/mnd": "stk",
}

# Einingar som kan reknast om utan tap. Alt anna vert flagga, ikkje konvertert.
_OMREGNING = {("kg", "tonn"): 0.001, ("tonn", "kg"): 1000.0}


def normaliser_enhet(enhet) -> str:
    """'Stk.' → 'stk', 'm²' → 'm2', 'RS' → 'rs'. Tom streng når ukjend."""
    e = _norm(enhet).replace(".", "").strip()
    if not e or e in ("nan", "none", "-", "--", "ingen"):
        return ""
    return _ENHET_SYNONYM.get(e, e)


def enhet_status(modellenhet, kontraktsenhet) -> tuple[str, float]:
    """
    Samanlikn eininga modellen målte i mot eininga posten gjerast opp i.

    Returnerer (status, faktor):
        "ok"        modellmengda kan brukast som ho er (faktor 1.0)
        "omregnet"  same storleik, ulik eining — modellmengda skalerast
        "ukjent"    ei av einingane manglar — modellmengda vert brukt,
                    men markert
        "avvik"     ulike storleikar (t.d. modell m2 mot post stk) —
                    modellmengda er ikkje samanliknbar og vert halden
                    utanfor prognosen
    """
    m, k = normaliser_enhet(modellenhet), normaliser_enhet(kontraktsenhet)
    if not m or not k:
        return "ukjent", 1.0
    if m == k:
        return "ok", 1.0
    faktor = _OMREGNING.get((m, k))
    if faktor:
        return "omregnet", faktor
    return "avvik", 1.0


def _er_rundsum(enhet) -> bool:
    return _norm(enhet).replace(".", "") in RUNDSUM


def _sorteringsnokkel(post: str):
    """Naturleg sortering av postnummer: 1303.3.5 før 1303.30.5."""
    return tuple(
        (0, int(b)) if b.isdigit() else (1, b)
        for b in re.split(r"[.\-]", str(post))
    )


# =============================================================================
# 1. Grunnlag: kva hovudark høyrer posten til?
# =============================================================================

def les_grunnlag(isy_path: str | list[str]) -> dict[str, str]:
    """
    Returner {postnummer: "NS3420" | "Prosesskoden"}.

    Same logikk som mengde_oversikt.les_isy brukar for å finne post-kolonna,
    men her held vi kjeldene frå kvarandre slik at rapporten kan delast i to
    arkfaner slik FOB har bestilt.
    """
    stiar = [isy_path] if isinstance(isy_path, str) else list(isy_path)
    grunnlag: dict[str, str] = {}

    for sti in stiar:
        if not sti or not os.path.exists(sti):
            continue
        df = pd.read_excel(sti, engine="openpyxl")
        cols = list(df.columns)

        kilde = None
        postkol = None
        for k in NS3420_POSTKOL:
            if k in cols:
                kilde, postkol = GRUNNLAG_NS, k
                break
        if postkol is None:
            for k in PROSESS_POSTKOL:
                if k in cols:
                    kilde, postkol = GRUNNLAG_PK, k
                    break
        if postkol is None:
            postkol = cols[0]
            kilde = GRUNNLAG_PK if "prosess" in _norm(sti) else GRUNNLAG_NS

        n = 0
        for v in df[postkol].dropna():
            p = str(v).strip()
            if p and p.lower() != "nan":
                grunnlag.setdefault(p, kilde)
                n += 1
        print(f"   Grunnlag {kilde:<13} {n:>6} postar  ({os.path.basename(sti)})")

    return grunnlag


def les_opsjonsposter(isy_path: str | list[str]) -> set[str]:
    """Postnummer der "All tekst" i ISY-kontrakten inneheld OPSJON."""
    stiar = [isy_path] if isinstance(isy_path, str) else list(isy_path)
    opsjon: set[str] = set()

    for sti in stiar:
        if not sti or not os.path.exists(sti):
            continue
        df = pd.read_excel(sti, engine="openpyxl")
        postkol = next((k for k in NS3420_POSTKOL + PROSESS_POSTKOL
                        if k in df.columns), None)
        tekstkol = next((k for k in ("All tekst", "Tekst - alt. (uten linjeskift)")
                         if k in df.columns), None)
        if postkol is None or tekstkol is None:
            continue

        treff = df[tekstkol].astype(str).str.contains(OPSJON_TEKST, regex=False)
        for v in df.loc[treff, postkol].dropna():
            p = str(v).strip()
            if p and p.lower() != "nan":
                opsjon.add(p)

    if opsjon:
        print(f"   Opsjon: {len(opsjon):,} postar med «{OPSJON_TEKST}» i teksten")
    return opsjon


# =============================================================================
# 2. Omega: kontraktsverdiar via gruppeoverskriftene
# =============================================================================

def les_kontraktsverdier(omega_path: str) -> pd.DataFrame:
    """
    Les kontraktsmengder og kontraktsverdiar frå Omega.

    Omega-uttrekket har to headerrader: ei med gruppeoverskrifter
    ('Kontraktsmengde', 'Kontraktsverdier', 'Utført') og ei med kolonnenamn.
    Kolonnenamna gjentek seg mellom gruppene ('Opprinnelig', 'Gjeldende',
    'Prognose'), og Omega dedupliserer dei med talsuffiks som varierer mellom
    eksportar. Difor les vi posisjonen til kvar gruppe i staden for å stole
    på namna.

    Returnerer ein DataFrame med kolonnane:
        post, tittel, enhet, enhetspris,
        km_opprinnelig, km_endringer, km_gjeldende, km_prognose,
        kv_opprinnelig, kv_endringer, kv_revidert, kv_gjeldende, kv_prognose,
        utf_faktisk, utf_rapportert
    """
    from base_lesning import finn_omega_ark

    wb = openpyxl.load_workbook(omega_path, read_only=True, data_only=True)
    ws = wb[finn_omega_ark(omega_path)]
    rader = list(ws.iter_rows(values_only=True))
    wb.close()

    # Finn headerraden: den som inneheld 'Navn (Prosesskode)'
    hdr_idx = None
    for i, r in enumerate(rader[:10]):
        if any(_norm(c) == "navn (prosesskode)" for c in r):
            hdr_idx = i
            break
    if hdr_idx is None:
        raise ValueError(f"Fann ikkje headerrad i {os.path.basename(omega_path)}")

    header = [_norm(c) for c in rader[hdr_idx]]
    gruppe_rad = [_norm(c) for c in rader[hdr_idx - 1]] if hdr_idx > 0 else []

    # Kolonneindeks for kvar gruppe: frå gruppeoverskrifta til neste
    gruppe_start = {g: i for i, g in enumerate(gruppe_rad) if g in OMEGA_GRUPPER}
    grenser = sorted(gruppe_start.values()) + [len(header)]

    kart: dict[str, int] = {}
    for gruppe, start in gruppe_start.items():
        slutt = next(g for g in grenser if g > start)
        felt = OMEGA_GRUPPER[gruppe]
        prefiks = {"kontraktsmengde": "km", "kontraktsverdier": "kv", "utført": "utf"}[gruppe]
        for kol in range(start, slutt):
            navn = re.sub(r"\d+$", "", header[kol]).strip()
            if navn in felt:
                kart[f"{prefiks}_{navn}"] = kol

    for enkel, treff in (("post", "navn (prosesskode)"), ("tittel", "tittel"),
                         ("enhet", "enhet"), ("enhetspris", "enhetspris"),
                         ("kommentar", "kommentarer"),
                         ("kommentar_forrige", "kommentar fra forrige delmålebrev"),
                         ("sres", "sres")):
        for i, h in enumerate(header):
            if h == treff:
                kart[enkel] = i
                break

    ugyldig = {"", "nan", "none", "total", "sum"}
    ut = []
    for r in rader[hdr_idx + 1:]:
        if kart.get("post") is None or kart["post"] >= len(r):
            continue
        post = str(r[kart["post"]] or "").strip()
        if post.lower() in ugyldig:
            continue        # totalsumrada nedst — ville dobla alle verdiar
        rad = {"post": post}
        for navn, kol in kart.items():
            if navn == "post":
                continue
            v = r[kol] if kol < len(r) else None
            rad[navn] = (v if navn in ("tittel", "enhet", "kommentar",
                                       "kommentar_forrige", "sres")
                         else _tal(v))
        ut.append(rad)

    df = pd.DataFrame(ut).drop_duplicates(subset=["post"], keep="first")
    for k in ("kommentar", "kommentar_forrige"):
        if k not in df.columns:
            df[k] = ""
        df[k] = (df[k].fillna("").astype(str)
                 .str.replace("_x000D_", " ", regex=False).str.strip())
    for k in ("km_opprinnelig", "km_endringer", "km_gjeldende", "km_prognose",
              "kv_opprinnelig", "kv_endringer", "kv_revidert", "kv_gjeldende",
              "kv_prognose", "utf_faktisk", "utf_rapportert", "enhetspris"):
        if k not in df.columns:
            df[k] = None
        df[k] = pd.to_numeric(df[k], errors="coerce")

    # Manglar kontraktsverdiar? Rekn dei ut frå mengde x enhetspris.
    if df["kv_opprinnelig"].fillna(0).sum() == 0:
        df["kv_opprinnelig"] = df["km_opprinnelig"] * df["enhetspris"]
    if df["kv_gjeldende"].fillna(0).sum() == 0:
        df["kv_gjeldende"] = df["km_gjeldende"] * df["enhetspris"]
    if df["kv_endringer"].fillna(0).sum() == 0:
        df["kv_endringer"] = df["kv_gjeldende"].fillna(0) - df["kv_opprinnelig"].fillna(0)

    print(f"   Omega kontraktsverdiar: {len(df):,} postar, "
          f"gjeldande {df['kv_gjeldende'].sum():,.0f} kr")
    return df


# =============================================================================
# 3. Les mengde_oversikt
# =============================================================================

def les_oversikt(oversikt_excel: str) -> pd.DataFrame:
    """Les postradene frå mengde_oversikt_*.xlsx (I_modell + Ikke_i_modell)."""
    deler = []
    for ark in OVERSIKT_ARK:
        try:
            df = pd.read_excel(oversikt_excel, sheet_name=ark, engine="openpyxl")
        except ValueError:
            continue
        df = df.rename(columns={k: v for k, v in OVERSIKT_KOL.items() if k in df.columns})
        df["oversiktsark"] = ark
        deler.append(df)
    if not deler:
        raise ValueError(f"Fann ingen postark i {os.path.basename(oversikt_excel)}")

    df = pd.concat(deler, ignore_index=True)
    for k in OVERSIKT_KOL.values():
        if k not in df.columns:
            df[k] = None
    df["post"] = df["post"].astype(str).str.strip()
    df = df[~df["post"].str.lower().isin(["", "nan", "none"])]
    for k in ("modell", "apm", "enhetspris", "kontrakt", "rapportert"):
        df[k] = pd.to_numeric(df[k], errors="coerce")
    df = df.drop_duplicates(subset=["post"], keep="first")
    print(f"   Mengde-oversikt: {len(df):,} postar "
          f"({os.path.basename(oversikt_excel)})")
    return df


def les_bim_utan_omega(oversikt_excel: str) -> pd.DataFrame:
    """Les BIM_utan_Omega — modellpostar utan kontraktsmatch."""
    try:
        df = pd.read_excel(oversikt_excel, sheet_name=OVERSIKT_BIM, engine="openpyxl")
    except ValueError:
        return pd.DataFrame(columns=["post", "fag", "modell", "enhet", "kilde"])
    df = df.rename(columns={
        "Mengdelistepost": "post", "Fag": "fag",
        "Mengde fra modell": "modell", "Enhet": "enhet", "Kilde": "kilde",
    })
    df["post"] = df["post"].astype(str).str.strip()
    df["modell"] = pd.to_numeric(df["modell"], errors="coerce")
    return df[df["post"].str.lower().isin(["", "nan"]) == False]


# =============================================================================
# 4. Prognoseberegning
# =============================================================================

SIKKERHET_NIVAA = ("Helt sikker", "Sikker", "Middels", "Usikker")


def _nesten_lik(a, b) -> bool:
    """True når a og b er innanfor SIKKERHET_LIK_PST av den største."""
    return abs(a - b) <= max(abs(a), abs(b)) * (SIKKERHET_LIK_PST / 100.0)


def prognose_sikkerhet(r) -> tuple[str, str]:
    """Kor sikker prognosen på rada er, og kvifor. Sjå SIKKERHET_* over."""
    if r["omega_null"]:
        return "Helt sikker", "Satt til 0 i Omega"
    if r["maalebrev_ferdig"]:
        return "Helt sikker", "Målebrev avsluttet"

    if r["apm_til_kontrakt"]:
        return "Usikker", "APM spriker mot fakturert"
    if r["apm_under_kontrakt"]:
        return "Usikker", "APM langt under kontrakt"
    if r["kabel_til_kontrakt"]:
        return "Usikker", "Kabelliste spriker mot kontrakt"

    prognose = r["prognose"]
    if pd.isna(prognose):
        return "Usikker", "Mangler grunnlag"

    kjelder = {
        "Modell": r["modell_i_prognose"],
        "APM": r["apm"],
        "Kabelliste": r["kabel"],
        "Kontrakt": r["kontrakt"],
    }
    kjelder = {k: v for k, v in kjelder.items() if pd.notna(v)}
    like = [k for k, v in kjelder.items() if _nesten_lik(v, prognose)]

    # Ein Omega-prognose er alt eit bevisst val — éi stadfestande kjelde held.
    # Ein utrekna prognose kjem frå éi av kjeldene, så ho treng éi til.
    if r["overstyrt_av_omega"]:
        if r["omega_kommentert"]:
            return "Sikker", "Manuelt satt i Omega (kommentar)"
        if like:
            return "Sikker", "Omega-prognose ≈ " + " og ".join(like)
    elif len(like) >= 2:
        return "Sikker", " ≈ ".join(like)

    if r["varsel"]:
        return "Usikker", "Omega-prognose spriker mot beregnet"
    if r["enhetsstatus"] == "avvik":
        return "Usikker", "Enhetsavvik i modell"

    verdiar = list(kjelder.values())
    storst = max((abs(v) for v in verdiar), default=0)
    if len(verdiar) >= 2 and storst > 0:
        sprik = (max(verdiar) - min(verdiar)) / storst * 100
        if sprik > SIKKERHET_SPRIK_PST:
            lav = min(kjelder, key=kjelder.get)
            hoy = max(kjelder, key=kjelder.get)
            return "Usikker", f"{hoy} og {lav} spriker ({sprik:.0f} %)"

    if list(kjelder) == ["Kontrakt"]:
        return "Middels", "Kun kontraktsmengde"
    if len(kjelder) == 1:
        return "Middels", f"Kun {next(iter(kjelder))}"
    if r["overstyrt_av_omega"]:
        return "Middels", "Omega-prognose uten støtte i kildene"
    return "Middels", "Kildene er nær, men ikke like"


def bygg_prognose(
    oversikt: pd.DataFrame,
    omega: pd.DataFrame,
    grunnlag: dict[str, str],
    kabel: dict[str, dict] | None = None,
    omega_overstyrer: bool = OMEGA_OVERSTYRER_STANDARD,
    bruk_forrige_kommentar: bool = False,
    opsjon: set[str] | None = None,
) -> pd.DataFrame:
    """Slå saman kjeldene og rekn ut prognosen per post."""
    df = oversikt.merge(omega, on="post", how="outer", suffixes=("", "_om"))

    # Kabelmengder frå kabellistene — kablane finst ikkje som objekt i modellen
    kabel = kabel or {}
    df["kabel"] = df["post"].map(lambda p: kabel.get(p, {}).get("mengde"))
    df["kabeltyper"] = df["post"].map(lambda p: kabel.get(p, {}).get("kabeltyper", ""))
    df["kabel"] = pd.to_numeric(df["kabel"], errors="coerce")

    df["tittel"] = df["tittel"].fillna(df.get("tittel_om", "")).fillna("")
    df["enhet"] = df["enhet"].fillna(df.get("enhet_om", "")).fillna("")
    df["enhetspris"] = df["enhetspris"].fillna(df.get("enhetspris_om"))
    df["kontrakt"] = df["kontrakt"].fillna(df["km_gjeldende"])
    df["rapportert"] = df["rapportert"].fillna(df["utf_rapportert"])
    for k in ("fag", "apm_kilde", "beskrivelse", "oversiktsark",
              "kommentar", "kommentar_forrige"):
        if k not in df.columns:
            df[k] = ""
        df[k] = df[k].fillna("")

    df["grunnlag"] = df["post"].map(grunnlag).fillna(GRUNNLAG_NY)

    # Rund sum: modellen tel ikkje opp fleire runde summar, så modellmengda
    # held seg utanfor prognosen for desse postane.
    df["rundsum"] = df["enhet"].map(_er_rundsum)

    # Enhetskontroll: modellen kan ha målt noko anna enn posten gjerast opp i
    # (t.d. areal på eit sandwichpanel som skal avreknast per stk). Då er
    # modellmengda ikkje samanliknbar, og ho vert halden utanfor prognosen.
    if "modellenhet" not in df.columns:
        df["modellenhet"] = ""
    df["modellenhet"] = df["modellenhet"].fillna("")
    status = df.apply(lambda r: enhet_status(r["modellenhet"], r["enhet"]), axis=1)
    df["enhetsstatus"] = [t[0] for t in status]
    df["enhetsfaktor"] = [t[1] for t in status]
    df.loc[df["modell"].isna(), "enhetsstatus"] = ""

    # Rekn om der einingane er same storleik (kg ↔ tonn)
    df["modell_omregnet"] = df["modell"] * df["enhetsfaktor"]

    modell_i_prognose = df["modell_omregnet"].where(
        ~df["rundsum"] & (df["enhetsstatus"] != "avvik"))
    df["modell_i_prognose"] = modell_i_prognose

    par = pd.concat([modell_i_prognose, df["apm"], df["kabel"]], axis=1)
    har_tall = modell_i_prognose.notna() | df["apm"].notna() | df["kabel"].notna()
    df["prognose"] = par.max(axis=1, skipna=True)
    df.loc[~har_tall, "prognose"] = df.loc[~har_tall, "kontrakt"]

    def kilde(r):
        modell_ute = r["rundsum"] or r["enhetsstatus"] == "avvik"
        kandidatar = {
            "Modell": None if modell_ute else r["modell_omregnet"],
            "APM": r["apm"],
            "Kabelliste": r["kabel"],
        }
        med_tall = {k: v for k, v in kandidatar.items() if pd.notna(v)}
        if not med_tall:
            if r["enhetsstatus"] == "avvik" and pd.notna(r["kontrakt"]):
                return "Kontraktsmengde (enhetsavvik i modell)"
            if r["rundsum"] and pd.notna(r["kontrakt"]):
                return "Kontraktsmengde (rund sum)"
            return "Kontraktsmengde" if pd.notna(r["kontrakt"]) else "Mangler grunnlag"
        toppverdi = max(med_tall.values())
        vinnarar = [k for k, v in med_tall.items() if abs(v - toppverdi) < 1e-9]
        return " = ".join(vinnarar)

    df["prognosekilde"] = df.apply(kilde, axis=1)

    # --- APM-mengda mot andre kjelder -------------------------------------
    # Kontroll 1: er APM-mengda lik modellmengda eller kontraktsmengda, er ho
    # stadfesta uavhengig, og prognosen er sikrare. Ho står uansett kontroll 2.
    # Kontroll 2: elles, avvik ho meir enn APM_MAKS_AVVIK_RAPPORTERT_PST frå
    # det som er fakturert hittil, vert kontraktsmengda brukt i staden.
    # Ligg før kabelregelen og Omega-overstyringa.
    df["apm_bekrefta"] = False
    df["apm_til_kontrakt"] = False
    fra_apm = df["prognosekilde"].str.contains("APM", regex=False)

    def _er_lik(a, b):
        """True der a og b er like innanfor APM_BEKREFTA_PST prosent av b."""
        return ((a - b).abs() <= b.abs() * (APM_BEKREFTA_PST / 100.0)
                ).fillna(False)

    # Modellmengda tel berre som stadfesting når ho er med i prognosen —
    # ei rundsum- eller enhetsavvik-mengd er ikkje samanliknbar.
    bekrefta = (fra_apm & (_er_lik(df["apm"], modell_i_prognose)
                           | _er_lik(df["apm"], df["kontrakt"]))
                ).fillna(False).astype(bool)
    df["apm_bekrefta"] = bekrefta

    if APM_MAKS_AVVIK_RAPPORTERT_PST is not None:
        fakturert = df["rapportert"].notna() & (df["rapportert"] != 0)
        avvik_rapp_pst = ((df["apm"] - df["rapportert"]).abs()
                          / df["rapportert"].abs() * 100)
        apm_til_kontrakt = (fra_apm & ~bekrefta & fakturert
                            & df["kontrakt"].notna()
                            & (avvik_rapp_pst > APM_MAKS_AVVIK_RAPPORTERT_PST)
                            ).fillna(False).astype(bool)

        df["apm_til_kontrakt"] = apm_til_kontrakt
        df.loc[apm_til_kontrakt, "prognose"] = df.loc[apm_til_kontrakt, "kontrakt"]
        df.loc[apm_til_kontrakt, "prognosekilde"] = (
            f"Kontraktsmengde (APM avvik > {APM_MAKS_AVVIK_RAPPORTERT_PST:g} %"
            f" mot fakturert)")
        if bekrefta.any() or apm_til_kontrakt.any():
            print(f"   APM-kontroll: {int(bekrefta.sum()):,} poster bekreftet av "
                  f"modell/kontrakt, {int(apm_til_kontrakt.sum()):,} satt til "
                  f"kontraktsmengde "
                  f"(avvik > {APM_MAKS_AVVIK_RAPPORTERT_PST:g} % mot fakturert)")

    # Kontroll 3: ligg APM-mengda meir enn APM_MAKS_UNDER_KONTRAKT_PST under
    # kontraktsmengda, vert kontraktsmengda brukt i staden.
    df["apm_under_kontrakt"] = False
    if APM_MAKS_UNDER_KONTRAKT_PST is not None:
        fra_apm = df["prognosekilde"].str.contains("APM", regex=False)
        grense = df["kontrakt"] * (1 - APM_MAKS_UNDER_KONTRAKT_PST / 100.0)
        under = (fra_apm & (df["kontrakt"] > 0) & (df["apm"] < grense)
                 ).fillna(False).astype(bool)

        df["apm_under_kontrakt"] = under
        df.loc[under, "prognose"] = df.loc[under, "kontrakt"]
        df.loc[under, "prognosekilde"] = (
            f"Kontraktsmengde (APM > {APM_MAKS_UNDER_KONTRAKT_PST:g} %"
            f" under kontrakt)")
        if under.any():
            print(f"   APM-kontroll: {int(under.sum()):,} poster satt til "
                  f"kontraktsmengde (APM > {APM_MAKS_UNDER_KONTRAKT_PST:g} % "
                  f"under kontrakt)")

    # --- Kabelprognose med stort sprik mot kontrakt -----------------------
    # Når prognosen kjem frå kabellista og verdien hennar avvik meir enn
    # KABEL_MAKS_AVVIK_KR frå kontraktsverdien (opp eller ned), vert
    # kontraktsmengda brukt i staden. Kabelmengda står att i "kabel" til
    # kontroll. Ligg før Omega-overstyringa, så ein avtalt Omega-prognose
    # framleis vinn.
    df["kabel_til_kontrakt"] = False
    if KABEL_MAKS_AVVIK_KR is not None:
        kv = df["kv_gjeldende"].where(df["kv_gjeldende"].notna(),
                                      df["kontrakt"] * df["enhetspris"])
        kabel_avvik_kr = (df["prognose"] * df["enhetspris"] - kv).abs()
        fra_kabel = df["prognosekilde"].str.contains("Kabelliste", regex=False)
        til_kontrakt = (fra_kabel & df["kontrakt"].notna()
                        & (kabel_avvik_kr > KABEL_MAKS_AVVIK_KR)
                        ).fillna(False).astype(bool)

        df["kabel_til_kontrakt"] = til_kontrakt
        df.loc[til_kontrakt, "prognose"] = df.loc[til_kontrakt, "kontrakt"]
        df.loc[til_kontrakt, "prognosekilde"] = (
            f"Kontraktsmengde (kabelliste avvik > {KABEL_MAKS_AVVIK_KR:,.0f} kr)")
        if til_kontrakt.any():
            print(f"   Kabelprognose erstattet med kontraktsmengde for "
                  f"{int(til_kontrakt.sum()):,} poster "
                  f"(avvik > {KABEL_MAKS_AVVIK_KR:,.0f} kr)")

    # --- Ferdig målebrev --------------------------------------------------
    # Er SRES sann i Omega, er målebrevet gjort opp. Då er den fakturerte
    # (rapporterte) mengda endeleg, og prognosen skal vere lik ho.
    # Ligg før Omega-overstyringa: ein prognose som er sett bevisst i Omega
    # (0 eller kommentert) går framfor.
    df["maalebrev_ferdig"] = False
    if BRUK_FERDIG_MAALEBREV:
        if "sres" not in df.columns:
            df["sres"] = None
        ferdig = df["sres"].map(er_sann).astype(bool)
        har_mengd = ferdig & df["rapportert"].notna()

        df["maalebrev_ferdig"] = har_mengd
        df.loc[har_mengd, "prognose"] = df.loc[har_mengd, "rapportert"]
        df.loc[har_mengd, "prognosekilde"] = "Fakturert mengde (målebrev ferdig)"
        df.loc[har_mengd, "kabel_til_kontrakt"] = False
        df.loc[har_mengd, "apm_til_kontrakt"] = False
        df.loc[har_mengd, "apm_under_kontrakt"] = False

        utan_mengd = int((ferdig & df["rapportert"].isna()).sum())
        if ferdig.any():
            print(f"   Ferdige målebrev: {int(har_mengd.sum()):,} poster satt "
                  f"til fakturert mengde")
            if utan_mengd:
                print(f"      [ADVARSEL] {utan_mengd:,} ferdige poster manglar "
                      f"rapportert mengde — prognosen står urørt")

    # --- Omega-prognosen overstyrer ---------------------------------------
    # Der prognosen i Omega er sett bevisst — nullstilt, kommentert, eller
    # ulik kontraktsmengda — er han avtalt og skal gjelde framfor den
    # utrekna. Gjeld september 2026; slå av med omega_overstyrer=False.
    df["omega_prognose"] = df["km_prognose"]
    df["kommentargrunn"] = df["kommentar"].map(kommentar_grunn)
    if bruk_forrige_kommentar:
        forrige = df["kommentar_forrige"].map(kommentar_grunn)
        df["kommentargrunn"] = df["kommentargrunn"].where(
            df["kommentargrunn"] != "", forrige)
    df["har_kommentar"] = df["kommentargrunn"] != ""

    satt = df["omega_prognose"].notna()

    # Faste reglar — gjeld uavhengig av omega_overstyrer:
    #   - prognosen er sett til 0: posten går ut, og 0 skal stå
    #   - posten har ein kvalifiserande kommentar: prognosen er lagt inn
    #     manuelt i Omega, og den skal brukast
    #   - posten startar med eit prefiks i OMEGA_BEHOLD_PREFIKS (01.): den
    #     gamle prognosen i Omega skal stå urørt
    er_null = satt & (df["omega_prognose"] == 0)
    kommentert = satt & df["har_kommentar"]

    if OMEGA_BEHOLD_PREFIKS:
        har_prefiks = df["post"].astype(str).str.strip().str.startswith(
            OMEGA_BEHOLD_PREFIKS)
    else:
        har_prefiks = pd.Series(False, index=df.index)
    er_behold = satt & har_prefiks

    grunn = pd.Series("", index=df.index, dtype=object)
    grunn[kommentert] = ("Omega prognose ("
                         + df.loc[kommentert, "kommentargrunn"] + ")")

    if omega_overstyrer:
        # Brei, mellombels regel: all prognose som skil seg frå kontraktsmengda
        # vert rekna som avtalt. Slå av med omega_overstyrer=False — då står
        # dei faste reglane over att.
        avvik_kontrakt = (satt & df["km_gjeldende"].notna()
                          & (df["omega_prognose"] != df["km_gjeldende"]))
        grunn[avvik_kontrakt] = "Omega prognose (avvik mot kontrakt)"
        overstyr = er_null | kommentert | er_behold | avvik_kontrakt
    else:
        avvik_kontrakt = pd.Series(False, index=df.index)
        overstyr = er_null | kommentert | er_behold

    grunn[er_behold] = "Omega prognose (behelder gammel)"
    grunn[er_null] = "Omega prognose (satt til 0)"

    # Ein post med prefiks, men utan prognose i Omega, har ingen gammal
    # prognose å behalde. Han fell tilbake på kontraktsmengda — det skjer
    # lenger nede, etter at beregnet_prognose er teken vare på.
    utan_gammal = har_prefiks & ~satt

    df["overstyrt_av_omega"] = overstyr
    df["omega_null"] = er_null
    df["omega_kommentert"] = kommentert & ~er_null
    df["beregnet_prognose"] = df["prognose"]      # før overstyringa
    df.loc[overstyr, "prognose"] = df.loc[overstyr, "omega_prognose"]
    df.loc[overstyr, "prognosekilde"] = grunn[overstyr]
    # Omega har siste ordet: kabel- og APM-regelen gjeld ikkje lenger
    df.loc[overstyr, "kabel_til_kontrakt"] = False
    df.loc[overstyr, "apm_til_kontrakt"] = False
    df.loc[overstyr, "apm_under_kontrakt"] = False
    df.loc[overstyr, "maalebrev_ferdig"] = False
    print(f"   Omega-prognosen overstyrer {int(overstyr.sum()):,} poster "
          f"({int(er_null.sum())} satt til 0, "
          f"{int(kommentert.sum())} med kvalifisert kommentar, "
          f"{int(er_behold.sum())} beholder gammel, "
          f"{int(avvik_kontrakt.sum())} med avvik mot kontrakt)")

    # Prefiks-post utan gammal prognose i Omega: fell tilbake på
    # kontraktsmengda. Står her fordi beregnet_prognose då er teken vare på,
    # og fordi desse postane ikkje er med i overstyringa over.
    til_kontrakt_pre = utan_gammal & df["kontrakt"].notna()
    df.loc[til_kontrakt_pre, "prognose"] = df.loc[til_kontrakt_pre, "kontrakt"]
    df.loc[til_kontrakt_pre, "prognosekilde"] = (
        "Kontraktsmengde (ingen gammel prognose i Omega)")
    df.loc[til_kontrakt_pre, "kabel_til_kontrakt"] = False
    df.loc[til_kontrakt_pre, "apm_til_kontrakt"] = False
    df.loc[til_kontrakt_pre, "apm_under_kontrakt"] = False
    df.loc[til_kontrakt_pre, "maalebrev_ferdig"] = False

    if utan_gammal.any():
        print(f"   {int(til_kontrakt_pre.sum()):,} poster med prefiks "
              f"{'/'.join(OMEGA_BEHOLD_PREFIKS)} manglar prognose i Omega "
              f"— satt til kontraktsmengde")
        utan_kontrakt = int((utan_gammal & df["kontrakt"].isna()).sum())
        if utan_kontrakt:
            print(f"      [ADVARSEL] {utan_kontrakt:,} av dei manglar òg "
                  f"kontraktsmengde — den beregnede står")

    # Tekstkolonne til rapporten. Står etter Omega-overstyringa, så ho viser
    # det som faktisk gjeld. "Bekreftet" seier noko om APM-mengda si kvalitet
    # og står att sjølv om Omega overstyrer prognosen.
    df["apm_kontroll"] = ""
    df.loc[df["apm_bekrefta"], "apm_kontroll"] = "Bekreftet"
    df.loc[df["apm_til_kontrakt"], "apm_kontroll"] = "Avvik mot fakturert"
    df.loc[df["apm_under_kontrakt"], "apm_kontroll"] = "Under kontraktsmengde"

    # Sprik mellom overstyringa og det som er rekna ut frå modell/APM/kabler
    df["avvik_mot_beregnet"] = df["prognose"] - df["beregnet_prognose"]
    med_grunnlag = df["overstyrt_av_omega"] & df["beregnet_prognose"].notna() & (
        df["modell"].notna() | df["apm"].notna() | df["kabel"].notna())
    avvik_kr = (df["avvik_mot_beregnet"] * df["enhetspris"]).abs()
    avvik_pst = (df["avvik_mot_beregnet"] / df["beregnet_prognose"].replace(0, pd.NA)
                 ).abs() * 100
    df["varsel"] = ""
    df.loc[med_grunnlag & ((avvik_pst > VARSEL_AVVIK_PST)
                           | (avvik_kr > VARSEL_AVVIK_KR)), "varsel"] = "Vurder på nytt"

    sikkerhet = df.apply(prognose_sikkerhet, axis=1)
    df["sikkerhet"] = [t[0] for t in sikkerhet]
    df["sikkerhet_grunn"] = [t[1] for t in sikkerhet]
    try:
        import feil_poster
        for _, r in df[df["enhetsstatus"] == "avvik"].iterrows():
            feil_poster.registrer(
                feil_poster.FEIL_ENHET, kilde="Modell", post=r["post"],
                detalj=f"Enhet i modell: {r['modellenhet']}, posten: {r['enhet']} "
                       f"— modellmengden er holdt utenfor prognosen")
    except Exception as exc:          # feillista skal aldri velte prognosen
        print(f"   [ADVARSEL] Feillista (enhet): {exc}")

    df["opsjon"] = df["post"].astype(str).str.strip().isin(opsjon or set())
    df.loc[df["opsjon"], "sikkerhet_grunn"] = (
        df.loc[df["opsjon"], "sikkerhet_grunn"] + " · OPSJON")
    print("   Sikkerhet: " + ", ".join(
        f"{n} {int((df['sikkerhet'] == n).sum()):,}" for n in SIKKERHET_NIVAA))

    df["kontraktsverdi"] = df["kv_gjeldende"].where(
        df["kv_gjeldende"].notna(), df["kontrakt"] * df["enhetspris"])
    df["prognoseverdi"] = df["prognose"] * df["enhetspris"]
    df["avvik_mengde"] = df["prognose"] - df["kontrakt"]
    df["avvik_verdi"] = df["prognoseverdi"] - df["kontraktsverdi"]
    df["avvik_pst"] = (df["avvik_mengde"] / df["kontrakt"] * 100)
    df["avvik_pst"] = df["avvik_pst"].replace([float("inf"), float("-inf")], pd.NA)

    df = df.sort_values("post", key=lambda s: s.map(_sorteringsnokkel))
    return df.reset_index(drop=True)


# =============================================================================
# 5. Excel-skriving
# =============================================================================

KOLONNER_POST = [
    ("Post", "post", 20, None),
    ("Beskrivelse", "tittel", 42, None),
    ("Fag", "fag", 7, None),
    ("Enhet", "enhet", 8, None),
    ("Enhet i modell", "modellenhet", 12, None),
    ("Enhetskontroll", "enhetsstatus", 13, None),
    ("Kontraktsmengde", "kontrakt", 16, _TAL),
    ("Mengde fra modell", "modell", 17, _TAL),
    ("APM", "apm", 14, _TAL),
    ("APM-kontroll", "apm_kontroll", 18, None),
    ("Kabelliste", "kabel", 13, _TAL),
    ("Prognose", "prognose", 14, _TAL),
    ("Omega prognose", "omega_prognose", 15, _TAL),
    ("Prognosekilde", "prognosekilde", 26, None),
    ("Sikkerhet", "sikkerhet", 12, None),
    ("Sikkerhet grunn", "sikkerhet_grunn", 30, None),
    ("Omega kommentar", "kommentar", 40, None),
    ("Avvik mengde", "avvik_mengde", 14, _TAL),
    ("Avvik %", "avvik_pst", 10, _PST),
    ("Enhetspris", "enhetspris", 13, _TAL),
    ("Kontraktsverdi", "kontraktsverdi", 16, _KR),
    ("Prognoseverdi", "prognoseverdi", 16, _KR),
    ("Avvik verdi", "avvik_verdi", 15, _KR),
    ("Rapportert", "rapportert", 14, _TAL),
    ("APM-kilde", "apm_kilde", 22, None),
]


def _skriv_tabell(ws, df, kolonner, tabellnavn, stil="TableStyleMedium2", startrad=1):
    for j, (hdr, _, bredde, _) in enumerate(kolonner, 1):
        c = ws.cell(row=startrad, column=j, value=hdr)
        c.font, c.fill, c.alignment = _HDR_FONT, _HDR_FILL, _TOP
        ws.column_dimensions[get_column_letter(j)].width = bredde

    for i, (_, rad) in enumerate(df.iterrows(), start=startrad + 1):
        for j, (_, nokkel, _, fmt) in enumerate(kolonner, 1):
            v = rad.get(nokkel)
            if v is None or (isinstance(v, float) and pd.isna(v)) or (
                    not isinstance(v, (int, float)) and pd.isna(v)):
                v = None
            elif not isinstance(v, (int, float)):
                v = str(v)
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.alignment = _CAL10, _TOP
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt
            if nokkel == "sikkerhet" and v in _SIKKERHET_FILL:
                c.fill = _SIKKERHET_FILL[v]

    ws.freeze_panes = ws.cell(row=startrad + 1, column=1)
    if len(df):
        ref = f"A{startrad}:{get_column_letter(len(kolonner))}{startrad + len(df)}"
        tbl = Table(displayName=tabellnavn, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name=stil, showRowStripes=True)
        ws.add_table(tbl)
    else:
        ws.auto_filter.ref = ws.dimensions


def _skriv_forside(ws, df, omega, meta):
    def sum_(kol, maske=None):
        s = df.loc[maske, kol] if maske is not None else df[kol]
        return float(pd.to_numeric(s, errors="coerce").fillna(0).sum())

    ny = df["grunnlag"] == GRUNNLAG_NY

    kontrakt_signert = float(omega["kv_opprinnelig"].fillna(0).sum())
    kontrakt_endring = float(omega["kv_endringer"].fillna(0).sum())
    kontrakt_gjeldende = float(omega["kv_gjeldende"].fillna(0).sum())
    omega_prognose = float(omega["kv_prognose"].fillna(0).sum())

    prognose_total = sum_("prognoseverdi")
    prognose_nye = sum_("prognoseverdi", ny)

    for kol, br in (("A", 3), ("B", 62), ("C", 20), ("D", 44)):
        ws.column_dimensions[kol].width = br
    ws.sheet_view.showGridLines = False

    r = [2]

    def tittelrad(tekst, storleik=16, farge="1F4E79", kursiv=False):
        c = ws.cell(row=r[0], column=2, value=tekst)
        c.font = Font(size=storleik, bold=storleik >= 12, color=farge, italic=kursiv)
        r[0] += 1

    def blokk(tekst):
        r[0] += 1
        for kol in (2, 3, 4):
            ws.cell(row=r[0], column=kol).fill = _TIT_FILL
        c = ws.cell(row=r[0], column=2, value=tekst)
        c.font = Font(bold=True, size=11, color="1F4E79")
        r[0] += 1

    def linje(tekst, verdi, fmt=_KR, note="", fet=False, gul=False):
        ws.cell(row=r[0], column=2, value=tekst).font = Font(bold=fet, size=10)
        c = ws.cell(row=r[0], column=3, value=verdi)
        c.number_format, c.font, c.alignment = fmt, Font(bold=fet, size=10), _HOYRE
        if gul:
            c.fill = _GUL_FILL
            c.border = Border(_TYNN, _TYNN, _TYNN, _TYNN)
        if note:
            ws.cell(row=r[0], column=4, value=note).font = Font(size=9, color="808080")
        r[0] += 1

    tittelrad("PRIS- OG MENGDERAPPORT — PROGNOSE")
    tittelrad(f"{meta['prosjekt']}   |   {meta['periode']}", 11, "404040", kursiv=True)
    tittelrad(f"Generert {meta['tid']}", 9, "808080")

    blokk("KONTRAKT")
    linje("Kontraktssum ved signert kontrakt", kontrakt_signert,
          note="Omega: kontraktsverdi opprinnelig")
    linje("Kontraktsendringer etter kontraktsinngåelse", kontrakt_endring,
          note="Omega: kontraktsverdi endringer")
    linje("Gjeldende kontraktsverdi", kontrakt_gjeldende, fet=True,
          note="Omega: kontraktsverdi gjeldende")

    blokk("REVISJONSBESKRIVELSER")
    linje("Nye poster og prosesser utenfor kontraktsgrunnlaget", int(ny.sum()), fmt="#,##0",
          note="Poster i modell/Omega som mangler i NS3420- og prosesskodearket")
    linje("Prognoseverdi for disse postene", prognose_nye,
          note="Krever avtalt pris — oppdater grunnlagsfilene")
    linje("Planlagte revisjonsbeskrivelser — forventet omfang", None,
          note="FYLLES INN AV PGF", gul=True)
    linje("Planlagte revisjonsbeskrivelser — forventet kostnad", None,
          note="FYLLES INN AV PGF", gul=True)

    blokk("PROGNOSE")
    linje("Prognoseverdi — poster i kontraktsgrunnlaget", prognose_total - prognose_nye)
    linje("Prognoseverdi — nye poster", prognose_nye)
    linje("TOTAL PROGNOSE", prognose_total, fet=True)
    linje("Avvik mot gjeldende kontraktsverdi", prognose_total - kontrakt_gjeldende,
          note="Positivt tall = prognose over kontrakt")
    if omega_prognose:
        linje("Til sammenligning: prognose i Omega", omega_prognose,
              note="Omega: kontraktsverdi prognose")

    blokk("GRUNNLAG")
    for tekst, verdi in (
        ("Poster i rapporten", len(df)),
        ("  — fra NS3420-arket", int((df["grunnlag"] == GRUNNLAG_NS).sum())),
        ("  — fra prosesskodearket", int((df["grunnlag"] == GRUNNLAG_PK).sum())),
        ("  — nye poster", int(ny.sum())),
        ("Poster med mengde fra modell", int(df["modell"].notna().sum())),
        ("Poster med APM-mengde", int(df["apm"].notna().sum())),
        ("Kabelposter fra kabelliste", int(df["kabel"].notna().sum())),
        ("Poster uten grunnlag (kontraktsmengde brukt)",
         int((df["modell"].isna() & df["apm"].isna() & df["kabel"].isna()).sum())),
        ("Poster uten enhetspris", int(df["enhetspris"].isna().sum())),
        ("Poster med enhetsavvik modell/post (se eget ark)",
         int((df["enhetsstatus"] == "avvik").sum())),
        ("Poster der Omega-prognosen overstyrer (se eget ark)",
         int(df["overstyrt_av_omega"].sum())),
    ):
        linje(tekst, verdi, fmt="#,##0")

    blokk("METODE")
    for t in (
        "Prognose = største av mengde fra modell og APM-mengde når minst én har tall.",
        "Prognose = kontraktsmengde når verken modell eller APM har tall.",
        "Mengde fra modell er hentet fra mengde_oversikt (ARM/RIB via produksjonsenhet,",
        "øvrige fag via FOB_Mengde). APM-mengde er summert fra APM-målebrevene i PDF.",
        "APM-mengden kontrolleres mot de andre kildene: er den lik mengden fra modell",
        f"eller kontraktsmengden (innenfor {APM_BEKREFTA_PST:g} %), regnes den som",
        "bekreftet og blir stående.",
        *((f"Ellers, avviker den mer enn {APM_MAKS_AVVIK_RAPPORTERT_PST:g} % fra "
           "mengden som er fakturert hittil,",
           "brukes kontraktsmengden i stedet. Er ingenting fakturert ennå, er det",
           "ingenting å sammenligne med, og APM-mengden blir stående.")
          if APM_MAKS_AVVIK_RAPPORTERT_PST is not None else ()),
        *((f"Er APM-mengden mer enn {APM_MAKS_UNDER_KONTRAKT_PST:g} % lavere enn "
           "kontraktsmengden,",
           "brukes kontraktsmengden i stedet.")
          if APM_MAKS_UNDER_KONTRAKT_PST is not None else ()),
        "Kolonnen Sikkerhet sier hvor sikker prognosen er:",
        "  Helt sikker: satt til 0 i Omega, eller målebrevet er avsluttet.",
        f"  Sikker: minst to kilder (modell, APM, kabelliste, kontrakt) er innenfor "
        f"{SIKKERHET_LIK_PST:g} % av prognosen,",
        "  eller Omega-prognosen er kommentert eller støttet av en kilde.",
        "  Usikker: kontraktsmengden er satt inn fordi APM/kabelliste spriker, Omega-",
        f"  prognosen skal vurderes på nytt, enhetsavvik, eller kildene spriker mer enn "
        f"{SIKKERHET_SPRIK_PST:g} %.",
        "  Middels: resten, f.eks. bare én kilde eller bare kontraktsmengde.",
        f"  Står «{OPSJON_TEKST}» i kontraktsteksten (All tekst), er det lagt til i "
        "Sikkerhet grunn.",
        "Rund sum-poster får prognose fra kontrakt/APM — modellen teller ikke antall.",
        "Har modellen målt i en annen enhet enn posten gjøres opp i (f.eks. areal på",
        "et element som avregnes per stk), holdes modellmengden utenfor prognosen og",
        "posten listes på arket Enhetsavvik. kg og tonn regnes om automatisk.",
        "Er malebrevet ferdig (SRES er sann i Omega), er mengden gjort opp, og",
        "prognosen settes lik den fakturerte (rapporterte) mengden.",
        *((f"Poster med prosesskode som starter på "
           f"{'/'.join(OMEGA_BEHOLD_PREFIKS)} beholder prognosen som allerede",
           "står i Omega, uansett hva som er beregnet. Mangler posten prognose i",
           "Omega, brukes kontraktsmengden.")
          if OMEGA_BEHOLD_PREFIKS else ()),
        "Er prognosen i Omega satt til 0, eller har posten en kommentar som viser at",
        "prognosen er lagt inn manuelt, gjelder Omega-prognosen framfor den beregnede.",
        "Disse to gjelder alltid. I tillegg gjelder, for september 2026, at enhver",
        "Omega-prognose som er ulik kontraktsmengden regnes som avtalt.",
        "En kommentar overstyrer bare når den sier at posten er utgått, eller at",
        "mengden er regulert som følge av en EO eller PF. Andre kommentarer gjør ikke noe.",
        "Overstyringen står selv om mengdene endrer seg, men markeres Vurder på nytt",
        f"når spriket mot det beregnede er over {VARSEL_AVVIK_PST:g} % eller "
        f"{VARSEL_AVVIK_KR:,.0f} kr.",
        "Kabler finnes ikke som objekt i modellen: meterposter får prognose fra",
        "lengdene i kabellistene, koblet på kabelkonstruksjon, ledertall og tverrsnitt.",
        *((f"Avviker prognoseverdien fra kabellisten mer enn {KABEL_MAKS_AVVIK_KR:,.0f} kr",
           "fra kontraktsverdien (opp eller ned), brukes kontraktsmengden i stedet.")
          if KABEL_MAKS_AVVIK_KR is not None else ()),
        "Kontraktsverdiene er hentet direkte fra Omega, uten totalsumraden.",
    ):
        ws.cell(row=r[0], column=2, value=t).font = Font(size=9)
        r[0] += 1


def _skriv_revisjonsark(ws, df, bim):
    """Nye poster/prosesser og utgåtte element."""
    ws.column_dimensions["A"].width = 3
    ws.sheet_view.showGridLines = False

    ny = df[df["grunnlag"] == GRUNNLAG_NY].sort_values("prognoseverdi", ascending=False)
    ut = df[
        df["kontrakt"].notna() & (df["kontrakt"] != 0)
        & df["modell"].isna() & df["apm"].isna()
        & (df["grunnlag"] != GRUNNLAG_NY)
    ].sort_values("kontraktsverdi", ascending=False)

    rad = 1
    ws.cell(row=rad, column=1, value="NYE POSTER OG PROSESSER").font = Font(
        bold=True, size=12, color="1F4E79")
    rad += 1
    ws.cell(row=rad, column=1,
            value="Poster med mengde i modell eller APM som ikke finnes i "
                  "NS3420-arket eller prosesskodearket. Disse kommer fra "
                  "revisjonsbeskrivelser og trenger avtalt pris.").font = Font(
        size=9, color="808080")
    rad += 2
    _skriv_tabell(ws, ny, KOLONNER_POST, "Tabell_Nye_poster",
                  stil="TableStyleMedium3", startrad=rad)

    rad += len(ny) + 3
    ws.cell(row=rad, column=1, value="MULIG UTGÅTTE ELEMENTER").font = Font(
        bold=True, size=12, color="1F4E79")
    rad += 1
    ws.cell(row=rad, column=1,
            value="Poster med kontraktsmengde, men uten mengde i modell og uten "
                  "APM. Kan være utgått i revisjonsbeskrivelse, eller ikke "
                  "prosjektert/utført ennå.").font = Font(size=9, color="808080")
    rad += 2
    _skriv_tabell(ws, ut, KOLONNER_POST, "Tabell_Utgatte",
                  stil="TableStyleMedium7", startrad=rad)


def _skriv_kontroll(ws, df, omega, bim):
    rader = [
        ("Poster totalt", len(df)),
        ("Poster fra NS3420-arket", int((df["grunnlag"] == GRUNNLAG_NS).sum())),
        ("Poster fra prosesskodearket", int((df["grunnlag"] == GRUNNLAG_PK).sum())),
        ("Nye poster (revisjonsbeskrivelser)", int((df["grunnlag"] == GRUNNLAG_NY).sum())),
        ("Poster i Omega", len(omega)),
        ("Modellposter uten Omega-match (BIM_utan_Omega)", len(bim)),
        ("Poster med mengde fra modell", int(df["modell"].notna().sum())),
        ("Poster med APM-mengde", int(df["apm"].notna().sum())),
        ("Prognose fra modell", int((df["prognosekilde"] == "Modell").sum())),
        ("Prognose fra APM", int((df["prognosekilde"] == "APM").sum())),
        ("Prognose der modell = APM", int((df["prognosekilde"] == "Modell = APM").sum())),
        ("Prognose fra kontraktsmengde",
         int(df["prognosekilde"].str.startswith("Kontraktsmengde").sum())),
        ("Poster uten grunnlag", int((df["prognosekilde"] == "Mangler grunnlag").sum())),
        ("Rund sum-poster (modell holdt utenfor)", int(df["rundsum"].sum())),
        ("Poster der modellens enhet ikke matcher posten (modell holdt utenfor)",
         int((df["enhetsstatus"] == "avvik").sum())),
        ("Poster der modellmengden er regnet om (kg/tonn)",
         int((df["enhetsstatus"] == "omregnet").sum())),
        ("Poster uten kjent enhet i modellen",
         int((df["enhetsstatus"] == "ukjent").sum())),
        ("Poster der Omega-prognosen overstyrer",
         int(df["overstyrt_av_omega"].sum())),
        ("  - herav satt til 0 i Omega",
         int((df["prognosekilde"] == "Omega prognose (satt til 0)").sum())),
        ("  - herav med kommentar om utgått eller mengderegulering",
         int((df["kommentargrunn"] != "").sum())),
        ("Poster med kommentar i Omega som IKKE utløser overstyring",
         int(((df["kommentar"].astype(str).str.strip() != "")
              & (df["kommentargrunn"] == "")).sum())),
        ("  - herav med prognose ulik kontraktsmengde",
         int((df["prognosekilde"] == "Omega prognose (avvik mot kontrakt)").sum())),
        (f"  - herav som bør vurderes på nytt (sprik > {VARSEL_AVVIK_PST:g} % "
         f"eller {VARSEL_AVVIK_KR:,.0f} kr mot beregnet)",
         int((df["varsel"] != "").sum())),
        ("Poster uten enhetspris", int(df["enhetspris"].isna().sum())),
        (f"Poster med prefiks {'/'.join(OMEGA_BEHOLD_PREFIKS)} som beholder "
         f"prognosen fra Omega" if OMEGA_BEHOLD_PREFIKS
         else "Prefiks-regelen er av",
         int((df["prognosekilde"] == "Omega prognose (behelder gammel)").sum())),
        ("  - herav uten prognose i Omega, satt til kontraktsmengde",
         int((df["prognosekilde"]
              == "Kontraktsmengde (ingen gammel prognose i Omega)").sum())),
        ("Poster med ferdig målebrev (prognose = fakturert mengde)",
         int(df["maalebrev_ferdig"].sum())),
        ("APM-poster bekreftet av modell/kontraktsmengde",
         int(df["apm_bekrefta"].sum())),
        (f"APM-poster satt til kontraktsmengde (avvik > "
         f"{APM_MAKS_AVVIK_RAPPORTERT_PST:g} % mot fakturert)"
         if APM_MAKS_AVVIK_RAPPORTERT_PST is not None
         else "APM-poster satt til kontraktsmengde (kontroll av)",
         int(df["apm_til_kontrakt"].sum())),
        (f"APM-poster satt til kontraktsmengde (APM > "
         f"{APM_MAKS_UNDER_KONTRAKT_PST:g} % under kontrakt)"
         if APM_MAKS_UNDER_KONTRAKT_PST is not None
         else "APM-poster under kontrakt (kontroll av)",
         int(df["apm_under_kontrakt"].sum())),
        *((f"Prognose {n.lower()}", int((df["sikkerhet"] == n).sum()))
          for n in SIKKERHET_NIVAA),
        (f"Poster med {OPSJON_TEKST} i kontraktsteksten", int(df["opsjon"].sum())),
        ("Kabelposter med mengde fra kabelliste", int(df["kabel"].notna().sum())),
        ("Kabelposter satt til kontraktsmengde (stort avvik)",
         int(df["kabel_til_kontrakt"].sum())),
        ("Poster der prognose > 2 x kontraktsmengde",
         int((df["prognose"] > 2 * df["kontrakt"].where(df["kontrakt"] > 0)).sum())),
    ]
    kontroll = pd.DataFrame(rader, columns=["kontroll", "antall"])
    _skriv_tabell(ws, kontroll,
                  [("Kontroll", "kontroll", 52, None), ("Antall", "antall", 14, "#,##0")],
                  "Tabell_Kontroll", stil="TableStyleMedium4")


# =============================================================================
# 6. Hovudflyt
# =============================================================================

def finn_kabellister(*mapper) -> list[str]:
    """Finn kabellister (PF-*-LC-*.xlsx) i dei oppgitte mappene."""
    funne: list[str] = []
    for mappe in mapper:
        if not mappe or not os.path.isdir(mappe):
            continue
        for navn in sorted(os.listdir(mappe)):
            n = navn.lower()
            if n.endswith((".xlsx", ".xls")) and not n.startswith("~$") \
                    and "kabelliste" in n:
                funne.append(os.path.join(mappe, navn))
    return funne


def _oppl_kolnamn(verdi) -> str:
    """Normaliser kolonnenamn: 'Forecast Qty' og 'ForecastQty' -> 'forecastqty'."""
    return re.sub(r"[\s_]+", "", str(verdi or "")).lower()


def skriv_omega_opplasting(
    df: pd.DataFrame,
    kilde_sti: str,
    ut_sti: str | None = None,
) -> str:
    """Lag ein kopi av Omega-opplastingsarket med den nye prognosen.

    Kjelda — MAL_omega_opplastning.xlsx eller eit ekte BOQNT-uttrekk med
    fleire tusen rader — vert kopiert som ho er, slik at alle kolonnar,
    rader og format står urørte. Berre ForecastQty vert skriven, og berre
    på rader der prosesskoden i Name finst i prognosen.

    Parameters
    ----------
    df        : prognosen frå bygg_prognose (treng kolonnane post og prognose)
    kilde_sti : opplastingsarket som skal kopierast
    ut_sti    : ny fil (standard: ved sida av kjelda, med tidsstempel)

    Returns
    -------
    Stien til den nye fila.
    """
    if ut_sti is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = re.sub(r"^MAL[_\- ]*", "",
                      os.path.splitext(os.path.basename(kilde_sti))[0],
                      flags=re.IGNORECASE)
        ut_sti = os.path.join(os.path.dirname(os.path.abspath(kilde_sti)),
                              f"{base}_{ts}.xlsx")

    try:
        from lagring import vent_til_skrivbar
        vent_til_skrivbar(ut_sti)
    except ImportError:
        pass

    # Kopier heile fila — då står alt vi ikkje skriv til att urørt
    shutil.copyfile(kilde_sti, ut_sti)
    wb = openpyxl.load_workbook(ut_sti)
    ws = wb.active

    # Finn rada som har både Name og ForecastQty
    hdr_rad = kol_post = kol_prog = None
    for r in range(1, min(ws.max_row, 20) + 1):
        funne: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            n = _oppl_kolnamn(ws.cell(row=r, column=c).value)
            if n in (OMEGA_OPPL_KOL_POST, OMEGA_OPPL_KOL_PROGNOSE):
                funne.setdefault(n, c)
        if len(funne) == 2:
            hdr_rad = r
            kol_post = funne[OMEGA_OPPL_KOL_POST]
            kol_prog = funne[OMEGA_OPPL_KOL_PROGNOSE]
            break

    if hdr_rad is None:
        raise ValueError(
            f"Fann ikkje kolonnane '{OMEGA_OPPL_KOL_POST}' og "
            f"'{OMEGA_OPPL_KOL_PROGNOSE}' i {os.path.basename(kilde_sti)}")

    prognose = {str(p).strip(): float(v)
                for p, v in zip(df["post"], df["prognose"])
                if pd.notna(v) and str(p).strip()}

    oppdatert = 0
    funne_postar: set[str] = set()
    fleire_gonger: dict[str, int] = {}
    ukjende: set[str] = set()

    for r in range(hdr_rad + 1, ws.max_row + 1):
        post = str(ws.cell(row=r, column=kol_post).value or "").strip()
        if not post:
            continue
        if post not in prognose:
            # Hopp over header- og typerader; berre reelle kodar vert talde
            if post.count(".") >= 2:
                ukjende.add(post)
            continue
        ws.cell(row=r, column=kol_prog).value = round(prognose[post],
                                                     OMEGA_OPPL_DESIMALAR)
        oppdatert += 1
        if post in funne_postar:
            fleire_gonger[post] = fleire_gonger.get(post, 1) + 1
        funne_postar.add(post)

    wb.save(ut_sti)

    manglar = set(prognose) - funne_postar
    print(f"   Omega-opplasting: {oppdatert:,} rader oppdatert "
          f"({len(funne_postar):,} poster) i {os.path.basename(ut_sti)}")
    if ukjende:
        print(f"      {len(ukjende):,} poster i arket finst ikkje i prognosen "
              f"— ForecastQty står urørt")
    if manglar:
        print(f"      {len(manglar):,} poster i prognosen finst ikkje i arket")
    if fleire_gonger:
        print(f"      [ADVARSEL] {len(fleire_gonger):,} poster står på fleire "
              f"rader i arket og fekk same prognose på kvar. Summen kan bli "
              f"for høg: {', '.join(sorted(fleire_gonger)[:5])}"
              + (" …" if len(fleire_gonger) > 5 else ""))

    return ut_sti


def lag_prognose(
    oversikt_excel: str,
    omega_path: str,
    isy_path: str | list[str],
    kabelliste_stiar: list[str] | None = None,
    omega_overstyrer: bool = OMEGA_OVERSTYRER_STANDARD,
    bruk_forrige_kommentar: bool = False,
    periode: str = "",
    prosjekt: str = "K7 — Lysaker og Vækerø stasjon",
    utdir: str | None = None,
    omega_mal: str | None = None,
    omega_ut: str | None = None,
) -> str:
    """
    Lag prognosearket ut frå ein ferdig mengde_oversikt_*.xlsx.

    Parameters
    ----------
    oversikt_excel : sti til mengde_oversikt_*.xlsx frå mengde_oversikt.lag_oversikt
    omega_path     : sti til Omega-uttrekket
    isy_path       : ISY Linker-fil(er) — NS3420-arket og/eller prosesskodearket
    kabelliste_stiar : kabellister (PF-*-LC-*.xlsx). None = leit i mappa til
                     Omega-fila og i mappa til oversikta.
    omega_overstyrer : la prognosen frå Omega gå framfor den utrekna der han
                     er sett bevisst (0, kommentert, eller ulik kontrakts-
                     mengda). Gjeld september 2026 — sett False seinare.
    bruk_forrige_kommentar : rekn òg "Kommentar fra forrige delmålebrev" som
                     kommentar. Av som standard.
    periode        : tekst på forsida, t.d. "September 2026"
    utdir          : utmappe (standard: same mappe som oversikt_excel)
    omega_mal      : Omega-opplastingsark (MAL_omega_opplastning.xlsx eller eit
                     ekte BOQNT-uttrekk). Vert kopiert til utmappa med
                     ForecastQty oppdatert. None = hopp over.
    omega_ut       : eksplisitt filnamn for opplastingsarket. None = automatisk
                     namn i utmappa. Brukt av GUI-et, som treng å vite stien
                     for å kunne opne fila etterpå.

    Returns
    -------
    Sti til den ferdige prognosefila.
    """
    utdir = utdir or os.path.dirname(os.path.abspath(oversikt_excel))

    print("[1/4] Leser mengde-oversikt …")
    oversikt = les_oversikt(oversikt_excel)
    bim = les_bim_utan_omega(oversikt_excel)

    print("[2/4] Leser Omega og kontraktsgrunnlag …")
    omega = les_kontraktsverdier(omega_path)
    grunnlag = les_grunnlag(isy_path)
    opsjon = les_opsjonsposter(isy_path)

    if kabelliste_stiar is None:
        kabelliste_stiar = finn_kabellister(
            os.path.dirname(os.path.abspath(omega_path)),
            os.path.dirname(os.path.abspath(oversikt_excel)),
        )
    kabel: dict = {}
    if kabelliste_stiar:
        try:
            from kabel_prognose import les_kabelmengder
            kabel = les_kabelmengder(kabelliste_stiar, isy_path)
        except Exception as exc:      # kabellista skal aldri velte rapporten
            print(f"   [ADVARSEL] Kabelprognosen feila: {exc}")
    else:
        print("   Ingen kabellister funne — kabelprognose hoppa over")

    print("[3/4] Beregner prognose …")
    df = bygg_prognose(oversikt, omega, grunnlag, kabel,
                       omega_overstyrer=omega_overstyrer,
                       bruk_forrige_kommentar=bruk_forrige_kommentar,
                       opsjon=opsjon)

    print("[4/4] Skriver rapport …")
    wb = openpyxl.Workbook()
    ws_forside = wb.active
    ws_forside.title = "Forside"
    ws_ns = wb.create_sheet("NS3420")
    ws_pk = wb.create_sheet("Prosesskoden")
    ws_rev = wb.create_sheet("Revisjonsbeskrivelser")
    ws_alle = wb.create_sheet("Alle poster")
    ws_kon = wb.create_sheet("Kontroll")

    meta = {
        "prosjekt": prosjekt,
        "periode": periode or datetime.now().strftime("%B %Y"),
        "tid": datetime.now().strftime("%d.%m.%Y %H:%M"),
    }
    _skriv_forside(ws_forside, df, omega, meta)
    _skriv_tabell(ws_ns, df[df["grunnlag"] == GRUNNLAG_NS], KOLONNER_POST,
                  "Tabell_NS3420", stil="TableStyleMedium2")
    _skriv_tabell(ws_pk, df[df["grunnlag"] == GRUNNLAG_PK], KOLONNER_POST,
                  "Tabell_Prosesskoden", stil="TableStyleMedium5")
    _skriv_revisjonsark(ws_rev, df, bim)
    ws_om = wb.create_sheet("Omega overstyrer")
    over = df[df["overstyrt_av_omega"]].copy()
    over["_v"] = (over["avvik_mot_beregnet"].fillna(0)
                  * over["enhetspris"].fillna(0)).abs()
    over = over.sort_values(["varsel", "_v"], ascending=[False, False])
    _skriv_tabell(ws_om, over,
                  [("Post", "post", 20, None), ("Beskrivelse", "tittel", 40, None),
                   ("Enhet", "enhet", 8, None),
                   ("Kontraktsmengde", "kontrakt", 16, _TAL),
                   ("Beregnet fra modell/APM", "modell", 17, _TAL),
                   ("APM", "apm", 13, _TAL),
                   ("Omega prognose", "omega_prognose", 15, _TAL),
                   ("Prognose", "prognose", 14, _TAL),
                   ("Prognosekilde", "prognosekilde", 26, None),
                   ("Beregnet prognose", "beregnet_prognose", 16, _TAL),
                   ("Avvik mot beregnet", "avvik_mot_beregnet", 16, _TAL),
                   ("Varsel", "varsel", 14, None),
                   ("Prognoseverdi", "prognoseverdi", 16, _KR),
                   ("Avvik verdi", "avvik_verdi", 15, _KR),
                   ("Kommentargrunn", "kommentargrunn", 20, None),
                   ("Omega kommentar", "kommentar", 50, None)],
                  "Tabell_Omega_overstyrer", stil="TableStyleMedium10")

    ws_enh = wb.create_sheet("Enhetsavvik")
    avvik = df[df["enhetsstatus"] == "avvik"].copy()
    avvik["_v"] = (avvik["modell"] * avvik["enhetspris"]).fillna(0)
    avvik = avvik.sort_values("_v", ascending=False)
    _skriv_tabell(ws_enh, avvik,
                  [("Post", "post", 20, None), ("Beskrivelse", "tittel", 42, None),
                   ("Fag", "fag", 7, None),
                   ("Enhet i post", "enhet", 11, None),
                   ("Enhet i modell", "modellenhet", 12, None),
                   ("Kontraktsmengde", "kontrakt", 16, _TAL),
                   ("Mengde fra modell", "modell", 17, _TAL),
                   ("APM", "apm", 13, _TAL),
                   ("Prognose", "prognose", 14, _TAL),
                   ("Prognosekilde", "prognosekilde", 26, None),
                   ("Enhetspris", "enhetspris", 13, _TAL)],
                  "Tabell_Enhetsavvik", stil="TableStyleMedium8")

    ws_kab = wb.create_sheet("Kabler")
    kabler = df[df["kabel"].notna()].sort_values("post", key=lambda s: s.map(_sorteringsnokkel))
    _skriv_tabell(ws_kab, kabler,
                  [("Post", "post", 20, None), ("Beskrivelse", "tittel", 40, None),
                   ("Enhet", "enhet", 8, None),
                   ("Kontraktsmengde", "kontrakt", 16, _TAL),
                   ("Fra kabelliste", "kabel", 15, _TAL),
                   ("APM", "apm", 13, _TAL),
                   ("Prognose", "prognose", 14, _TAL),
                   ("Enhetspris", "enhetspris", 13, _TAL),
                   ("Prognoseverdi", "prognoseverdi", 16, _KR),
                   ("Avvik verdi", "avvik_verdi", 15, _KR),
                   ("Prognosekilde", "prognosekilde", 34, None),
                   ("Kabeltyper", "kabeltyper", 46, None)],
                  "Tabell_Kabler", stil="TableStyleMedium9")
    _skriv_tabell(ws_alle, df, KOLONNER_POST, "Tabell_Alle", stil="TableStyleMedium6")
    _skriv_kontroll(ws_kon, df, omega, bim)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    stikk = re.sub(r"[^\w]+", "_", periode).strip("_") or "prognose"
    utfil = os.path.join(utdir, f"prognose_{stikk}_{ts}.xlsx")

    try:
        from lagring import vent_til_skrivbar
        vent_til_skrivbar(utfil)
    except ImportError:
        pass
    wb.save(utfil)

    print(f"\n[OK] Prognose lagret: {utfil}")
    print(f"     Poster:               {len(df):,}")
    print(f"     Nye poster:           {int((df['grunnlag'] == GRUNNLAG_NY).sum()):,}")
    print(f"     Gjeldende kontrakt:   {omega['kv_gjeldende'].fillna(0).sum():>18,.0f} kr")
    print(f"     Total prognose:       {df['prognoseverdi'].fillna(0).sum():>18,.0f} kr")

    if omega_mal:
        try:
            oppl = skriv_omega_opplasting(
                df, omega_mal,
                omega_ut or os.path.join(
                    utdir, f"omega_opplastning_{stikk}_{ts}.xlsx"))
            print(f"     Omega-opplasting:     {oppl}")
        except Exception as exc:   # skal aldri velte sjolve prognosen
            print(f"   [ADVARSEL] Omega-opplastingsarket feila: {exc}")

    return utfil


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="Lag prognoseark frå mengde_oversikt")
    p.add_argument("oversikt", help="mengde_oversikt_*.xlsx")
    p.add_argument("omega", help="Omega-uttrekk (.xlsx)")
    p.add_argument("isy", nargs="+", help="ISY Linker-filer (NS3420 / prosesskoden)")
    p.add_argument("--periode", default="", help='T.d. "September 2026"')
    p.add_argument("--omega-mal", dest="omega_mal", default=None,
                   help="Omega-opplastingsark som skal kopierast med ny prognose")
    a = p.parse_args()
    lag_prognose(a.oversikt, a.omega, a.isy, periode=a.periode,
                 omega_mal=a.omega_mal)
