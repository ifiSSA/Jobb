# -*- coding: utf-8 -*-
"""
Lager Excel-oversikt over alle Omega-poster med:
  - Mengde fra BIM-modell  (ARM/RIB via PE-matching, andre fag via FOB_Mengde)
  - Fag-kode (3 bokstavar) henta frå Path-kolonnen i CSV
  - Enhetspris og beskrivelse fra ISY Linker
  - Kontraktsmengde og rapportert mengde fra Omega

Tre ark:
  1. I_modell       – Omega-poster som finnes i BIM-modellen
  2. Ikke_i_modell  – Omega-poster som IKKE finnes i BIM-modellen
  3. ISY_Omega_avvik – Postar der ISY-mengde ≠ Omega-kontraktsmengde

Bruk — kjør direkte:
    python mengde_oversikt.py

Eller importer og bruk arm_rib_excel for å hoppe over ARM/RIB-analysen:
    from mengde_oversikt import lag_oversikt
    lag_oversikt(arm_rib_excel=r"C:\\...\\resultat_mengdeuttak_....xlsx")
"""

from __future__ import annotations

import os
import re
import sys
import warnings

from datetime import datetime

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ── Filstiar ──────────────────────────────────────────────────────────────────
CSV_PATH   = r"C:\Users\ifi\Downloads\selection-basket (34).csv"
OMEGA_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang\Omega_K7_April.xlsx"
)
ISY_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\Mengdelistpost NS3420 isylinker.xlsx"
)
APM_EXCEL = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang\Maalebrev_samlet.xlsx"
)
from lagring import EXCEL_MAPPE as UTDIR   # <kodemappa>/excel/ — ikkje i git

# ── Kolonnenavn i Sammendrag-arket frå Mengdeuttak_Armering ──────────────────
SAM_ARM_POST = "OMEGA-post: Armering"
SAM_BET_POST = "OMEGA-post: Betong"
# Armeringsvekt-kolonna heiter "Armeringsvekt modell (kg)" etter at Total_Vekt
# vart lagt om til å bruke ARM-modellen. Eldre rapportar har "Armeringsvekt (kg)"
# — begge namn vert prøvd, i denne rekkjefølgja.
SAM_ARM_VEKT = ("Armeringsvekt modell (kg)", "Armeringsvekt (kg)")
SAM_VOL      = "Volum (m3)"

# ── Omega-kolonnar ────────────────────────────────────────────────────────────
OM_KODE    = "Navn (Prosesskode)"
OM_TITTEL  = "Tittel"
OM_ENHET   = "Enhet"
OM_EPREIS  = "Enhetspris"
OM_KONTR   = "Kontraktsmengde"
OM_RAPPORT = "Rapportert mengde (Kumulativ)"

# ── CSV-kolonnar (etter / → . normalisering) ──────────────────────────────────
FOB_MENG_KOL = "FOB_Info.FOB_Mengdelistepost"
FOB_MENG     = "FOB_Info.FOB_Mengde"
_ENHET_KOL   = "FOB_Info.FOB_Mengdeenhet"

# Toleranse for ISY vs Omega mengde-samanlikning (relativ, 0.1 %)
_AVVIK_TOL = 0.001

# ── Område-deteksjon (prosesskode-prefiks) ────────────────────────────────────
_LYSAKER_PREFIX  = ("1303", "1407")
_VAAKEROO_PREFIX = ("1302",)


# =============================================================================
# Hjelpefunksjon: Fag-kode frå Path
# =============================================================================

# Regex: finn disiplinkode i IFC-filnamn med mønster _DISC_[SUB_]K7-LP
# Døme: "_MEK_ROR_K7-LP" → "MEK"
#        "_RIB_ARM_K7-LP" → "RIB"
#        "_RIE_K7-LP"     → "RIE"
_FAG_RE    = re.compile(r"_([A-Z]{3,4})_(?:[A-Z]{3,4}_)?K7-LP", re.IGNORECASE)
_APM_NR_RE = re.compile(r"APM\s*(\d+)", re.IGNORECASE)


def _fag_fra_sti(sti: str) -> str:
    """Ekstraher 3-bokstavs fagkode frå filsti. Tom streng om ikkje funnen."""
    m = _FAG_RE.search(str(sti))
    return m.group(1).upper() if m else ""


# =============================================================================
# Hjelpefunksjonar for område-deteksjon
# =============================================================================

def _detect_omraade(df: pd.DataFrame) -> tuple[bool, bool]:
    """Returner (har_lysaker, har_vaakeroo) basert på FOB_Mengdelistepost-prefiks."""
    koder = df[FOB_MENG_KOL].dropna().astype(str).str.strip()
    koder = koder[~koder.isin(["", "nan"])]
    return (
        bool(koder.str.startswith(_LYSAKER_PREFIX).any()),
        bool(koder.str.startswith(_VAAKEROO_PREFIX).any()),
    )


def _har_arm_og_rib(df: pd.DataFrame) -> bool:
    """Sjekk om DataFrame har både ARM- og RIB-element (basert på Path)."""
    from csv_til_armrib import ARM_PATH_TOKEN, RIB_PATH_TOKEN
    path = df["Path"].astype(str)
    har_arm = path.str.contains(ARM_PATH_TOKEN, regex=False, na=False).any()
    har_rib = (
        path.str.contains(RIB_PATH_TOKEN, regex=False, na=False)
        & ~path.str.contains(ARM_PATH_TOKEN, regex=False, na=False)
    ).any()
    return har_arm and har_rib


def _filtrer_til_omraade(df: pd.DataFrame, prefiks: tuple[str, ...]) -> pd.DataFrame:
    """Filtrer DataFrame til rader som høyrer til eit område.

    Rader med FOB_Mengdelistepost-prefiks → filtrer på prefiks.
    Rader utan mengdelistepost (t.d. ARM-element) → tas alltid med,
    sidan PE-matchinga i Mengdeuttak_Armering handterer områdeskillet.
    """
    from csv_til_armrib import ARM_PATH_TOKEN
    kode_s = df[FOB_MENG_KOL].fillna("").astype(str).str.strip()
    har_kode  = ~kode_s.isin(["", "nan"])
    rett_kode = kode_s.str.startswith(prefiks) & har_kode
    arm_mask  = df["Path"].astype(str).str.contains(ARM_PATH_TOKEN, regex=False, na=False)
    return df[rett_kode | (~har_kode & arm_mask)].copy()


def _rename_med_suffix(sti: str, suffix: str) -> str:
    """Legg til suffix (t.d. 'Lysaker') i filnamnet like føre .xlsx."""
    base, ext = os.path.splitext(sti)
    ny = f"{base}_{suffix}{ext}"
    os.rename(sti, ny)
    return ny


def kjor_arm_rib_per_omraade(
    csv_path: str | list[str],
    omega_path: str,
    fob_mengde_mode: bool = False,
    df: pd.DataFrame | None = None,
) -> list[str]:
    """Kjør ARM/RIB-analyse (Mengdeuttak_Armering), evt. separat per område.

    Returnerer liste med Excel-stiar — éin per område viss begge har ARM+RIB,
    elles éin kombinert fil.
    """
    from csv_til_armrib import kjor_mengdeuttak, les_csv

    df_full = df if df is not None else les_csv(csv_path, lp_only=True)
    har_ly, har_vk = _detect_omraade(df_full)

    if har_ly and har_vk:
        df_ly = _filtrer_til_omraade(df_full, _LYSAKER_PREFIX)
        df_vk = _filtrer_til_omraade(df_full, _VAAKEROO_PREFIX)
        ly_ok = _har_arm_og_rib(df_ly)
        vk_ok = _har_arm_og_rib(df_vk)

        if ly_ok and vk_ok:
            print(
                "[INFO] Begge område (Lysaker + Vækerø) har ARM+RIB "
                "— køyrer separat analyse per område"
            )
            print(f"\n{'─' * 50}\n  LYA — Lysaker ({len(df_ly):,} rader)\n{'─' * 50}")
            excel_ly = kjor_mengdeuttak(omega_path=omega_path, df_override=df_ly,
                                        fob_mengde_mode=fob_mengde_mode)
            excel_ly = _rename_med_suffix(excel_ly, "LYA")
            print(f"[OK] {excel_ly}")

            print(f"\n{'─' * 50}\n  VEK — Vækerø ({len(df_vk):,} rader)\n{'─' * 50}")
            excel_vk = kjor_mengdeuttak(omega_path=omega_path, df_override=df_vk,
                                        fob_mengde_mode=fob_mengde_mode)
            excel_vk = _rename_med_suffix(excel_vk, "VEK")
            print(f"[OK] {excel_vk}")

            return [excel_ly, excel_vk]

        print(
            f"[INFO] Lysaker har ARM+RIB: {ly_ok}, Vækerø har ARM+RIB: {vk_ok} "
            "— køyrer kombinert analyse"
        )
    elif har_ly:
        print("[INFO] Berre Lysaker-data detektert")
    elif har_vk:
        print("[INFO] Berre Vækerø-data detektert")
    else:
        print("[INFO] Ingen kjent område-prefiks funne — køyrer kombinert analyse")

    excel = kjor_mengdeuttak(csv_path=csv_path, omega_path=omega_path,
                             fob_mengde_mode=fob_mengde_mode)
    return [excel]


# =============================================================================
# 1. Les ISY Linker
# =============================================================================

_ISY_CACHE: dict = {}


def les_isy(isy_path: str | list[str]) -> dict[str, dict]:
    """
    Mapping: post_nr (str) → {beskrivelse, enhetspris, mengde}.
    'mengde' brukast for samanlikning med Omega-kontraktsmengde.
    Aksepterer éin fil eller ei liste — seinare filer overstyrer tidlegare
    ved same postnummer. Cachar resultatet på (path, mtime).
    """
    stiar = [isy_path] if isinstance(isy_path, str) else list(isy_path)

    # Cache-nøkkel: (sti, mtime) for kvar fil — invalideres om fila endrast
    try:
        cache_key = tuple((s, os.path.getmtime(s)) for s in stiar if os.path.exists(s))
        if cache_key and cache_key in _ISY_CACHE:
            kakra = _ISY_CACHE[cache_key]
            print(f"   ISY Linker: (cached, {len(kakra)} postar)")
            return kakra
    except OSError:
        cache_key = None

    def _rens(s) -> str:
        return (
            str(s)
            .replace("_x000D_", "")
            .replace("\r", "")
            .replace("\n\n", "\n")
            .strip()[:600]
        )

    # Kolonne-aliaser: (post_nr_kol, beskrivelse_kol)
    _POST_KOLS  = ["Post nummer", "ProsessID: prosess, sted og element", "Prosess"]
    _BESK_KOLS  = ["All tekst", "Tekst - alt. (uten linjeskift)"]
    _STED_KODER = ("1302", "1303", "1407")

    isy_map: dict[str, dict] = {}
    for sti in stiar:
        print(f"   ISY Linker: {os.path.basename(sti)}")
        df = pd.read_excel(sti, engine="openpyxl")
        cols = list(df.columns)

        post_kol = next((k for k in _POST_KOLS if k in cols), None)
        besk_kol = next((k for k in _BESK_KOLS if k in cols), None)
        if post_kol is None:
            raise ValueError(
                f"Finner ingen post-kolonne i {os.path.basename(sti)}. "
                f"Kolonner: {cols}"
            )
        print(f"     -> post-kolonne: '{post_kol}', beskrivelse: '{besk_kol}'")

        df[post_kol] = df[post_kol].astype(str).str.strip()
        for _, row in df.iterrows():
            pnr = str(row.get(post_kol, "")).strip()
            if not pnr or pnr == "nan":
                continue
            isy_map[pnr] = {
                "beskrivelse": _rens(row.get(besk_kol, "") if besk_kol else ""),
                "enhetspris":  row.get("Pris: Anbud", None),
                "mengde":      pd.to_numeric(row.get("Mengde"), errors="coerce"),
            }

    # ── Berik sted-spesifikke poster med basetekst frå 00.xxx ─────────────────
    # Mønster: 00.81.1213  ←→  1302.81.1213 / 1303.81.1213 / 1407.81.1213
    base_map = {
        pnr: entry
        for pnr, entry in isy_map.items()
        if pnr.startswith("00.")
    }
    for base_pnr, base_entry in base_map.items():
        suffix = base_pnr[3:]          # strip leading "00."
        base_besk = base_entry["beskrivelse"]
        for sted in _STED_KODER:
            sted_pnr = f"{sted}.{suffix}"
            if sted_pnr in isy_map:
                # Sted-post finnes — berik beskrivelse viss base har meir tekst
                existing = isy_map[sted_pnr]
                if len(base_besk) > len(existing["beskrivelse"]):
                    existing["beskrivelse"] = base_besk
            else:
                # Sted-post finnes ikke — registrer som fallback frå base
                isy_map[sted_pnr] = {**base_entry}

    n_fallback = sum(
        1 for pnr in isy_map
        if pnr[:4] in {s + "." for s in _STED_KODER} or
           any(pnr.startswith(s + ".") for s in _STED_KODER)
    )
    print(f"     -> {len(base_map)} base-poster (00.xxx), "
          f"{n_fallback} sted-spesifikke oppslag registrert")
    if cache_key:
        _ISY_CACHE[cache_key] = isy_map
    return isy_map


# =============================================================================
# 2. ARM- og RIB-mengder frå Sammendrag
# =============================================================================

def les_arm_rib_fra_sammendrag(arm_rib_excel: str) -> tuple[dict, dict]:
    """
    Les Sammendrag-arket (header på rad 2) og returner:
        arm_dict : {OMEGA-post Armering kode → total armeringsvekt (kg)}
        bet_dict : {OMEGA-post Betong kode   → total betongvolum   (m³)}
    """
    df = pd.read_excel(
        arm_rib_excel, sheet_name="Sammendrag", engine="openpyxl", header=1
    )

    def _grp(post_kol, verdi_kol):
        # verdi_kol kan vere ein streng eller fleire alternative kolonnenamn
        kandidatar = (verdi_kol,) if isinstance(verdi_kol, str) else tuple(verdi_kol)
        verdi_kol = next((k for k in kandidatar if k in df.columns), None)
        if post_kol not in df.columns or verdi_kol is None:
            print(f"   [ADVARSEL] Fann ingen av kolonnene {kandidatar} i Sammendrag "
                  f"— mengdene vert 0. Kolonner: {list(df.columns)}")
            return {}
        grp = df.dropna(subset=[post_kol]).groupby(post_kol)[verdi_kol].sum()
        return {
            str(k).strip(): float(v)
            for k, v in grp.items()
            if str(k).strip() not in ("nan", "")
        }

    arm_dict = _grp(SAM_ARM_POST, SAM_ARM_VEKT)
    bet_dict = _grp(SAM_BET_POST, SAM_VOL)
    print(f"   ARM-koder med mengde: {len(arm_dict):,}")
    print(f"   Betong-koder med mengde: {len(bet_dict):,}")
    return arm_dict, bet_dict


# =============================================================================
# 3. Andre fag frå CSV + disiplin-mapping
# =============================================================================

def _tal_serie(s) -> "pd.Series":
    """Første talet i kvar celle, tolerant for einingssuffiks og ';'-lister.

        "42.668m²"          -> 42.668
        "12,793;1382,711*"  -> 12.793
        "1382,711*"         -> 1382.711

    FOB_Mengde brukar komma som desimalteikn, PGF-kolonnane punktum, og
    fleire av dei har eininga hengande på talet. Begge delar vert handtert.
    """
    tekst = s.astype(str).str.split(";").str[0].str.strip()
    tal = tekst.str.extract(r"^(-?\d+(?:[.,]\d+)?)", expand=False)
    return pd.to_numeric(tal.str.replace(",", ".", regex=False), errors="coerce")


def _finn_ark_areal_kol(df) -> str | None:
    """Finn ARK_AREAL-kolonnen uansett gruppeprefiks.

    Kolonnenamna er normaliserte frå "/" til "." lenger oppe, så vi ser på
    siste ledd: "ARK_Info.ARK_AREAL" og "PGF_ARK_Info.ARK_AREAL" gjev begge
    treff.
    """
    for kol in df.columns:
        if str(kol).replace("/", ".").split(".")[-1].strip().lower() == "ark_areal":
            return kol
    return None


def les_andre_fag(
    csv_path: str | list[str] | None = None,
    df: pd.DataFrame | None = None,
):
    """
    Les CSV, filtrer bort ARM og RIB, og returner:
        other_dict    : {mengdelistepost kode → sum FOB_Mengde}
        disiplin_dict : {mengdelistepost kode → fagkode (3 bokstavar)}

    Fagkoden vert henta frå Path-kolonnen i CSV via _fag_fra_sti().
    Semicolon-separerte koder vert behandla kvar for seg.
    """
    from csv_til_armrib import les_csv, ARM_PATH_TOKEN, RIB_PATH_TOKEN

    if df is None:
        df = les_csv(csv_path, lp_only=True)
    path_s = df["Path"].astype(str)
    arm_mask   = path_s.str.contains(ARM_PATH_TOKEN, regex=False, na=False)
    rib_mask   = path_s.str.contains(RIB_PATH_TOKEN, regex=False, na=False) & ~arm_mask
    df_other   = df[~arm_mask & ~rib_mask].copy()

    print(f"   Andre fag: {len(df_other):,} rader")

    if FOB_MENG_KOL not in df_other.columns or FOB_MENG not in df_other.columns:
        return {}, {}

    kolonner = [FOB_MENG_KOL, FOB_MENG, "Path"]
    if _ENHET_KOL in df_other.columns:
        kolonner.append(_ENHET_KOL)
    _ark_kol = _finn_ark_areal_kol(df_other)
    if _ark_kol:
        kolonner.append(_ark_kol)
    sub = df_other[kolonner].copy()
    if _ENHET_KOL not in sub.columns:
        sub[_ENHET_KOL] = ""
    # Første talet i lista — same ledd som "kode" og "_enhet" under. Før
    # vart heile strengen parsa, så element med fleire mengder
    # ("12,793;1382,711;12,793") gav NaN og vart talde som 0.
    sub["verdi"] = _tal_serie(sub[FOB_MENG]).fillna(0.0)
    sub["_fag"] = sub["Path"].astype(str).apply(_fag_fra_sti)
    sub[FOB_MENG_KOL] = sub[FOB_MENG_KOL].astype(str).str.strip()
    sub = sub[
        sub[FOB_MENG_KOL].notna()
        & (sub[FOB_MENG_KOL] != "")
        & (sub[FOB_MENG_KOL] != "nan")
    ]
    # Ta første kode frå semicolon-separert liste
    sub["kode"] = sub[FOB_MENG_KOL].str.split(";").str[0].str.strip()
    sub = sub[sub["kode"] != ""]

    # Summer per eining i tillegg til totalen. Same post kan ha element målt
    # i ulike einingar (t.d. både Stk og m²) — då skal berre den eininga
    # posten faktisk gjerast opp i telje med.
    from prognose_ark import normaliser_enhet as _n_enhet
    sub["_enhet"] = sub.get(_ENHET_KOL, "").astype(str).str.split(";").str[0].map(_n_enhet)

    # ARK_AREAL er det målte arealet i ARK-modellane, og er meir påliteleg
    # enn FOB_Mengde. Men det er eit AREAL: det kan berre brukast der posten
    # faktisk gjerast opp i m². Er eininga stk eller m, står FOB_Mengde.
    if _ark_kol:
        ark = _tal_serie(sub[_ark_kol])
        bruk = ark.notna() & (ark > 0) & (sub["_enhet"] == "m2")
        n_bruk = int(bruk.sum())
        if n_bruk:
            endra = int((bruk & (sub["verdi"] != ark)).sum())
            sub.loc[bruk, "verdi"] = ark[bruk]
            print(f"   ARK_AREAL brukt i staden for FOB_Mengde: "
                  f"{n_bruk:,} element ({endra:,} med ulik verdi)")
        n_hoppa = int((ark.notna() & (ark > 0) & (sub["_enhet"] != "m2")).sum())
        if n_hoppa:
            print(f"   ARK_AREAL fanst, men eininga er ikkje m² for "
                  f"{n_hoppa:,} element — FOB_Mengde står")
    else:
        print("   Ingen ARK_AREAL-kolonne i datagrunnlaget")

    per_enhet: dict = {}
    for (kode, enh), verdi in sub.groupby(["kode", "_enhet"])["verdi"].sum().items():
        per_enhet.setdefault(str(kode), {})[str(enh)] = float(verdi)

    other_dict   = sub.groupby("kode")["verdi"].sum().to_dict()
    disiplin_raw = (
        sub.groupby("kode")["_fag"]
        .agg(lambda x: x.mode().iloc[0] if len(x) > 0 else "")
    )
    disiplin_dict = disiplin_raw.to_dict()

    n_blanda = sum(1 for v in per_enhet.values() if len([e for e in v if e]) > 1)
    print(f"   Unike mengdelistepostar (andre fag): {len(other_dict):,} "
          f"({n_blanda} med fleire einingar)")
    return (
        {str(k): float(v) for k, v in other_dict.items()},
        {str(k): str(v)   for k, v in disiplin_dict.items()},
        per_enhet,
    )


# =============================================================================
# 3b. Rein-stål-postar på RIB-element (.07.-postar utan betong-par)
# =============================================================================

def les_staal_fob_fra_rib(
    csv_path: str | list[str] | None = None,
    df: pd.DataFrame | None = None,
) -> tuple[dict[str, float], dict[str, str]]:
    """Les stål-mengder (kg) frå FOB_Mengde for RIB-element som IKKJE er ARM.

    Mengdeuttak_Armering filtrerer ut rein-stål-element utan betong-par
    (sjå filteret `if not groups_med_betong: continue`), så desse manglar
    i Sammendrag-arket. Denne funksjonen plukkar dei opp direkte frå CSV
    slik at dei kjem med i hovuddokumentet.

    Bladkode-logikk: FOB kan duplisere mengda på fleire hierarki-nivå,
    t.d. "1303.01.07.223.7;1303.01.07.223.7.1" = "9,792;9,792" (same kg).
    Vi behald berre den mest spesifikke (bladnoden) for å unngå dobbel-
    telling, men legg samstundes verdien òg under foreldrekoden slik at
    Omega-oppslag fungerer uansett kva nivå Omega-koden ligg på.

    Returnerer:
        staal_dict    : {post-kode → sum kg}
        disiplin_dict : {post-kode → "RIB"}
    """
    from csv_til_armrib import les_csv, ARM_PATH_TOKEN, RIB_PATH_TOKEN

    if df is None:
        df = les_csv(csv_path, lp_only=True)

    if FOB_MENG_KOL not in df.columns or FOB_MENG not in df.columns:
        return {}, {}

    path_s = df["Path"].astype(str)
    arm_mask = path_s.str.contains(ARM_PATH_TOKEN, regex=False, na=False)
    rib_mask = path_s.str.contains(RIB_PATH_TOKEN, regex=False, na=False) & ~arm_mask
    df_rib   = df[rib_mask].copy()

    print(f"   RIB-element (ikkje ARM): {len(df_rib):,} rader")

    enhet_kol = "FOB_Info.FOB_Mengdeenhet"

    staal_sum: dict[str, float] = {}
    n_rader_med_staal = 0

    # iterrows() byggjer eit Series-objekt per rad og brukte ~8 s på dette
    # datasettet. Vi hentar dei tre kolonnane som rå arrayar i staden — same
    # logikk, ein brøkdel av tida.
    _koder_arr = df_rib[FOB_MENG_KOL].fillna("").astype(str).to_numpy()
    _meng_arr  = df_rib[FOB_MENG].fillna("").astype(str).to_numpy()
    _enhet_arr = (df_rib[enhet_kol].fillna("").astype(str).to_numpy()
                  if enhet_kol in df_rib.columns
                  else np.full(len(df_rib), "", dtype=object))

    for kode_str, meng_str, enhet_str in zip(_koder_arr, _meng_arr, _enhet_arr):
        kode_str  = kode_str.strip()
        meng_str  = meng_str.strip().replace(",", ".")
        enhet_str = enhet_str.strip().lower()

        koder   = [k.strip() for k in kode_str.split(";")  if k.strip() and k.strip() != "nan"]
        mengder = [m.strip() for m in meng_str.split(";")  if m.strip() and m.strip() != "nan"]
        enheter = [e.strip() for e in enhet_str.split(";") if e.strip()]

        if not koder or not mengder:
            continue

        # Bladkode: kode som IKKJE er prefiks av nokon annan kode i lista
        def _er_blad(idx: int) -> bool:
            k = koder[idx]
            pref = k + "."
            for j, other in enumerate(koder):
                if j != idx and other.startswith(pref):
                    return False
            return True

        # Finn éin representativ verdi: bladnoden si mengde
        # (i dei observerte data er parent;child same verdi, så det held)
        blad_indeksar = [i for i in range(len(koder)) if _er_blad(i)]
        if not blad_indeksar:
            continue

        # Berre kg-postar (stål-mengder). Andre einingar (stk, lm, m³ osv.)
        # vert handtert av andre dict-ar (other_dict, bet_dict).
        rad_telt = False
        for i in blad_indeksar:
            e = enheter[i] if i < len(enheter) else (enheter[0] if enheter else "")
            if e != "kg":
                continue
            v = pd.to_numeric(mengder[i], errors="coerce")
            if pd.isna(v) or v == 0:
                continue
            k = koder[i]
            # Legg verdien både på bladkoden og foreldrekoden(ane) som finst
            # i same liste — slik at Omega-oppslag treffer uansett nivå.
            staal_sum[k] = staal_sum.get(k, 0.0) + float(v)
            for j, other in enumerate(koder):
                if j != i and k.startswith(other + "."):
                    staal_sum[other] = staal_sum.get(other, 0.0) + float(v)
            rad_telt = True
        if rad_telt:
            n_rader_med_staal += 1

    disiplin_dict = {k: "RIB" for k in staal_sum}
    print(f"   Stål-postar på RIB (utan betong-par): {len(staal_sum):,} kodar "
          f"frå {n_rader_med_staal:,} rader")
    return staal_sum, disiplin_dict


# =============================================================================
# 3c. Einingar (frå CSV) for bladkodar — brukast i BIM_utan_Omega-arket
# =============================================================================

def les_enheter_fra_csv(
    csv_path: str | list[str] | None = None,
    df: pd.DataFrame | None = None,
) -> dict[str, str]:
    """Returner {post-kode → FOB_Mengdeenhet} for kodar i CSV.

    Brukast for å vise rett eining i BIM_utan_Omega-arket der vi ikkje har
    Omega-eining (sidan koden ikkje finst i Omega).
    """
    from csv_til_armrib import les_csv

    if df is None:
        df = les_csv(csv_path, lp_only=True)
    enhet_kol = "FOB_Info.FOB_Mengdeenhet"
    if FOB_MENG_KOL not in df.columns or enhet_kol not in df.columns:
        return {}

    result: dict[str, str] = {}

    # Same grep som i les_staal_fob_fra_rib: rå kolonnar i staden for
    # iterrows(), som brukte ~21 s på 833 000 rader.
    _koder_arr = df[FOB_MENG_KOL].fillna("").astype(str).to_numpy()
    _enhet_arr = df[enhet_kol].fillna("").astype(str).to_numpy()

    for kode_str, enhet_str in zip(_koder_arr, _enhet_arr):
        kode_str  = kode_str.strip()
        enhet_str = enhet_str.strip()
        koder   = [k.strip() for k in kode_str.split(";")  if k.strip() and k.strip() != "nan"]
        enheter = [e.strip() for e in enhet_str.split(";") if e.strip()]
        if not koder:
            continue
        for i, k in enumerate(koder):
            if k in result:
                continue
            e = enheter[i] if i < len(enheter) else (enheter[0] if enheter else "")
            if e:
                result[k] = e
    return result


# =============================================================================
# 3d. APM-mengder frå Maalebrev_samlet (Post sammenstilling)
# =============================================================================

def _les_rettede_apm_poster(apm_excel: str) -> None:
    """Meld APM-postar som vart retta då PDF-ane vart lesne, inn i feillista.

    Rettingane ligg i arket "Rettede poster" i Maalebrev_samlet.xlsx, så dei
    kjem med sjølv når APM-steget er hoppa over. Eldre filer utan arket gjev
    berre ei melding.
    """
    try:
        from pdf_maalebrev_til_excel import ARK_RETTEDE, registrer_rettede
        try:
            df = pd.read_excel(apm_excel, sheet_name=ARK_RETTEDE, engine="openpyxl")
        except ValueError:
            print(f"   [INFO] '{ARK_RETTEDE}' finst ikkje i APM-fila — køyr APM-steget "
                  f"på nytt for å få automatisk retta postar med i feillista")
            return
        registrer_rettede(df.fillna("").to_dict("records"))
        print(f"   APM: {len(df):,} retta/feil postar lagt i feillista")
    except Exception as exc:          # feillista skal aldri velte oversikta
        print(f"   [ADVARSEL] Feillista (APM): {exc}")


def les_apm_fra_maalebrev(apm_excel: str) -> dict[str, dict]:
    """Les 'Post sammenstilling'-arket frå Maalebrev_samlet.xlsx.

    Returnerer {post_kode: {"apm_mengde": float|None, "apm_kilde": str}}
    der apm_kilde er semikolonseparert liste av unike APM-nummer,
    t.d. "APM146; APM149".
    """
    if not apm_excel:
        print("   Ingen APM-fil — APM-mengder vert ikkje lesne")
        return {}
    if not os.path.exists(apm_excel):
        print(f"   [ADVARSEL] APM-fil ikkje funne: {apm_excel}")
        return {}

    _les_rettede_apm_poster(apm_excel)

    try:
        df = pd.read_excel(apm_excel, sheet_name="Post sammenstilling", engine="openpyxl")
    except Exception as e:
        print(f"   [ADVARSEL] Kunne ikkje lese APM-fil: {e}")
        return {}

    if "Post" not in df.columns or "Mengde fra APM" not in df.columns:
        print("   [ADVARSEL] Manglande kolonnar i 'Post sammenstilling'")
        return {}

    df["Post"] = (
        df["Post"].astype(str).str.strip()
        .str.replace("-", ".", regex=False)   # normaliser skiljetegn
    )
    apm_kol = "APM" if "APM" in df.columns else None

    # Hopp over [TOTAL]-aggregat-rader — summer berre individuelle APM-rader
    if apm_kol:
        df = df[df[apm_kol].astype(str).str.strip() != "[TOTAL]"]
    df = df[~df["Post"].isin(["", "nan"])].copy()
    df["_mengde"] = pd.to_numeric(df["Mengde fra APM"], errors="coerce").fillna(0.0)

    result: dict[str, dict] = {}
    for post, grp in df.groupby("Post"):
        total = float(grp["_mengde"].sum())
        if apm_kol:
            namn_liste = sorted({
                str(v).strip()
                for v in grp[apm_kol].dropna()
                if str(v).strip() not in ("", "nan", "[TOTAL]")
            })
            kort = []
            for n in namn_liste:
                m = _APM_NR_RE.search(n)
                kort.append(f"APM{m.group(1)}" if m else n[:20])
            apm_str = "; ".join(sorted(set(kort)))
        else:
            apm_str = ""
        result[str(post)] = {
            "apm_mengde": total if total != 0.0 else None,
            "apm_kilde":  apm_str,
        }

    print(f"   APM-postar frå Maalebrev_samlet: {len(result):,}")
    return result


# =============================================================================
# 4. Vel riktig mengde basert på Omega-enhet
# =============================================================================

def vel_modellenhet(
    kode: str,
    enhet: str,
    other_per_enhet: dict | None,
    enhet_csv: dict | None = None,
) -> str:
    """Eininga modellmengda faktisk vart henta i — same val som vel_mengde().

    Blir brukt til enhetskontrollen i prognosen, så flagget følgjer den
    mengda som faktisk er brukt og ikkje berre den første eininga i fila.
    """
    from prognose_ark import normaliser_enhet as _n_enhet
    per = (other_per_enhet or {}).get(kode)
    if per:
        reelle = {k: v for k, v in per.items() if k}
        mål = _n_enhet(enhet)
        if mål and mål in reelle:
            return mål
        if len(reelle) == 1:
            return next(iter(reelle))
        if reelle:
            return "; ".join(sorted(reelle))
    return (enhet_csv or {}).get(kode, "")


def vel_mengde(
    kode: str,
    enhet: str,
    arm_dict: dict,
    bet_dict: dict,
    other_dict: dict,
    other_per_enhet: dict | None = None,
) -> float | None:
    e = enhet.lower().strip()
    if e == "kg":
        return arm_dict.get(kode)
    if e in ("m3", "m³"):
        return bet_dict.get(kode)

    # Andre einingar (rs, stk, lm, m, m2 …). Har posten element målt i fleire
    # einingar, brukar vi berre dei som er målt i same eining som posten
    # gjerast opp i — elles vert t.d. eit areal talt som stykktal.
    if other_per_enhet:
        from prognose_ark import normaliser_enhet as _n_enhet
        per = other_per_enhet.get(kode)
        if per:
            mål = _n_enhet(enhet)
            if mål and mål in per:
                return per[mål]
            reelle = {k: v for k, v in per.items() if k}
            if mål and len(reelle) > 1:
                # Fleire einingar, ingen av dei er posten sin → ikkje
                # samanliknbart. La enhetskontrollen i prognosen fange det.
                return other_dict.get(kode)

    v = other_dict.get(kode)
    if v is not None:
        return v
    return arm_dict.get(kode) or bet_dict.get(kode)


# =============================================================================
# 5a. Betong / forskaling konsistenssjekk
# =============================================================================

# Regex for å hente tykkelse (mm) frå gruppeoverskrift som "Veggfundament 250 B35"
_TYKKELSE_RE = re.compile(r'\b(\d{2,4})\s*(?:mm|$|\s+B\d)', re.IGNORECASE)

# Konstruksjonstypar der volum/tykkelse = forskalingsareal (éi side)
_ENKELTSIDE_TYPER = {"vegg", "yttervegg", "innervegg", "skjermvegg", "parapet", "plate", "dekke"}
# Typar vi hoppar over (geometri passer ikkje formelen)
_SKIP_TYPER       = {"søyle", "soyle", "bjelke", "ramme"}


def _parse_tykkelse_mm(tittel: str) -> float | None:
    """Hent tykkelse i mm frå gruppeoverskrift. Returnerer mm som float eller None."""
    m = _TYKKELSE_RE.search(tittel)
    return float(m.group(1)) if m else None


def _konstruksjonstype(tittel: str) -> str:
    """Hent første ord i tittelen som konstruksjonstype."""
    return tittel.split()[0].lower() if tittel.strip() else ""


def _forskaling_faktor(ktype: str) -> float:
    """Antal sider med forskaling: vegg/parapet = 2, plate/dekke = 1."""
    if any(s in ktype for s in ("vegg", "parapet")):
        return 2.0
    return 1.0


def bygg_betong_forskaling_ark(
    df_omega: pd.DataFrame,
    isy_map: dict[str, dict],
) -> list[dict]:
    """
    For kvar Normalbetong-post (m³):
      - finn gruppehovud (strip siste .N) og tykkelse frå ISY-tittelen
      - finn tilhøyrande forskalingspost (m² + 'FORSKALING' i ISY-tittel)
      - rekn ut forventa rapportert forskaling = rapp_betong/tykkelse × sider
      - samanlikn med rapportert forskalingsmengde i Omega
    """
    omega_idx = {str(r[OM_KODE]).strip(): r for _, r in df_omega.iterrows()}

    resultat: list[dict] = []
    sette: set = set()

    for _, row in df_omega.iterrows():
        kode  = str(row[OM_KODE]).strip()
        tittel = str(row.get(OM_TITTEL, "") or "").upper()
        enhet  = str(row.get(OM_ENHET,  "") or "").lower().strip()

        if "NORMALBETONG" not in tittel or enhet not in ("m3", "m³"):
            continue
        if kode in sette:
            continue
        sette.add(kode)

        parts = kode.rsplit(".", 1)
        if len(parts) < 2:
            continue
        gruppe_pnr = parts[0]

        gruppe_tittel = isy_map.get(gruppe_pnr, {}).get("beskrivelse", "").split("\n")[0].strip()
        ktype = _konstruksjonstype(gruppe_tittel)
        if any(s in ktype for s in _SKIP_TYPER):
            continue

        tykkelse_mm = _parse_tykkelse_mm(gruppe_tittel)
        if not tykkelse_mm:
            continue
        tykkelse_m = tykkelse_mm / 1000.0
        fk_faktor  = _forskaling_faktor(ktype)

        volum_omega = pd.to_numeric(row.get(OM_KONTR),   errors="coerce")
        rapp_betong = pd.to_numeric(row.get(OM_RAPPORT), errors="coerce")

        # Finn forskalingspost: same gruppe, m², 'FORSKALING' i ISY-tittel
        forskaling_kode = ""
        fk_tittel = ""
        fk_rapp = None
        gruppe_prefix = gruppe_pnr + "."
        for k2, r2 in omega_idx.items():
            if not k2.startswith(gruppe_prefix):
                continue
            e2 = str(r2.get(OM_ENHET, "") or "").lower().strip()
            if e2 not in ("m2", "m²"):
                continue
            isy_tittel2 = isy_map.get(k2, {}).get("beskrivelse", "").split("\n")[0].upper()
            if "FORSKALING" in isy_tittel2:
                forskaling_kode = k2
                fk_tittel = str(r2.get(OM_TITTEL, "") or "")
                fk_rapp = pd.to_numeric(r2.get(OM_RAPPORT), errors="coerce")
                break

        forventa_rapp_forsk = (round(float(rapp_betong) / tykkelse_m * fk_faktor, 2)
                               if pd.notna(rapp_betong) else None)

        if (forventa_rapp_forsk and fk_rapp is not None and pd.notna(fk_rapp)
                and forventa_rapp_forsk != 0):
            avvik_rapp     = round(float(fk_rapp) - float(forventa_rapp_forsk), 2)
            avvik_rapp_pst = round(float(fk_rapp) / float(forventa_rapp_forsk) * 100 - 100, 1)
        else:
            avvik_rapp     = None
            avvik_rapp_pst = None

        resultat.append({
            "gruppe_pnr":              gruppe_pnr,
            "gruppe_tittel":           gruppe_tittel,
            "tykkelse_mm":             tykkelse_mm,
            "betong_kode":             kode,
            "volum_omega":             round(float(volum_omega), 3) if pd.notna(volum_omega) else None,
            "rapp_betong":             round(float(rapp_betong), 3) if pd.notna(rapp_betong) else None,
            "forventa_rapp_forskaling": forventa_rapp_forsk,
            "forskaling_kode":         forskaling_kode,
            "forskaling_tittel":       fk_tittel,
            "forskaling_rapp":         round(float(fk_rapp), 2) if fk_rapp is not None and pd.notna(fk_rapp) else None,
            "avvik_rapp":              avvik_rapp,
            "avvik_rapp_pst":          avvik_rapp_pst,
        })

    resultat.sort(key=lambda r: r["betong_kode"])
    return resultat


KOLONNER_RAPP: list[tuple[str, int, bool]] = [
    ("Gruppe (prefiks)",              22, False),
    ("Beskrivelse",                   30, False),
    ("Tykkelse (mm)",                 14, True),
    ("Betong-post",                   22, False),
    ("Rapportert betong (m³)",        20, True),
    ("Forv. forskal. rapp. (m²)",     22, True),
    ("Forskaling-post",               22, False),
    ("Forskaling tittel",             30, False),
    ("Rapportert forskaling (m²)",    22, True),
    ("Avvik rapp. (m²)",              16, True),
    ("Avvik rapp. (%)",               14, True),
]

KOLONNER_MATERIAL: list[tuple[str, int, bool]] = [
    ("Fagkapittel",            34, False),
    ("Mengdelistepost",        24, False),
    ("Dominerande material",   22, False),
    ("Avvikande material",     22, False),
    ("Antall element",         14, True),
    ("Usikker",                10, False),
    ("Handtering",             26, False),
]


# =============================================================================
# 5. ISY vs Omega avvik
# =============================================================================

def finn_isy_omega_avvik(
    df_omega: pd.DataFrame,
    isy_map: dict[str, dict],
) -> list[dict]:
    """
    Returner liste av postar der ISY-mengde og Omega-kontraktsmengde
    avviker med meir enn _AVVIK_TOL (relativt).
    """
    avvik: list[dict] = []
    for _, row in df_omega.iterrows():
        kode = str(row[OM_KODE]).strip()
        isy  = isy_map.get(kode, {})
        isy_mengde = isy.get("mengde")
        if pd.isna(isy_mengde) or isy_mengde is None:
            continue
        omega_kontr = pd.to_numeric(row.get(OM_KONTR), errors="coerce")
        if pd.isna(omega_kontr):
            continue

        isy_v   = float(isy_mengde)
        omega_v = float(omega_kontr)

        # Relativ differanse (unngå divisjon på null)
        ref = max(abs(isy_v), abs(omega_v), 1e-9)
        if abs(isy_v - omega_v) / ref > _AVVIK_TOL:
            avvik.append({
                "kode":           kode,
                "tittel":         str(row.get(OM_TITTEL, "") or ""),
                "enhet":          str(row.get(OM_ENHET,  "") or ""),
                "isy_mengde":     isy_v,
                "omega_kontrakt": omega_v,
                "differanse":     isy_v - omega_v,
            })
    return avvik


# =============================================================================
# 6. Excel-skriving
# =============================================================================

_HDR_FILL   = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT   = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
_ALT_FILL   = PatternFill("solid", fgColor="DCE6F1")
_MISS_FILL  = PatternFill("solid", fgColor="FCE4D6")
_AVVIK_FILL = PatternFill("solid", fgColor="FFF2CC")
_WRAP       = Alignment(wrap_text=True, vertical="top")
_TOP        = Alignment(vertical="top")
_CAL10      = Font(name="Calibri", size=10)

# Fargeskala for %-kolonnar (% av rapportert / % av kontrakt)
# ≤ 100 % → grøn, > 100 % → raud
_GRNN_LETT  = PatternFill("solid", fgColor="C6EFCE")
_RAUD_LETT  = PatternFill("solid", fgColor="FFC7CE")
_LINK_FONT  = Font(name="Calibri", size=10, color="0563C1", underline="single")

# Kolonnar: (header, bredde, er-tal-kolonne)
# I_modell:       mengde_modell / rapportert  → "% av rapportert"
# Ikke_i_modell:  rapportert / kontrakt       → "% av kontrakt"
KOLONNER_I: list[tuple[str, int, bool]] = [
    ("Mengdelistepost",           18, False),
    ("Fag",                        6, False),
    ("Tittel (Omega)",            30, False),
    ("Enhet i modell",            13, False),
    ("Mengde fra modell",         16, True),
    ("APM-mengde",                14, True),
    ("% av rapportert",           12, True),
    ("Enhet",                      8, False),
    ("Enhetspris",                14, True),
    ("Kontraktsmengde",           16, True),
    ("Rapportert mengde",         16, True),
    ("APM-kilde",                 22, False),
    ("Beskrivelse (ISY Linker)",  65, False),
]

KOLONNER_NEI: list[tuple[str, int, bool]] = [
    ("Mengdelistepost",           18, False),
    ("Fag",                        6, False),
    ("Tittel (Omega)",            30, False),
    ("Enhet i modell",            13, False),
    ("Mengde fra modell",         16, True),
    ("APM-mengde",                14, True),
    ("% av kontrakt",             12, True),
    ("Enhet",                      8, False),
    ("Enhetspris",                14, True),
    ("Kontraktsmengde",           16, True),
    ("Rapportert mengde",         16, True),
    ("APM-kilde",                 22, False),
    ("Beskrivelse (ISY Linker)",  65, False),
]

KOLONNER = KOLONNER_I  # bakoverkompatibelt alias

AVVIK_KOLONNER: list[tuple[str, int, bool]] = [
    ("Mengdelistepost",     18, False),
    ("Tittel",              35, False),
    ("Enhet",                8, False),
    ("ISY Mengde",          16, True),
    ("Omega Kontraktsmengde", 20, True),
    ("Differanse",          16, True),
]

# BIM_utan_Omega: postar med mengde i BIM-modell, men ikkje i Omega-kontrakt
KOLONNER_BIM_UTAN_OMEGA: list[tuple[str, int, bool]] = [
    ("Mengdelistepost",          18, False),
    ("Fag",                       6, False),
    ("Mengde fra modell",        16, True),
    ("Enhet",                     8, False),
    ("Kilde",                    14, False),
]


def _skriv_ark(
    ws,
    tittel: str,
    data: list[dict],
    kolonner: list[tuple[str, int, bool]],
    tabel_stil: str = "TableStyleMedium2",
) -> None:
    ws.title = tittel[:31]

    for ci, (hdr, bredde, _) in enumerate(kolonner, 1):
        c = ws.cell(row=1, column=ci, value=hdr)
        c.font      = _HDR_FONT
        c.fill      = _HDR_FILL
        c.alignment = _TOP
        ws.column_dimensions[get_column_letter(ci)].width = bredde

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 18

    for ri, rad in enumerate(data, 2):
        vals = [rad.get(k) for k in _rad_nøklar(kolonner)]
        for ci, (val, (_, _, er_tal)) in enumerate(zip(vals, kolonner), 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font      = _CAL10
            c.alignment = _WRAP if (not er_tal and ci == len(kolonner)) else _TOP
            if er_tal and isinstance(val, (int, float)):
                hdr = kolonner[ci - 1][0]
                c.number_format = "0.0%" if hdr in ("% av kontrakt", "% av rapportert") else "#,##0.00"

    n_rows = len(data)
    n_cols = len(kolonner)
    if n_rows > 0:
        tbl_ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"
        tbl_name = "Tabell_" + ws.title.replace(" ", "_").replace("/", "_")
        tbl = Table(displayName=tbl_name, ref=tbl_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name=tabel_stil,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tbl)
    else:
        ws.auto_filter.ref = ws.dimensions


def _rad_nøklar(kolonner: list[tuple[str, int, bool]]) -> list[str]:
    """Konverter kolonneheadar til dict-nøklar."""
    mapping = {
        "Mengdelistepost":      "kode",
        "Fag":                  "fag",
        "Tittel (Omega)":       "tittel",
        "Mengde fra modell":    "mengde_modell",
        "Enhet i modell":       "modellenhet",
        "% av kontrakt":        "pst_av_kontrakt",
        "% av rapportert":      "pst_av_kontrakt",
        "Enhet":                "enhet",
        "Enhetspris":           "enhetspris",
        "Kontraktsmengde":      "kontraktsmengde",
        "Rapportert mengde":    "rapportert",
        "Beskrivelse (ISY Linker)": "beskrivelse",
        "Tittel":               "tittel",
        "ISY Mengde":           "isy_mengde",
        "Omega Kontraktsmengde": "omega_kontrakt",
        "Differanse":           "differanse",
        # Rapp_kontroll-ark
        "Gruppe (prefiks)":              "gruppe_pnr",
        "Beskrivelse":                   "gruppe_tittel",
        "Tykkelse (mm)":                 "tykkelse_mm",
        "Betong-post":                   "betong_kode",
        "Rapportert betong (m³)":        "rapp_betong",
        "Forv. forskal. rapp. (m²)":     "forventa_rapp_forskaling",
        "Forskaling-post":               "forskaling_kode",
        "Forskaling tittel":             "forskaling_tittel",
        "Rapportert forskaling (m²)":    "forskaling_rapp",
        "Avvik rapp. (m²)":              "avvik_rapp",
        "Avvik rapp. (%)":               "avvik_rapp_pst",
        # BIM_utan_Omega
        "Kilde":                "kilde",
        # APM-integrasjon
        "APM-mengde":           "apm_mengde",
        "APM-kilde":            "apm_kilde",
        # Material_kontroll
        "Fagkapittel":          "fagkapittel",
        "Dominerande material": "dom_material",
        "Avvikande material":   "feil_material",
        "Antall element":       "antall",
        "Usikker":              "usikker",
        "Handtering":           "handtering",
    }
    return [mapping.get(hdr, hdr) for hdr, _, _ in kolonner]


# =============================================================================
# 6b. Post-prosessering: fargelegging og APM-hyperlenker
# =============================================================================

def _fargelegg_pst_kolonne(ws, kolonner: list, hdr_namn: str) -> None:
    """Fargelegg %-kolonne: ≤ 100 % grøn, > 100 % raud."""
    col_idx = next((i for i, (h, _, _) in enumerate(kolonner, 1) if h == hdr_namn), None)
    if col_idx is None:
        return
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        v = row[col_idx - 1].value
        if not isinstance(v, (int, float)):
            continue
        row[col_idx - 1].fill = _GRNN_LETT if v <= 1.0 else _RAUD_LETT


def _legg_apm_hyperlenkjer(ws, kolonner: list, apm_link_target: str) -> None:
    """Set hyperlenke på APM-kilde-celler → opnar Maalebrev_samlet.xlsx på
    'Post sammenstilling'-arket. Brukar Windows-sti utan URL-encoding —
    det er det Excel forventar for lokale filer. Target kan vere absolutt
    (desktop) eller berre filnamn (relativ, web-modus).
    """
    col_idx = next((i for i, (h, _, _) in enumerate(kolonner, 1) if h == "APM-kilde"), None)
    if col_idx is None or not apm_link_target:
        return
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[col_idx - 1]
        if not (cell.value and str(cell.value).strip()):
            continue
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate,
            target=apm_link_target,
            location="'Post sammenstilling'!A1",
            display=str(cell.value),
        )
        cell.font = _LINK_FONT


# =============================================================================
# Hovudflyt
# =============================================================================

def lag_oversikt(
    csv_path:        str | list[str]  = CSV_PATH,
    omega_path:      str              = OMEGA_PATH,
    isy_path:        str | list[str]  = ISY_PATH,
    apm_excel:       str              = APM_EXCEL,
    apm_link_target: str | None       = None,     # override for hyperlenker (web → basename)
    arm_rib_excel:   str | None       = None,     # bakoverkompatibel enkelt-fil
    arm_rib_excels:  list[str] | None = None,     # ny: éin fil per område
    df_csv:          "pd.DataFrame | None" = None,  # ferdiglesne CSV-ar
    utdir:           str              = UTDIR,
) -> str:
    """
    Generer mengde-oversikt Excel med tre ark:
      I_modell, Ikke_i_modell, ISY_Omega_avvik.
    """

    # ── Nullstill material-avviksloggen for denne køyringa ───────────────────
    import csv_til_armrib as _cta
    _cta.nullstill_material_avvik()

    # CSV-ane vart tidlegare lesne på nytt i kvart steg — fire fulle
    # gjennomlesingar av dei same filene. No vert dei lesne éin gong og
    # sendt vidare.
    if df_csv is None:
        from csv_til_armrib import les_csv
        print("[0/6] Leser CSV-ane (éin gong for heile køyringa) …")
        df_csv = les_csv(csv_path, lp_only=True)

    # ── Normaliser arm_rib_excels ─────────────────────────────────────────────
    if arm_rib_excels is None and arm_rib_excel is not None:
        arm_rib_excels = [arm_rib_excel]

    # ── Steg 1: ARM/RIB-analyse ──────────────────────────────────────────────
    if arm_rib_excels is None:
        print("[1/6] Kjører ARM/RIB-analyse (Mengdeuttak_Armering.py) …")
        arm_rib_excels = kjor_arm_rib_per_omraade(
            csv_path=csv_path, omega_path=omega_path, df=df_csv)
        for p in arm_rib_excels:
            if not p or not os.path.exists(p):
                raise RuntimeError(f"ARM/RIB-analysen produserte ingen output-fil: {p}")
    else:
        print("[1/6] Bruker eksisterande ARM/RIB-Excel:")
        for p in arm_rib_excels:
            print(f"      {p}")

    # ── Steg 2: ARM/RIB-mengder frå Sammendrag (kombinert frå alle filer) ────
    print("[2/6] Leser ARM/RIB-mengder frå Sammendrag …")
    arm_dict: dict = {}
    bet_dict: dict = {}
    for excel in arm_rib_excels:
        a, b = les_arm_rib_fra_sammendrag(excel)
        arm_dict.update(a)
        bet_dict.update(b)

    # ── Steg 3: Andre fag + disiplin-mapping frå CSV ─────────────────────────
    print("[3/6] Ekstraher andre fag og fagkodar frå CSV …")
    other_dict, disiplin_dict, other_per_enhet = les_andre_fag(df=df_csv)

    # Legg til "RIB" for alle ARM- og Betong-koder
    for kode in arm_dict:
        disiplin_dict.setdefault(kode, "RIB")
    for kode in bet_dict:
        disiplin_dict.setdefault(kode, "RIB")

    # ── Steg 3b: Rein-stål-postar på RIB (utan betong-par) ───────────────────
    # Mengdeuttak_Armering filtrerer desse vekk (har ikkje betong-par), men
    # FOB_Mengde har gyldig stålvekt i kg. Supplér arm_dict for manglande kodar.
    print("[3b/6] Plukkar opp rein-stål-postar frå RIB-element direkte i CSV …")
    staal_rib_dict, staal_disiplin = les_staal_fob_fra_rib(df=df_csv)
    nye_kodar = 0
    for kode, mengde in staal_rib_dict.items():
        if kode not in arm_dict or not arm_dict[kode]:
            arm_dict[kode] = mengde
            disiplin_dict.setdefault(kode, staal_disiplin.get(kode, "RIB"))
            nye_kodar += 1
    print(f"   Nye stål-kodar lagt til arm_dict: {nye_kodar:,}")

    # ── Steg 3c: APM-mengder frå Maalebrev_samlet ───────────────────────────
    print("[3c/6] Leser APM-mengder frå Maalebrev_samlet …")
    apm_dict = les_apm_fra_maalebrev(apm_excel)

    # ── Steg 4: Omega + ISY Linker ───────────────────────────────────────────
    print("[4/6] Leser Omega og ISY Linker …")
    from base_lesning import les_omega
    df_omega = les_omega(omega_path)
    df_omega[OM_EPREIS] = pd.to_numeric(
        df_omega.get(OM_EPREIS, pd.Series(dtype=float)), errors="coerce"
    )
    isy_map = les_isy(isy_path)

    # ── Steg 5: ISY vs Omega avvik ───────────────────────────────────────────
    print("[5/6] Sjekkar ISY vs Omega mengder …")
    avvik_rader = finn_isy_omega_avvik(df_omega, isy_map)
    print(f"   Avvik funne: {len(avvik_rader):,} postar")

    # ── Steg 6: Bygg og skriv Excel ──────────────────────────────────────────
    print("[6/6] Bygger Excel-oversikt …")

    # Eininga modellen sjølv brukar. Trengst for å avdekke tilfelle der
    # modellen har målt t.d. areal medan posten skal gjerast opp per stk —
    # då er modellmengda ikkje samanliknbar med kontraktsmengda.
    enhet_csv = les_enheter_fra_csv(df=df_csv)

    rader_i:   list[dict] = []
    rader_nei: list[dict] = []

    for _, om_row in df_omega.iterrows():
        kode  = str(om_row[OM_KODE]).strip()
        if not kode or kode == "nan":
            continue

        enhet = str(om_row.get(OM_ENHET, "") or "")
        isy   = isy_map.get(kode, {})

        base = {
            "kode":            kode,
            "fag":             disiplin_dict.get(kode, ""),
            "tittel":          str(om_row.get(OM_TITTEL, "") or ""),
            "enhet":           enhet,
            "enhetspris":      (
                isy.get("enhetspris")
                or pd.to_numeric(om_row.get(OM_EPREIS), errors="coerce")
            ),
            "kontraktsmengde": pd.to_numeric(om_row.get(OM_KONTR),   errors="coerce"),
            "rapportert":      pd.to_numeric(om_row.get(OM_RAPPORT), errors="coerce"),
            "beskrivelse":     isy.get("beskrivelse", ""),
            "modellenhet":     vel_modellenhet(kode, enhet, other_per_enhet, enhet_csv),
        }

        # Prøv både original og normalisert (dash→dot) — Omega kan ha
        # "1407-10.43.482" medan APM-dict-nøklane er "1407.10.43.482".
        apm_info = apm_dict.get(kode) or apm_dict.get(kode.replace("-", ".")) or {}
        base["apm_mengde"] = apm_info.get("apm_mengde")
        base["apm_kilde"]  = apm_info.get("apm_kilde", "")

        mengde = vel_mengde(kode, enhet, arm_dict, bet_dict, other_dict,
                            other_per_enhet)
        if mengde is not None:
            rapp = base["rapportert"]
            pst = (float(rapp) / mengde) if (pd.notna(rapp) and mengde != 0) else None
            rader_i.append({**base, "mengde_modell": mengde, "pst_av_kontrakt": pst})
        else:
            kontr = base["kontraktsmengde"]
            rapp  = base["rapportert"]
            pst   = (float(rapp) / float(kontr)) if (pd.notna(rapp) and pd.notna(kontr) and float(kontr) != 0) else None
            rader_nei.append({**base, "mengde_modell": None, "pst_av_kontrakt": pst})

    # ── BIM_utan_Omega: kodar med mengde i BIM-modell som IKKJE finst i Omega ─
    # Mengdeuttak_Armering + andre fag kan ha kodar som ikkje er i kontrakta.
    # Tidlegare vart dei silently dropped — no listast dei her for synleg
    # avvik mellom modell og kontrakt.
    print("[6a/6] Byggjer BIM_utan_Omega …")
    omega_kodar_alle = set(df_omega[OM_KODE].astype(str).str.strip())

    bim_utan: list[dict] = []
    def _legg_til(kode: str, mengde, kilde: str, default_enhet: str):
        if not kode or kode == "nan" or kode in omega_kodar_alle:
            return
        if mengde is None or not isinstance(mengde, (int, float)) or mengde <= 0:
            return
        bim_utan.append({
            "kode":          kode,
            "fag":           disiplin_dict.get(kode, ""),
            "mengde_modell": float(mengde),
            "enhet":         enhet_csv.get(kode, default_enhet),
            "kilde":         kilde,
        })

    seen: set[str] = set()
    for kode, m in arm_dict.items():
        if kode in seen:
            continue
        seen.add(kode)
        _legg_til(kode, m, "ARM/stål", "kg")
    for kode, m in bet_dict.items():
        if kode in seen:
            continue
        seen.add(kode)
        _legg_til(kode, m, "Betong", "m³")
    for kode, m in other_dict.items():
        if kode in seen:
            continue
        seen.add(kode)
        _legg_til(kode, m, "Andre fag", "")

    # Sorter desc etter mengde for å vise største avvik øvst
    bim_utan.sort(key=lambda r: -(r.get("mengde_modell") or 0))
    print(f"   BIM-kodar utan Omega-match: {len(bim_utan):,}")

    # ── Rapp_kontroll: rapportert betong vs rapportert forskaling ────────────
    print("[6b/6] Bygger Rapp_kontroll …")
    betong_rader = bygg_betong_forskaling_ark(df_omega, isy_map)
    print(f"   Normalbetong-poster med tykkelse: {len(betong_rader):,}")

    # Rapp_kontroll: kun vegger der rapportert betong er innanfor ±10 % av kontraktsmengde
    rapp_rader = [
        r for r in betong_rader
        if (
            r.get("rapp_betong")     is not None
            and r.get("forskaling_rapp") is not None
            and r.get("volum_omega") is not None
            and r["volum_omega"] != 0
            and abs(r["rapp_betong"] / r["volum_omega"] - 1.0) <= 0.10
        )
    ]

    # ── Material_kontroll: element med avvikande material per post ───────────
    # Sikre avvik (framand materiale) er ekskludert frå mengda; usikre
    # (konstruksjonsmateriale mot konstruksjonsmateriale) er beheldt + flagga.
    material_rader = []
    for (fag, post, dom, feil), info in sorted(_cta.MATERIAL_AVVIK.items()):
        if isinstance(info, dict):
            antall  = info.get("antall", 0)
            usikker = info.get("usikker", False)
        else:  # bakoverkompat (gammal int-form)
            antall, usikker = info, False
        material_rader.append({
            "fagkapittel":   fag,
            "kode":          post,
            "dom_material":  dom,
            "feil_material": feil,
            "antall":        antall,
            "usikker":       "Ja" if usikker else "Nei",
            "handtering":    "Beheldt – sjekk manuelt" if usikker else "Ekskludert",
        })
    # Usikre øvst (treng tilsyn), deretter etter antall
    material_rader.sort(key=lambda r: (r["usikker"] != "Ja", -r["antall"]))
    n_eks = sum(1 for r in material_rader if r["usikker"] == "Nei")
    n_fl  = sum(1 for r in material_rader if r["usikker"] == "Ja")
    print(f"   Material-avvik: {n_eks:,} ekskludert, {n_fl:,} flagga (usikre)")

    wb    = openpyxl.Workbook()
    ws_i  = wb.active
    ws_ni = wb.create_sheet("Ikke_i_modell")
    ws_bo = wb.create_sheet("BIM_utan_Omega")
    ws_av = wb.create_sheet("ISY_Omega_avvik")
    ws_rp = wb.create_sheet("Rapp_kontroll")
    ws_mt = wb.create_sheet("Material_kontroll")

    _skriv_ark(ws_i,  "I_modell",          rader_i,      KOLONNER_I,              tabel_stil="TableStyleMedium2")
    _skriv_ark(ws_ni, "Ikke_i_modell",     rader_nei,    KOLONNER_NEI,            tabel_stil="TableStyleMedium3")
    _skriv_ark(ws_bo, "BIM_utan_Omega",    bim_utan,     KOLONNER_BIM_UTAN_OMEGA, tabel_stil="TableStyleMedium5")
    _skriv_ark(ws_av, "ISY_Omega_avvik",   avvik_rader,  AVVIK_KOLONNER,          tabel_stil="TableStyleMedium4")
    _skriv_ark(ws_rp, "Rapp_kontroll",     rapp_rader,   KOLONNER_RAPP,           tabel_stil="TableStyleMedium7")
    _skriv_ark(ws_mt, "Material_kontroll", material_rader, KOLONNER_MATERIAL,     tabel_stil="TableStyleMedium6")

    # Fargelegg "Usikker"-kolonna: gul = usikker (beheldt), raud = ekskludert
    _us_mt = next(i for i, (h, _, _) in enumerate(KOLONNER_MATERIAL, 1) if h == "Usikker")
    _gul_mt = PatternFill("solid", fgColor="FFEB9C")
    _raud_mt_fill = PatternFill("solid", fgColor="FFC7CE")
    for row in ws_mt.iter_rows(min_row=2, max_row=ws_mt.max_row):
        cell = row[_us_mt - 1]
        cell.fill = _gul_mt if str(cell.value) == "Ja" else _raud_mt_fill

    # Fargelegg %-kolonnar + set APM-hyperlenker
    _fargelegg_pst_kolonne(ws_i,  KOLONNER_I,   "% av rapportert")
    _fargelegg_pst_kolonne(ws_ni, KOLONNER_NEI, "% av kontrakt")
    link_target = apm_link_target if apm_link_target is not None else apm_excel
    if link_target and os.path.exists(apm_excel):
        _legg_apm_hyperlenkjer(ws_i,  KOLONNER_I,   link_target)
        _legg_apm_hyperlenkjer(ws_ni, KOLONNER_NEI, link_target)

    _GRNN = PatternFill("solid", fgColor="C6EFCE")
    _RAUD = PatternFill("solid", fgColor="FFC7CE")

    # Fargelegg avvik-kolonne i Rapp_kontroll (grense 3%)
    avvik_col_rp = next(i for i, (h, _, _) in enumerate(KOLONNER_RAPP, 1) if h == "Avvik rapp. (%)")
    for row in ws_rp.iter_rows(min_row=2, max_row=ws_rp.max_row):
        v = row[avvik_col_rp - 1].value
        if isinstance(v, (int, float)):
            row[avvik_col_rp - 1].fill = _GRNN if abs(v) <= 3 else _RAUD

    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(utdir, exist_ok=True)
    utfil = os.path.join(utdir, f"mengde_oversikt_{ts}.xlsx")
    from lagring import vent_til_skrivbar
    vent_til_skrivbar(utfil)
    wb.save(utfil)

    print(f"\n[OK] Lagret: {utfil}")
    print(f"     I modell:          {len(rader_i):,} Omega-postar")
    print(f"     Ikke i modell:     {len(rader_nei):,} Omega-postar")
    print(f"     BIM utan Omega:    {len(bim_utan):,} BIM-kodar (utan kontraktsmatch)")
    print(f"     ISY/Omega avvik:   {len(avvik_rader):,} postar")
    print(f"     Rapp_kontroll:     {len(rapp_rader):,} vegger (~100 % rapportert)")
    print(f"     Material_kontroll: {n_eks:,} ekskludert, {n_fl:,} flagga (usikre)")
    return utfil


if __name__ == "__main__":
    lag_oversikt()
