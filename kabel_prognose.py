# -*- coding: utf-8 -*-
"""
Kabelmengder til prognosen — frå kabellistene, ikkje frå BIM-modellen.

Kablane ligg ikkje som objekt i modellen, så prognosen for dei vert henta frå
kabellistene (PF-*-LC-*.xlsx, arket "Kabelliste") og kopla mot NS3420-postane
på kabeltypen slik han står i "All tekst".

Kabeltypen står i to former i NS3420, og begge vert lesne:

  1. Strukturert (meterpostane — sjølve kabelen)
        Kabeltype/kabelkonstruksjon: BFSI
        Ledertall/-materiale:        4+PE/Cu
        Ledertverrsnitt:             95 mm2

  2. Fritekst (stk-postane — terminering av kabelen)
        PUNKT Anvendelse: Terminering/tilkobling av stigeledning …
        Posten omfatter komplett tilkobling/terminering av
        BFXI 5x6 mm^2 Cu stigekabel.

Begge gjev same nøkkel: konstruksjon + ledertall + tverrsnitt, med materiale
(Cu/Al) og jordleiar (+j / +PE) som tilleggskriterium.

Mengde per rad i kabellista:

    Post med enhet "stk"  →  talet på taggar i "Til =ref Tag".
                             Kvar tagg er eitt punkt kabelen går til, og
                             punktprisen gjeld per punkt. Ei rad med 13
                             taggar er 13 stk.
    Post med enhet "m"    →  "Lengde (m)" for rada.

Same kabelgruppe kan difor treffe både ein meterpost (lengda) og ein
termineringspost (talet på punkt).

Bruk:
    from kabel_prognose import les_kabelmengder
    kabel = les_kabelmengder(kabelliste_stiar, isy_stiar)
"""

from __future__ import annotations

import os
import re

import pandas as pd


# =============================================================================
# Konstantar
# =============================================================================

KABEL_ARK = "Kabelliste"

KOL_TIL_TAG = "Til =ref Tag"
KOL_TYPE = "Kabeltype"
KOL_LENGDE = "Lengde (m)"
KOL_KABELNR = "Navn/merkenr."
KOL_OMRAADE = ("Prosjekteringsområde", "Prosjekterings-område")

# Kva prosesskodeprefiks kabellista høyrer til.
FIL_TIL_PREFIKS = (("LYA", "1303"), ("VEK", "1302"))

# Kjende kabelkonstruksjonar. Lista styrer kva som vert lese som konstruksjon
# i fritekst, slik at vilkårlege ord ikkje vert tolka som kabeltype.
KONSTRUKSJONAR = (
    "BFSI", "BFXI", "IFSI", "IFXI", "PFSP", "PFXP", "RFOU", "TFXP",
    "TSLF", "TXXP", "FXQJ", "EXQJ",
)

_KON_ALT = "|".join(sorted(set(KONSTRUKSJONAR), key=len, reverse=True))

# Strukturerte NS3420-felt (meterpostane)
_RE_KONSTRUKSJON = re.compile(r"Kabeltype/kabelkonstruksjon:\s*([^\s,;]+)", re.IGNORECASE)
_RE_LEDERTALL = re.compile(r"Ledertall/-?materiale:\s*([^\n\r]+)", re.IGNORECASE)
_RE_TVERRSNITT = re.compile(r"Ledertverrsnitt:\s*(\d+(?:[.,]\d+)?)\s*mm", re.IGNORECASE)

# Fritekst: "BFXI 5x6 mm^2 Cu", "BFXI BFSI 4x150 mm^2 Cu + j",
#           "IFSI 4x25 mm^2 Al stigekabel + PE skjerm"
_RE_FRITEKST = re.compile(
    rf"((?:(?:{_KON_ALT})\s+)+)"
    rf"(\d+)\s*[xX*]\s*(\d+(?:[.,]\d+)?)"
    rf"\s*(?:mm)",
    re.IGNORECASE,
)

# Kabeltype i kabellista: "BFSI 4x95+j Cu", "BFXI 5x2,5", "IFSI 4x25 Al + j"
_RE_LISTETYPE = re.compile(
    rf"^\s*((?:(?:{_KON_ALT})\s*)+)"
    rf"(\d+)\s*[xX*]\s*(\d+(?:[.,]\d+)?)",
    re.IGNORECASE,
)

_RE_JORD = re.compile(r"\+\s*(?:j\b|pe\b)", re.IGNORECASE)
_RE_MATERIALE = re.compile(r"\b(cu|al)\b", re.IGNORECASE)


# =============================================================================
# Parsing
# =============================================================================

def _tal(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if pd.notna(v) else None
    s = str(v).strip().replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _nokkel(konstruksjon: str, ledertall: int, tverrsnitt: float) -> tuple:
    return (konstruksjon.upper(), int(ledertall), round(float(tverrsnitt), 3))


def parse_kabeltype(tekst) -> dict | None:
    """"BFSI 4x95+j Cu" -> nøkkel {("BFSI", 4, 95.0)}, jord=True, materiale=CU."""
    s = str(tekst or "").strip()
    if not s:
        return None
    m = _RE_LISTETYPE.match(s)
    if not m:
        return None
    konstruksjonar = re.findall(_KON_ALT, m.group(1), re.IGNORECASE)
    ledertall = int(m.group(2))
    tverrsnitt = float(m.group(3).replace(",", "."))
    mat = _RE_MATERIALE.search(s)
    return {
        "nokler": {_nokkel(k, ledertall, tverrsnitt) for k in konstruksjonar},
        "jord": bool(_RE_JORD.search(s)),
        "materiale": mat.group(1).upper() if mat else None,
    }


def parse_ns3420_kabel(tekst) -> dict | None:
    """Kabelnøkkel frå NS3420-tekst: strukturert form først, så fritekst."""
    s = str(tekst or "").replace("_x000D_", "\n")
    if not s.strip():
        return None

    mk = _RE_KONSTRUKSJON.search(s)
    mt = _RE_TVERRSNITT.search(s)
    ml = _RE_LEDERTALL.search(s)
    if mk and mt and ml:
        mn = re.search(r"(\d+)", ml.group(1))
        if mn:
            mat = _RE_MATERIALE.search(ml.group(1))
            return {
                "nokler": {_nokkel(mk.group(1), int(mn.group(1)),
                                   float(mt.group(1).replace(",", ".")))},
                "jord": bool(re.search(r"\+\s*(pe|j)\b", ml.group(1), re.IGNORECASE)),
                "materiale": mat.group(1).upper() if mat else None,
                "form": "strukturert",
            }

    mf = _RE_FRITEKST.search(s)
    if not mf:
        return None
    konstruksjonar = re.findall(_KON_ALT, mf.group(1), re.IGNORECASE)
    ledertall = int(mf.group(2))
    tverrsnitt = float(mf.group(3).replace(",", "."))
    hale = s[mf.end(): mf.end() + 80]
    mat = _RE_MATERIALE.search(hale)
    return {
        "nokler": {_nokkel(k, ledertall, tverrsnitt) for k in konstruksjonar},
        "jord": bool(_RE_JORD.search(hale)),
        "materiale": mat.group(1).upper() if mat else None,
        "form": "fritekst",
    }


def tell_punkt(til_tag) -> int:
    """Tal på punkt kabelen går til = tal på taggar i 'Til =ref Tag'."""
    s = str(til_tag or "").strip()
    if not s or s.lower() in ("nan", "none", "-"):
        return 0
    return len([t for t in re.split(r"[,;\n]+", s) if t.strip()])


def _prefiks_for_fil(sti: str) -> str:
    navn = os.path.basename(sti).upper()
    for nokkel, prefiks in FIL_TIL_PREFIKS:
        if nokkel in navn:
            return prefiks
    return ""


# =============================================================================
# Les kjeldene
# =============================================================================

def les_kabelliste(sti: str) -> pd.DataFrame:
    """Les arket 'Kabelliste'. Rad 1 er gruppeoverskrifter, rad 2 er headeren."""
    df = pd.read_excel(sti, sheet_name=KABEL_ARK, engine="openpyxl", header=1)
    df.columns = [str(c).strip() for c in df.columns]

    omraadekol = next((k for k in KOL_OMRAADE if k in df.columns), None)
    manglar = [k for k in (KOL_TIL_TAG, KOL_TYPE, KOL_LENGDE) if k not in df.columns]
    if manglar:
        raise KeyError(f"{os.path.basename(sti)}: manglar kolonnane {manglar}")

    ut = pd.DataFrame({
        "kabelnr": df.get(KOL_KABELNR, "").astype(str).str.strip(),
        "omraade": (df[omraadekol].astype(str).str.strip() if omraadekol else ""),
        "kabeltype": df[KOL_TYPE].astype(str).str.strip(),
        "lengde_m": df[KOL_LENGDE].map(_tal),
        "antall_punkt": df[KOL_TIL_TAG].map(tell_punkt),
    })
    ut = ut[~ut["kabeltype"].str.lower().isin(["", "nan", "none"])].copy()
    ut["prefiks"] = _prefiks_for_fil(sti)
    ut["kilde"] = os.path.basename(sti)
    ut["parsed"] = ut["kabeltype"].map(parse_kabeltype)

    utolka = ut.loc[ut["parsed"].isna(), "kabeltype"]
    print(f"   Kabelliste {os.path.basename(sti)}: {len(ut):,} kablar, "
          f"{int(ut['antall_punkt'].sum()):,} punkt, "
          f"{ut['lengde_m'].fillna(0).sum():,.0f} m")
    if len(utolka):
        print(f"     [ADVARSEL] {utolka.nunique()} kabeltypar kunne ikkje tolkast: "
              f"{sorted(set(utolka))[:6]}")
    return ut[ut["parsed"].notna()].copy()


def les_ns3420_kabelposter(isy_path) -> pd.DataFrame:
    """Plukk ut NS3420-postane som skildrar ein kabel, og tolk kabelnøkkelen."""
    stiar = [isy_path] if isinstance(isy_path, str) else list(isy_path)
    rader = []
    for sti in stiar:
        if not sti or not os.path.exists(sti):
            continue
        df = pd.read_excel(sti, engine="openpyxl")
        postkol = next((k for k in ("Post nummer",
                                    "ProsessID: prosess, sted og element")
                        if k in df.columns), None)
        tekstkol = next((k for k in ("All tekst",
                                     "Tekst - alt. (uten linjeskift)")
                         if k in df.columns), None)
        if postkol is None or tekstkol is None:
            continue
        for _, r in df.iterrows():
            info = parse_ns3420_kabel(r.get(tekstkol))
            if not info:
                continue
            post = str(r.get(postkol, "")).strip()
            if not post or post.lower() == "nan":
                continue
            rader.append({
                "post": post,
                "nokler": info["nokler"],
                "jord": info["jord"],
                "materiale": info["materiale"],
                "form": info.get("form", ""),
                "enhet": str(r.get("Enhet", "") or "").strip().lower(),
                "kontraktsmengde": _tal(r.get("Mengde")),
            })
    df = pd.DataFrame(rader).drop_duplicates(subset=["post"], keep="first")
    if df.empty:
        print("   NS3420 kabelposter: 0")
        return df
    print(f"   NS3420 kabelposter: {len(df):,} "
          f"({(df['enhet'] == 'stk').sum()} stk, {(df['enhet'] == 'm').sum()} m; "
          f"{(df['form'] == 'strukturert').sum()} strukturert, "
          f"{(df['form'] == 'fritekst').sum()} fritekst)")
    return df


# =============================================================================
# Kopling
# =============================================================================

def _passar(kabel: dict, post) -> bool:
    """Jordleiar og materiale skal ikkje motseie kvarandre."""
    if kabel["jord"] != post["jord"]:
        return False
    if kabel["materiale"] and post["materiale"] and kabel["materiale"] != post["materiale"]:
        return False
    return True


def les_kabelmengder(kabelliste_stiar, isy_path) -> dict:
    """Returner {postnummer: {"mengde", "enhet", "antall_kabler", …}}."""
    stiar = [kabelliste_stiar] if isinstance(kabelliste_stiar, str) else list(kabelliste_stiar)
    stiar = [s for s in stiar if s and os.path.exists(s)]
    if not stiar:
        print("   Ingen kabellister oppgitt — hoppar over kabelprognosen")
        return {}

    kabler = pd.concat([les_kabelliste(s) for s in stiar], ignore_index=True)
    poster = les_ns3420_kabelposter(isy_path)
    if poster.empty or kabler.empty:
        return {}

    indeks: dict = {}
    for _, post in poster.iterrows():
        for n in post["nokler"]:
            indeks.setdefault(n, []).append(post)

    kabler["_grp"] = kabler.apply(
        lambda r: (tuple(sorted(r["parsed"]["nokler"])),
                   r["parsed"]["jord"], r["parsed"]["materiale"], r["prefiks"]),
        axis=1,
    )

    resultat: dict = {}
    uten_treff: dict = {}

    for (nokler, jord, materiale, prefiks), grp in kabler.groupby("_grp"):
        kabelinfo = {"jord": jord, "materiale": materiale}

        kandidatar = []
        for n in nokler:
            for post in indeks.get(n, []):
                if prefiks and not str(post["post"]).startswith(prefiks):
                    continue
                kandidatar.append(post)

        if not kandidatar:
            uten_treff[(nokler, jord, materiale, prefiks)] = {
                "kabler": len(grp),
                "punkt": int(grp["antall_punkt"].sum()),
                "meter": float(grp["lengde_m"].fillna(0).sum()),
                "typer": sorted(set(grp["kabeltype"])),
            }
            continue

        eksakte = [p for p in kandidatar if _passar(kabelinfo, p)]
        bruk = eksakte or kandidatar
        match = "eksakt" if eksakte else "uten jord/materiale"

        for post in bruk:
            if post["enhet"] == "stk":
                mengde = float(grp["antall_punkt"].sum())
            elif post["enhet"] in ("m", "lm"):
                mengde = float(grp["lengde_m"].fillna(0).sum())
            else:
                mengde = float(len(grp))

            rad = resultat.setdefault(str(post["post"]), {
                "mengde": 0.0, "enhet": post["enhet"], "antall_kabler": 0,
                "kabeltyper": set(), "kilde": set(), "match": set(),
            })
            rad["mengde"] += mengde
            rad["antall_kabler"] += len(grp)
            rad["kabeltyper"].update(grp["kabeltype"])
            rad["kilde"].update(grp["kilde"])
            rad["match"].add(match)

    for rad in resultat.values():
        rad["kabeltyper"] = "; ".join(sorted(rad["kabeltyper"]))
        rad["kilde"] = "; ".join(sorted(rad["kilde"]))
        rad["match"] = "; ".join(sorted(rad["match"]))

    n_stk = sum(1 for v in resultat.values() if v["enhet"] == "stk")
    n_m = sum(1 for v in resultat.values() if v["enhet"] in ("m", "lm"))
    print(f"   Kabelposter med prognosemengde: {len(resultat):,} ({n_stk} stk, {n_m} m)")

    if uten_treff:
        tapt_punkt = sum(v["punkt"] for v in uten_treff.values())
        tapt_meter = sum(v["meter"] for v in uten_treff.values())
        print(f"   [ADVARSEL] {len(uten_treff)} kabeltype/område-kombinasjonar "
              f"fann ingen NS3420-post ({tapt_punkt:,} punkt, {tapt_meter:,.0f} m):")
        for (nokler, jord, mat, prefiks), v in sorted(
                uten_treff.items(), key=lambda kv: -kv[1]["punkt"])[:12]:
            print(f"        {prefiks or '??'}  {', '.join(v['typer'][:2]):<28} "
                  f"{v['punkt']:>6} punkt, {v['meter']:>8,.0f} m")
    return resultat


def kabelrapport(kabelmengder: dict) -> pd.DataFrame:
    """Gjer resultatet om til ein DataFrame for skriving til Excel."""
    kolonner = ["post", "kabelmengde", "kabelenhet", "antall_kabler",
                "kabeltyper", "kabelkilde", "kabelmatch"]
    if not kabelmengder:
        return pd.DataFrame(columns=kolonner)
    return pd.DataFrame([
        {
            "post": post,
            "kabelmengde": v["mengde"],
            "kabelenhet": v["enhet"],
            "antall_kabler": v["antall_kabler"],
            "kabeltyper": v["kabeltyper"],
            "kabelkilde": v["kilde"],
            "kabelmatch": v["match"],
        }
        for post, v in sorted(kabelmengder.items())
    ])


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="Kabelmengder frå kabellister")
    p.add_argument("--kabelliste", nargs="+", required=True)
    p.add_argument("--isy", nargs="+", required=True)
    a = p.parse_args()
    print(kabelrapport(les_kabelmengder(a.kabelliste, a.isy)).to_string(index=False))
