# -*- coding: utf-8 -*-
"""
Finn kvar i kjeda ein mengde blir feil.

Mengda går gjennom fire ledd, og kvart ledd les det førre:

    CSV (selection basket)  ->  Mengdeuttak_Armering
                            ->  ARM/RIB-excel, arket "Sammendrag"
                            ->  mengde_oversikt_*.xlsx
                            ->  prognose_*.xlsx

Er talet feil i prognosen, men rett i Sammendrag, er prognosen bygd på ei
gammal mengde_oversikt. Er det feil allereie i CSV-en — fleire rader enn
GUID-ar — er same element lese to gonger, og då doblar mengda seg uansett
kor rett resten reknar.

Skriptet skriv berre ut tal for den posten du spør om. Ingenting vert sendt
nokon stad.

Bruk:
    python sjekk_volum_kjede.py 1303.01.05.241.8 ^
        --csv      "selection-basket.csv" ^
        --arm-rib  "resultat_mengdeuttak_....xlsx" ^
        --oversikt "mengde_oversikt_....xlsx" ^
        --prognose "prognose_....xlsx"

Alle fil-argumenta er valfrie — ta med dei du har.
"""

from __future__ import annotations

import argparse
import os
import sys

import pandas as pd


def _norm(s) -> str:
    return " ".join(str(s or "").split()).strip().lower()


def _finn_kol(df: pd.DataFrame, *namn) -> str | None:
    """Finn første kolonne som matchar eitt av namna (ufølsam for store bokstavar)."""
    ynskje = [_norm(n) for n in namn]
    for kol in df.columns:
        if _norm(kol) in ynskje:
            return kol
    return None


def _tal(v) -> float:
    """Parse eit tal som kan ha einingssuffiks: '12.793m³' -> 12.793."""
    s = str(v or "").strip()
    if not s:
        return 0.0
    ut = []
    for c in s:
        if c.isdigit() or c in ".,-":
            ut.append("." if c == "," else c)
        elif ut:
            break
    try:
        return float("".join(ut))
    except ValueError:
        return 0.0


def _les_ark(sti: str, header: int = 0) -> dict[str, pd.DataFrame]:
    """Les alle ark i ei Excel-fil. Returnerer {arknamn: df}."""
    try:
        return pd.read_excel(sti, sheet_name=None, engine="openpyxl", header=header)
    except Exception as feil:
        print(f"   [FEIL] Klarte ikkje lese {os.path.basename(sti)}: {feil}")
        return {}


def _treff(df: pd.DataFrame, postkol: str, post: str) -> pd.DataFrame:
    """Rader der postkolonnen er posten sjølv eller ein underpost av han."""
    s = df[postkol].astype(str).str.strip()
    return df[(s == post) | s.str.startswith(post + ".")]


def _overskrift(tekst: str) -> None:
    strek = "=" * 70
    print(f"\n{strek}\n{tekst}\n{strek}")


# =============================================================================
# Ledd 0 — CSV: er same element lese fleire gonger?
# =============================================================================

def rapporter_csv(stiar: list, post: str) -> None:
    _overskrift("LEDD 0 — CSV (raadata)")

    deler = []
    for sti in stiar:
        if not os.path.exists(sti):
            print(f"   [FEIL] Finn ikkje {sti}")
            continue
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                d = pd.read_csv(sti, sep=None, engine="python",
                                encoding=enc, dtype=str)
                deler.append(d)
                print(f"   Las {os.path.basename(sti)}: {len(d):,} rader")
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
    if not deler:
        return

    df = pd.concat(deler, ignore_index=True) if len(deler) > 1 else deler[0]

    postkol = _finn_kol(df, "FOB_Info/FOB_Mengdelistepost",
                        "FOB_Info.FOB_Mengdelistepost")
    volkol = _finn_kol(df, "PGF_Kalkyle/PGF_Volum", "PGF_Kalkyle.PGF_Volum")
    guidkol = _finn_kol(df, "GUID")
    if postkol is None or guidkol is None:
        print("   [FEIL] Fann ikkje GUID- og/eller mengdelistepost-kolonne")
        return

    def har_posten(celle) -> bool:
        for bit in str(celle or "").split(";"):
            b = bit.strip()
            if b == post or b.startswith(post + "."):
                return True
        return False

    rader = df[df[postkol].map(har_posten)]
    if rader.empty:
        print(f"   Ingen rader med posten {post}")
        return

    guid = rader[guidkol].astype(str).str.strip()
    n_rader, n_unike = len(rader), guid.nunique()
    print(f"\n   Rader med posten:      {n_rader:>12,}")
    print(f"   Unike GUID-ar:         {n_unike:>12,}")

    if n_rader == n_unike:
        print("\n   Kvart element er lese éin gong — inga dobbeltlesing her.")
    else:
        print(f"\n   [!] {n_rader - n_unike:,} rader er same element "
              f"lese fleire gonger")

    if volkol is not None:
        alle = rader[volkol].map(_tal).sum()
        unike = rader.assign(_g=guid).drop_duplicates("_g")[volkol].map(_tal).sum()
        merke = "   <-- dette er rett" if abs(alle - unike) > 1e-9 else ""
        print(f"\n   Volum, alle rader:     {alle:>12,.4f}")
        print(f"   Volum, unike GUID:     {unike:>12,.4f}{merke}")

    if n_rader != n_unike and "Path" in rader.columns:
        dobbel = rader[guid.duplicated(keep="first").values]
        modell = (dobbel["Path"].astype(str)
                  .str.extract(r"([^/\\]+\.ifc)", expand=False)
                  .fillna("(utan .ifc i Path)"))
        print("\n   Duplikata kjem frå:")
        for navn, antal in modell.value_counts().head(10).items():
            print(f"      {antal:>8,}  {navn}")


# =============================================================================
# Ledd 1-3 — Excel-filene
# =============================================================================

def rapporter_excel(tittel: str, sti: str, header: int,
                    postkolonner: tuple, verdikolonner: tuple,
                    post: str) -> None:
    _overskrift(f"{tittel}\n  {os.path.basename(sti)}")
    if not os.path.exists(sti):
        print("   [FEIL] Finn ikkje fila")
        return

    funne_noko = False
    for arknamn, df in _les_ark(sti, header=header).items():
        if df is None or df.empty:
            continue
        postkol = _finn_kol(df, *postkolonner)
        verdikol = _finn_kol(df, *verdikolonner)
        if postkol is None or verdikol is None:
            continue

        rader = _treff(df, postkol, post)
        if rader.empty:
            continue

        funne_noko = True
        print(f"\n  Ark '{arknamn}'  —  '{postkol}' / '{verdikol}'")
        sum_pr_post: dict[str, float] = {}
        for _, r in rader.iterrows():
            kode = str(r[postkol]).strip()
            v = pd.to_numeric(r[verdikol], errors="coerce")
            if pd.notna(v):
                sum_pr_post[kode] = sum_pr_post.get(kode, 0.0) + float(v)

        if not sum_pr_post:
            print("     (ingen tal)")
            continue

        for kode in sorted(sum_pr_post):
            merke = "  <-- posten du spurde om" if kode == post else ""
            print(f"     {kode:<30} {sum_pr_post[kode]:>14,.4f}{merke}")
        print(f"     {'SUM (post + underpostar)':<30} "
              f"{sum(sum_pr_post.values()):>14,.4f}")

    if not funne_noko:
        print("   Fann ingen rader for denne posten i denne fila.")


def main() -> int:
    p = argparse.ArgumentParser(
        description="Sjekk kvar i kjeda ein mengde blir feil.")
    p.add_argument("post", help="Prosesskoden, t.d. 1303.01.05.241.8")
    p.add_argument("--csv", nargs="+", default=None,
                   help="Selection-basket CSV-ar — tel GUID-duplikat")
    p.add_argument("--arm-rib", dest="arm_rib", default=None,
                   help="ARM/RIB-excel frå Mengdeuttak_Armering")
    p.add_argument("--oversikt", default=None, help="mengde_oversikt_*.xlsx")
    p.add_argument("--prognose", default=None, help="prognose_*.xlsx")
    a = p.parse_args()

    post = a.post.strip()
    if not (a.csv or a.arm_rib or a.oversikt or a.prognose):
        p.error("Oppgi minst ei fil "
                "(--csv, --arm-rib, --oversikt eller --prognose)")

    print(f"\nPost: {post}")
    print("Underpostar vert viste kvar for seg, så du ser om mengda ligg på")
    print("foreldrekoden eller på den prisberande posten.")

    if a.csv:
        rapporter_csv(a.csv, post)

    if a.arm_rib:
        # Sammendrag har headeren på rad 2 -> header=1
        rapporter_excel("LEDD 1 — ARM/RIB-excel (Sammendrag)", a.arm_rib, 1,
                        ("OMEGA-post: Betong",), ("Volum (m3)",), post)

    if a.oversikt:
        rapporter_excel("LEDD 2 — mengde_oversikt", a.oversikt, 0,
                        ("Mengdelistepost", "Post"),
                        ("Mengde fra modell",), post)

    if a.prognose:
        rapporter_excel("LEDD 3 — prognose (mengde frå modell)", a.prognose, 0,
                        ("Post", "Mengdelistepost"),
                        ("Mengde fra modell",), post)
        rapporter_excel("LEDD 3 — prognose (sjølve prognosen)", a.prognose, 0,
                        ("Post", "Mengdelistepost"),
                        ("Prognose",), post)

    print("\n" + "-" * 70)
    print("Slik les du resultatet:")
    print("  Fleire rader enn GUID-ar i ledd 0")
    print("        -> same element er lese fleire gonger; det doblar mengda")
    print("  Rett i ledd 0/1, feil i ledd 2/3")
    print("        -> kjeda må køyrast på nytt frå ARM/RIB-steget")
    print("  Feil allereie i ledd 1")
    print("        -> utrekninga er framleis feil")
    print("  Mengd på foreldrekoden")
    print("        -> skal liggje på den prisberande underposten i staden")
    return 0


if __name__ == "__main__":
    sys.exit(main())
