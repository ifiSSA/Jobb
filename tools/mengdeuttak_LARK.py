"""
Sammenligner mengdelistepost/kontraktsmengde mellom:
  - Omega_K7_Mars.xlsx  (BOQNT-ark)
  - LARK_planter.xlsx   (Tabell1-ark)

Semikolon-separerte celler i LARK-filen ekspanderes til én rad per element.
Etter elementradene for hver post kommer en totalrad med sum og Omega-data.
"""

import pandas as pd
import numpy as np

OMEGA_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang\Omega_K7_April.xlsx"
)

LARK_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\LARK_planter.xlsx"
)

RB14_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\Mengdelistpost NS3420 isylinker.xlsx"
)

# ── Les Omega ──────────────────────────────────────────────────────────────────
omega_raw = pd.read_excel(OMEGA_PATH, sheet_name="BOQNT", header=2)
# Kolonnenavn for kontraktsmengde varierer mellom Omega-eksporter:
#   Gammelt format (Mars): 'Kontraktsmengde'
#   Nytt format (April):   'Opprinnelig'
if "Kontraktsmengde" in omega_raw.columns:
    kontraktsmengde_kol = "Kontraktsmengde"
elif "Opprinnelig" in omega_raw.columns:
    kontraktsmengde_kol = "Opprinnelig"
else:
    raise ValueError(f"Finner ikke kontraktsmengde-kolonne i Omega. Kolonner: {list(omega_raw.columns)}")

omega = omega_raw[["Navn (Prosesskode)", "Tittel", "Enhet", "Enhetspris", kontraktsmengde_kol]].copy()
omega.columns = ["post", "tittel", "enhet_omega", "enhetspris", "kontraktsmengde"]
omega["post"] = omega["post"].astype(str).str.strip()
omega = omega[omega["post"].notna() & (omega["post"] != "nan")]

omega_sum = (
    omega.groupby("post", sort=False)
    .agg(
        kontraktsmengde=("kontraktsmengde", "sum"),
        enhet_omega=("enhet_omega", "first"),
        enhetspris=("enhetspris", "first"),
        tittel=("tittel", "first"),
    )
    .reset_index()
)

# Oppslag: stripped (uten punktum) -> kanonisk Omega-post
omega_stripped_lookup = {p.replace(".", ""): p for p in omega_sum["post"]}
omega_posts_set = set(omega_sum["post"])

# ── Les RB14 ──────────────────────────────────────────────────────────────────
rb14_raw = pd.read_excel(RB14_PATH, sheet_name="Ark1", header=0)
rb14_raw.columns = [str(c).strip() for c in rb14_raw.columns]
rb14 = rb14_raw[["Post nummer", "Enhet", "Mengde", "Pris: Anbud"]].copy()
rb14.columns = ["post", "enhet_rb14", "mengde_rb14", "enhetspris_rb14"]
rb14["post"] = rb14["post"].astype(str).str.strip()
rb14 = rb14[rb14["post"].notna() & (rb14["post"] != "nan") & (rb14["post"] != "None")]
rb14["mengde_rb14"] = pd.to_numeric(rb14["mengde_rb14"], errors="coerce")
rb14 = rb14[rb14["mengde_rb14"].notna() & (rb14["mengde_rb14"] > 0)]

rb14_sum = (
    rb14.groupby("post", sort=False)
    .agg(
        mengde_rb14=("mengde_rb14", "sum"),
        enhet_rb14=("enhet_rb14", "first"),
        enhetspris_rb14=("enhetspris_rb14", "first"),
    )
    .reset_index()
)
rb14_sum["mengde_rb14"] = rb14_sum["mengde_rb14"].round(2)

# ── Les LARK og ekspander til én rad per element ───────────────────────────────
lark_xl   = pd.ExcelFile(LARK_PATH)
lark_sheet = "LARK" if "LARK" in lark_xl.sheet_names else lark_xl.sheet_names[0]
lark_raw  = pd.read_excel(lark_xl, sheet_name=lark_sheet, header=0)
lark_raw.columns = [c.strip() for c in lark_raw.columns]

element_rows = []
for _, r in lark_raw.iterrows():
    posts_raw   = str(r["FOB_Info.FOB_Mengdelistepost"]) if pd.notna(r["FOB_Info.FOB_Mengdelistepost"]) else ""
    mengder_raw = str(r["FOB_Info.FOB_Mengde"])          if pd.notna(r["FOB_Info.FOB_Mengde"])          else ""
    enheter_raw = str(r["FOB_Info.FOB_Mengdeenhet"])     if pd.notna(r["FOB_Info.FOB_Mengdeenhet"])     else ""
    mengdetype  = str(r["PGF_Info.PGF_Mengdetype"])      if pd.notna(r["PGF_Info.PGF_Mengdetype"])      else ""
    fob_status  = str(r["FOB_Info.FOB_Status"])           if pd.notna(r["FOB_Info.FOB_Status"])           else ""

    posts   = [p.strip() for p in posts_raw.split(";")]
    mengder = [m.strip() for m in mengder_raw.split(";")]
    enheter = [e.strip() for e in enheter_raw.split(";")]

    for i, post in enumerate(posts):
        if not post or post == "nan":
            continue
        # Normaliser: hvis post ikke finnes direkte i Omega, prøv uten punktum
        if post not in omega_posts_set:
            canonical = omega_stripped_lookup.get(post.replace(".", ""))
            if canonical:
                post = canonical
        try:
            mengde = float(mengder[i]) if i < len(mengder) else 0.0
        except ValueError:
            mengde = 0.0
        enhet = enheter[i] if i < len(enheter) else ""
        element_rows.append({
            "post":        post,
            "rad_type":    "Element",
            "tittel":      "",
            "mengdetype":  mengdetype,
            "fob_status":  fob_status,
            "enhet_lark":  enhet,
            "mengde_lark": round(mengde, 2),
        })

lark_exp = pd.DataFrame(element_rows)

# ── Bygg totalrader per post ───────────────────────────────────────────────────
lark_sum = (
    lark_exp.groupby("post", sort=False)
    .agg(
        mengde_lark=("mengde_lark", "sum"),
        enhet_lark=("enhet_lark", "first"),
        mengdetype=("mengdetype", "first"),
        fob_status=("fob_status", "first"),
    )
    .reset_index()
)
lark_sum["rad_type"] = "TOTAL"
lark_sum["mengde_lark"] = lark_sum["mengde_lark"].round(2)

# Koble på Omega-data (kun for totalrader)
totals = pd.merge(lark_sum, omega_sum, on="post", how="left")
totals["kontraktsmengde"] = totals["kontraktsmengde"].round(2)
totals["avvik"]           = (totals["mengde_lark"] - totals["kontraktsmengde"]).round(2)
totals["kost_lark"]       = (totals["enhetspris"] * totals["mengde_lark"]).round(2)
totals["kost_kontrakt"]   = (totals["enhetspris"] * totals["kontraktsmengde"]).round(2)
totals["mengde_pst"]      = (totals["mengde_lark"] / totals["kontraktsmengde"] * 100 - 100).round(1)
totals["kilde"]           = totals["enhet_omega"].apply(lambda x: "Begge" if pd.notna(x) else "Kun i LARK")
totals["tittel"]          = totals["tittel"].fillna("")

# Koble på RB14-data
totals = pd.merge(totals, rb14_sum, on="post", how="left")
totals["avvik_rb14"]      = (totals["mengde_lark"] - totals["mengde_rb14"]).round(2)
totals["kost_rb14"]       = (totals["enhetspris_rb14"] * totals["mengde_rb14"]).round(2)
totals["mengde_pst_rb14"] = (totals["mengde_lark"] / totals["mengde_rb14"] * 100 - 100).round(1)
totals["kilde_rb14"]      = totals["enhet_rb14"].apply(lambda x: "Begge" if pd.notna(x) else "Kun i LARK")

# Bygg oppslag post -> tittel for elementrader
post_tittel = dict(zip(totals["post"], totals["tittel"]))
lark_exp["tittel"] = lark_exp["post"].map(post_tittel).fillna("")

# Element-rader har ikke Omega-kolonner – fyll med NaN
for col in ["kontraktsmengde", "enhet_omega", "enhetspris", "avvik",
            "kost_lark", "kost_kontrakt", "mengde_pst", "kilde",
            "mengde_rb14", "enhet_rb14", "enhetspris_rb14",
            "avvik_rb14", "mengde_pst_rb14", "kost_rb14", "kilde_rb14"]:
    lark_exp[col] = np.nan

# ── Sett sammen: elementer + totalrader, sorter per post ──────────────────────
all_cols = ["post", "rad_type", "tittel", "mengdetype", "fob_status",
            "enhet_lark", "enhet_omega", "enhetspris",
            "mengde_lark", "kontraktsmengde", "avvik", "mengde_pst",
            "kost_lark", "kost_kontrakt", "kilde",
            "mengde_rb14", "enhet_rb14", "enhetspris_rb14",
            "avvik_rb14", "mengde_pst_rb14", "kost_rb14", "kilde_rb14"]

combined = pd.concat([lark_exp[all_cols], totals[all_cols]], ignore_index=True)
combined = combined.sort_values(["post", "rad_type"], ascending=[True, False])
# rad_type: "Element" < "TOTAL" alfabetisk -> Element kommer først, TOTAL sist

combined = combined.rename(columns={
    "post":              "Mengdelistepost",
    "rad_type":          "Rad",
    "tittel":            "Tittel (Omega)",
    "mengdetype":        "Mengdetype (LARK)",
    "fob_status":        "FOB Status",
    "enhet_lark":        "Enhet (LARK)",
    "enhet_omega":       "Enhet (Omega)",
    "enhetspris":        "Enhetspris (Omega)",
    "mengde_lark":       "Mengde LARK",
    "kontraktsmengde":   "Kontraktsmengde (Omega)",
    "avvik":             "Avvik (LARK - Omega)",
    "mengde_pst":        "Mengdeøkning (%)",
    "kost_lark":         "Kostnad LARK (EP*Mengde)",
    "kost_kontrakt":     "Kostnad kontrakt (EP*Kontraktsmengde)",
    "kilde":             "Kilde",
    "mengde_rb14":       "Mengde RB14",
    "enhet_rb14":        "Enhet (RB14)",
    "enhetspris_rb14":   "Enhetspris (RB14)",
    "avvik_rb14":        "Avvik (LARK - RB14)",
    "mengde_pst_rb14":   "Mengdeøkning RB14 (%)",
    "kost_rb14":         "Kostnad RB14 (EP*Mengde)",
    "kilde_rb14":        "Kilde RB14",
})

col_order = [
    "Mengdelistepost",
    "Rad",
    "Tittel (Omega)",
    "Mengdetype (LARK)",
    "FOB Status",
    "Enhet (LARK)",
    "Enhet (Omega)",
    "Enhetspris (Omega)",
    "Mengde LARK",
    "Kontraktsmengde (Omega)",
    "Avvik (LARK - Omega)",
    "Mengdeøkning (%)",
    "Kostnad LARK (EP*Mengde)",
    "Kostnad kontrakt (EP*Kontraktsmengde)",
    "Kilde",
    "Mengde RB14",
    "Enhet (RB14)",
    "Enhetspris (RB14)",
    "Avvik (LARK - RB14)",
    "Mengdeøkning RB14 (%)",
    "Kostnad RB14 (EP*Mengde)",
    "Kilde RB14",
]
combined = combined[col_order]

# ── Bygg totalrad-ark ─────────────────────────────────────────────────────────
totals_sheet = combined[combined["Rad"] == "TOTAL"].copy()

# ── Bygg sammendragsark ───────────────────────────────────────────────────────
POST_KAT = {
    "1303.03.71.771.4":  "Pisk 1303.03.71.771.4",
    "1303.03.71.772.13": "Busker 1303.03.71.772.13",
    "1303.03.71.772.14": "Stauder 1303.03.71.772.14",
    "1303.03.71.772.15": "Trær 1303.03.71.772.15",
}

def kategoriser_post(post, enhet):
    if str(enhet).strip().lower() != "stk":
        return None
    for prefix, kat in POST_KAT.items():
        if str(post).startswith(prefix):
            return kat
    return None

ts = totals_sheet.copy()
ts["kategori"] = ts.apply(
    lambda r: kategoriser_post(r["Mengdelistepost"], r["Enhet (LARK)"]), axis=1
)
ts_cat = ts[ts["kategori"].notna()]

lark_cat = (
    ts_cat.groupby("kategori", sort=False)["Mengde LARK"]
    .sum().round(0).astype(int).reset_index()
    .rename(columns={"Mengde LARK": "Mengde LARK (modell)"})
)
omega_cat = (
    ts_cat.groupby("kategori", sort=False)["Kontraktsmengde (Omega)"]
    .sum().round(0).astype(int).reset_index()
)

sammendrag = pd.merge(lark_cat, omega_cat, on="kategori", how="outer")
rb14_cat = (
    ts_cat.groupby("kategori", sort=False)["Mengde RB14"]
    .sum().round(0).reset_index()
)
sammendrag = pd.merge(sammendrag, rb14_cat, on="kategori", how="outer").fillna(0)
sammendrag["Mengde LARK (modell)"]    = sammendrag["Mengde LARK (modell)"].astype(int)
sammendrag["Kontraktsmengde (Omega)"] = sammendrag["Kontraktsmengde (Omega)"].astype(int)
sammendrag["Mengde RB14"]             = sammendrag["Mengde RB14"].astype(int)
sammendrag["Økning vs Omega (stk)"]  = sammendrag["Mengde LARK (modell)"] - sammendrag["Kontraktsmengde (Omega)"]
sammendrag["Økning vs Omega (%)"]    = (
    (sammendrag["Mengde LARK (modell)"] / sammendrag["Kontraktsmengde (Omega)"] * 100 - 100)
    .round(1)
)
sammendrag["Økning vs RB14 (stk)"]   = sammendrag["Mengde LARK (modell)"] - sammendrag["Mengde RB14"]
sammendrag["Økning vs RB14 (%)"]     = (
    (sammendrag["Mengde LARK (modell)"] / sammendrag["Mengde RB14"] * 100 - 100)
    .round(1)
)
kat_order = ["Pisk 1303.03.71.771.4", "Busker 1303.03.71.772.13", "Stauder 1303.03.71.772.14", "Trær 1303.03.71.772.15"]
sammendrag["kategori"] = pd.Categorical(sammendrag["kategori"], categories=kat_order, ordered=True)
sammendrag = sammendrag.sort_values("kategori").rename(columns={"kategori": "Kategori"})

# ── Bygg Merknad-ark (kubikk + tekst fra Merknad2 og Merknad3) ───────────────
import re

def parse_merknad(val):
    """Returnerer (kubikk_float, tekst) fra strenger som '6m3 Moldholdig anleggsjord'."""
    if pd.isna(val) or str(val).strip() in ("", "-", "nan"):
        return (None, None)
    s = str(val).strip()
    m = re.match(r'^([\d]+(?:[.,][\d]+)?)\s*m3\s+(.*)', s, re.IGNORECASE)
    if m:
        kubikk = float(m.group(1).replace(",", "."))
        tekst  = m.group(2).strip()
        return (kubikk, tekst)
    return (None, s)   # tekst uten kubikk-verdi

merknad_rows = []
for _, r in lark_raw.iterrows():
    post_raw = str(r["FOB_Info.FOB_Mengdelistepost"]) if pd.notna(r["FOB_Info.FOB_Mengdelistepost"]) else ""
    # Bruk første post som identifikator
    foerste_post = post_raw.split(";")[0].strip() if post_raw else ""
    if foerste_post and foerste_post != "nan":
        if foerste_post not in omega_posts_set:
            canonical = omega_stripped_lookup.get(foerste_post.replace(".", ""))
            if canonical:
                foerste_post = canonical

    m2_raw = r["PGF_LARK_Info.Merknad2"] if "PGF_LARK_Info.Merknad2" in lark_raw.columns else None
    m3_raw = r["PGF_LARK_Info.Merknad3"] if "PGF_LARK_Info.Merknad3" in lark_raw.columns else None
    m2_kub, m2_txt = parse_merknad(m2_raw)
    m3_kub, m3_txt = parse_merknad(m3_raw)

    if m2_kub is None and m2_txt is None and m3_kub is None and m3_txt is None:
        continue

    merknad_rows.append({
        "Mengdelistepost":   foerste_post,
        "Klassebeskrivelse": str(r["PGF_Info.PGF_Klassebeskrivelse"]).strip() if pd.notna(r["PGF_Info.PGF_Klassebeskrivelse"]) else "",
        "Type":              str(r["Type"]).strip() if pd.notna(r["Type"]) else "",
        "FOB Status":        str(r["FOB_Info.FOB_Status"]).strip() if pd.notna(r["FOB_Info.FOB_Status"]) else "",
        "Merknad2 (råverdi)": str(m2_raw).strip() if pd.notna(m2_raw) else "",
        "Merknad2 kubikk (m3)": m2_kub,
        "Merknad2 tekst":    m2_txt if m2_txt else "",
        "Merknad3 (råverdi)": str(m3_raw).strip() if pd.notna(m3_raw) else "",
        "Merknad3 kubikk (m3)": m3_kub,
        "Merknad3 tekst":    m3_txt if m3_txt else "",
    })

merknad_df = pd.DataFrame(merknad_rows)
merknad_df = merknad_df.sort_values(
    ["Mengdelistepost", "Merknad3 tekst", "Merknad2 tekst"],
    na_position="last"
).reset_index(drop=True)

# ── Skriv til Excel som formatert tabell ──────────────────────────────────────
import os
OUTPUT = os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "excel"), "mengdeuttak_LARK_resultat.xlsx")
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    # Ark 1: Sammendrag
    sammendrag.to_excel(writer, sheet_name="Sammendrag", index=False)
    # Ark 2: Bare totalrader
    totals_sheet.to_excel(writer, sheet_name="Totalrader", index=False)
    # Ark 3: Full detaljert liste
    combined.to_excel(writer, sheet_name="Sammenligning", index=False)
    # Ark 4: Merknad / kubikk
    merknad_df.to_excel(writer, sheet_name="Merknad", index=False)

    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    ws = writer.sheets["Sammenligning"]

    # Definer Excel-tabell
    n_rows = len(combined)
    n_cols = len(col_order)
    table_ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"
    tbl = Table(displayName="MengdeLARK", ref=table_ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)

    red       = PatternFill("solid", fgColor="FF9999")
    yellow    = PatternFill("solid", fgColor="FFFF99")
    offwhite  = PatternFill("solid", fgColor="F5F5F0")
    bold      = Font(bold=True)

    rad_col   = col_order.index("Rad") + 1
    kilde_col = col_order.index("Kilde") + 1
    avvik_col = col_order.index("Avvik (LARK - Omega)") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        rad_val   = row[rad_col - 1].value
        kilde_val = row[kilde_col - 1].value
        avvik_val = row[avvik_col - 1].value

        if rad_val == "TOTAL":
            if kilde_val == "Kun i LARK":
                fill = red
            elif avvik_val is not None and abs(avvik_val) > 0.01:
                fill = yellow
            else:
                fill = None
            for cell in row:
                cell.font = bold
                if fill:
                    cell.fill = fill
        else:
            for cell in row:
                cell.fill = offwhite

    # Juster kolonnebredder Sammenligning
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

    # ── Formater Totalrader-arket ─────────────────────────────────────────────
    ws_tot = writer.sheets["Totalrader"]
    tot_cols = list(totals_sheet.columns)
    tbl_tot = Table(displayName="Totalrader", ref=f"A1:{get_column_letter(len(tot_cols))}{len(totals_sheet)+1}")
    tbl_tot.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_tot.add_table(tbl_tot)

    kilde_t = tot_cols.index("Kilde") + 1
    avvik_t = tot_cols.index("Avvik (LARK - Omega)") + 1
    for row in ws_tot.iter_rows(min_row=2, max_row=ws_tot.max_row):
        kilde_val = row[kilde_t - 1].value
        avvik_val = row[avvik_t - 1].value
        if kilde_val == "Kun i LARK":
            fill = red
        elif avvik_val is not None and abs(avvik_val) > 0.01:
            fill = yellow
        else:
            fill = None
        for cell in row:
            cell.font = bold
            if fill:
                cell.fill = fill
    for col in ws_tot.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws_tot.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

    # ── Formater Sammendrag-arket ─────────────────────────────────────────────
    ws_sum = writer.sheets["Sammendrag"]
    sam_cols = list(sammendrag.columns)
    tbl_sam = Table(displayName="Sammendrag", ref=f"A1:{get_column_letter(len(sam_cols))}{len(sammendrag)+1}")
    tbl_sam.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_sum.add_table(tbl_sam)

    green_fill = PatternFill("solid", fgColor="C6EFCE")
    red_fill   = PatternFill("solid", fgColor="FFC7CE")
    bold_large = Font(bold=True, size=12)

    okning_col = sam_cols.index("Økning vs Omega (stk)") + 1
    for row in ws_sum.iter_rows(min_row=2, max_row=ws_sum.max_row):
        okning_val = row[okning_col - 1].value
        fill = green_fill if (okning_val is not None and okning_val > 0) else red_fill
        for cell in row:
            cell.font  = bold_large
            cell.fill  = fill
            cell.alignment = Alignment(horizontal="center")
    # Header fet
    for cell in ws_sum[1]:
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center")
    for col in ws_sum.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws_sum.column_dimensions[col[0].column_letter].width = min(max_len + 6, 35)
    ws_sum.row_dimensions[1].height = 20

    # ── Formater Merknad-arket ────────────────────────────────────────────────
    ws_mrk = writer.sheets["Merknad"]
    mrk_cols = list(merknad_df.columns)
    tbl_mrk = Table(displayName="Merknad", ref=f"A1:{get_column_letter(len(mrk_cols))}{len(merknad_df)+1}")
    tbl_mrk.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_mrk.add_table(tbl_mrk)
    for col in ws_mrk.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws_mrk.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

n_totals  = (combined["Rad"] == "TOTAL").sum()
n_element = (combined["Rad"] == "Element").sum()
print(f"Resultat skrevet til: {OUTPUT}")
print(f"\nAntall totalrader (poster): {n_totals}")
print(f"Antall elementrader:        {n_element}")
print(f"  Begge kilder:   {(combined.loc[combined['Rad']=='TOTAL', 'Kilde'] == 'Begge').sum()}")
print(f"  Kun i LARK:     {(combined.loc[combined['Rad']=='TOTAL', 'Kilde'] == 'Kun i LARK').sum()}")
