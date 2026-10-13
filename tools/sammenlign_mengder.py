"""
Sammenligner mengdelistepost/kontraktsmengde mellom:
  - Omega_K7_Mars.xlsx  (BOQNT-ark, kolonne 'Navn (Prosesskode)' og 'Kontraktsmengde')
  - LARK_planter.xlsx   (Tabell1-ark, kolonne 'FOB_Info.FOB_Mengdelistepost' og 'FOB_Info.FOB_Mengde')

Semikolon-separerte celler i LARK-filen ekspanderes til separate poster.
"""

import pandas as pd
from collections import defaultdict

OMEGA_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang\Omega_K7_Mars.xlsx"
)

LARK_PATH = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\LARK_planter.xlsx"
)

# ── Les Omega ──────────────────────────────────────────────────────────────────
omega_raw = pd.read_excel(OMEGA_PATH, sheet_name="BOQNT", header=2)
# Relevante kolonner
omega = omega_raw[["Navn (Prosesskode)", "Enhet", "Kontraktsmengde"]].copy()
omega.columns = ["post", "enhet_omega", "kontraktsmengde"]
omega["post"] = omega["post"].astype(str).str.strip()
omega = omega[omega["post"].notna() & (omega["post"] != "nan")]

# Summer kontraktsmengde per post (kan forekomme flere ganger i Omega)
omega_sum = (
    omega.groupby("post", sort=False)
    .agg(
        kontraktsmengde=("kontraktsmengde", "sum"),
        enhet_omega=("enhet_omega", "first"),
    )
    .reset_index()
)

# ── Les LARK ──────────────────────────────────────────────────────────────────
lark_raw = pd.read_excel(LARK_PATH, sheet_name="Tabell1", header=0)
lark_raw.columns = [c.strip() for c in lark_raw.columns]

# Expand semicolon-separated rows
rows = []
for _, r in lark_raw.iterrows():
    posts_raw = str(r["FOB_Info.FOB_Mengdelistepost"]) if pd.notna(r["FOB_Info.FOB_Mengdelistepost"]) else ""
    mengder_raw = str(r["FOB_Info.FOB_Mengde"]) if pd.notna(r["FOB_Info.FOB_Mengde"]) else ""
    enheter_raw = str(r["FOB_Info.FOB_Mengdeenhet"]) if pd.notna(r["FOB_Info.FOB_Mengdeenhet"]) else ""

    posts = [p.strip() for p in posts_raw.split(";")]
    mengder = [m.strip() for m in mengder_raw.split(";")]
    enheter = [e.strip() for e in enheter_raw.split(";")]

    for i, post in enumerate(posts):
        if not post or post == "nan":
            continue
        try:
            mengde = float(mengder[i]) if i < len(mengder) else 0.0
        except ValueError:
            mengde = 0.0
        enhet = enheter[i] if i < len(enheter) else ""
        rows.append({"post": post, "mengde_lark": mengde, "enhet_lark": enhet})

lark_exp = pd.DataFrame(rows)

# Summer per post
lark_sum = (
    lark_exp.groupby("post", sort=False)
    .agg(
        mengde_lark=("mengde_lark", "sum"),
        enhet_lark=("enhet_lark", "first"),
    )
    .reset_index()
)

# ── Slå sammen ────────────────────────────────────────────────────────────────
merged = pd.merge(lark_sum, omega_sum, on="post", how="outer", indicator=True)
merged["_merge"] = merged["_merge"].map(
    {"left_only": "Kun i LARK", "right_only": "Kun i Omega", "both": "Begge"}
)

# Avrund
merged["mengde_lark"] = merged["mengde_lark"].round(2)
merged["kontraktsmengde"] = merged["kontraktsmengde"].round(2)

# Avvik
merged["avvik"] = (merged["mengde_lark"] - merged["kontraktsmengde"]).round(2)

merged = merged.rename(columns={
    "post": "Mengdelistepost",
    "enhet_lark": "Enhet (LARK)",
    "enhet_omega": "Enhet (Omega)",
    "mengde_lark": "Mengde LARK",
    "kontraktsmengde": "Kontraktsmengde (Omega)",
    "avvik": "Avvik (LARK - Omega)",
    "_merge": "Kilde",
})

col_order = [
    "Mengdelistepost",
    "Enhet (LARK)",
    "Enhet (Omega)",
    "Mengde LARK",
    "Kontraktsmengde (Omega)",
    "Avvik (LARK - Omega)",
    "Kilde",
]
merged = merged[col_order].sort_values("Mengdelistepost")
merged = merged[merged["Kilde"] != "Kun i Omega"]

# ── Skriv ut ──────────────────────────────────────────────────────────────────
import os
OUTPUT = os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "excel"), "sammenlign_mengder_resultat.xlsx")
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    merged.to_excel(writer, sheet_name="Sammenligning", index=False)

    # Enkel farging: rød = kun i én kilde, gul = avvik, grønn = ok
    from openpyxl.styles import PatternFill, Font
    ws = writer.sheets["Sammenligning"]

    # Juster kolonnebredder
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    red   = PatternFill("solid", fgColor="FF9999")
    yellow = PatternFill("solid", fgColor="FFFF99")
    green  = PatternFill("solid", fgColor="CCFFCC")

    header_row = 1
    kilde_col = col_order.index("Kilde") + 1
    avvik_col  = col_order.index("Avvik (LARK - Omega)") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        kilde = row[kilde_col - 1].value
        avvik_val = row[avvik_col - 1].value

        if kilde in ("Kun i LARK", "Kun i Omega"):
            fill = red
        elif avvik_val is not None and abs(avvik_val) > 0.01:
            fill = yellow
        else:
            fill = green

        for cell in row:
            cell.fill = fill

print(f"Resultat skrevet til: {OUTPUT}")
print(f"\nAntall poster totalt: {len(merged)}")
print(f"  Begge kilder:   {(merged['Kilde'] == 'Begge').sum()}")
print(f"  Kun i LARK:     {(merged['Kilde'] == 'Kun i LARK').sum()}")
print(f"  Kun i Omega:    {(merged['Kilde'] == 'Kun i Omega').sum()}")
print(f"\nPoster med avvik (begge kilder): {((merged['Kilde'] == 'Begge') & (merged['Avvik (LARK - Omega)'].abs() > 0.01)).sum()}")
