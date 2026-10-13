# -*- coding: utf-8 -*-
"""
Spor hvor armeringsvekt forsvinner mellom rå modell og slutt-rapport.

Reproduserer flyten i Mengdeuttak_Armering.py steg for steg, og logger
hvor mye vekt som faller bort på hvert filter/aggregering.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd

from csv_til_armrib import les_og_split_csv

CSV_PATH   = r"C:\Users\ifi\Downloads\selection-basket (30).csv"
OMEGA_PATH = r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\APM FOB Gjennomgang\Omega_K7_April.xlsx"

KEY     = "FOB_Info.FOB_Produksjonsenhet"
WEIGHT  = "PGF_RIB_Info.Vekt"
MENG    = "FOB_Info.FOB_Mengdelistepost"
VOL     = "PGF_Kalkyle.PGF_Volum"
TYPE    = "Type"
FAGKAP  = "PGF_Info.PGF_Fagkapittel"

ALLOWED_FAGKAP = {"05 Betongarbeider", "06 Betongkonstruksjoner, prefab"}


def vekt(df):
    return df[WEIGHT].fillna(0).astype(float).sum()


def hdr(tittel):
    print("\n" + "=" * 76)
    print(tittel)
    print("=" * 76)


df_arm, df_rib = les_og_split_csv(CSV_PATH)

start_vekt = vekt(df_arm)
start_n = len(df_arm)
print(f"\n[START] ARM: {start_n:,} jern, vekt {start_vekt:.2f} kg")

# ------------------------------------------------------------------
# STEG 1 – Filtrer RIB-rader (slik Mengdeuttak_Armering.py gjør)
# ------------------------------------------------------------------
hdr("STEG 1: RIB-filtrering — påvirker hvilke PE-er ARM kan kobles til")
n_rib_start = len(df_rib)

# 1a: fjern Fuge
n_fuge = df_rib[TYPE].astype(str).str.contains("fuge", case=False, na=False).sum()
df_rib = df_rib[~df_rib[TYPE].astype(str).str.contains("fuge", case=False, na=False)].copy()
print(f"  Fjerner Fuge i Type: -{n_fuge} rader -> {len(df_rib):,}")

# 1b: behold kun fagkapittel 05/06
n_feil_fag = (~df_rib[FAGKAP].isin(ALLOWED_FAGKAP)).sum()
df_rib = df_rib[df_rib[FAGKAP].isin(ALLOWED_FAGKAP)].copy()
print(f"  Beholder fagkap 05/06: -{n_feil_fag} rader -> {len(df_rib):,}")

# 1c: behold kun rader med volum > 0
n_uten_vol = (df_rib[VOL].isna() | (df_rib[VOL] <= 0)).sum()
df_rib_med_vol = df_rib[df_rib[VOL].notna() & (df_rib[VOL] > 0)].copy()
print(f"  Beholder volum > 0: -{n_uten_vol} rader -> {len(df_rib_med_vol):,}")

# Hvilke PE-er finnes i RIB etter filtrering?
pe_i_rib_str = df_rib_med_vol[KEY].dropna().astype(str)
pe_i_rib = set()
for v in pe_i_rib_str:
    for p in v.split(","):
        p = p.strip()
        if p and p.lower() != "nan":
            pe_i_rib.add(p.lower())
print(f"  Unike PE-er i filtrert RIB: {len(pe_i_rib):,}")

# ------------------------------------------------------------------
# STEG 2 – ARM-eksplodering på PE
# ------------------------------------------------------------------
hdr("STEG 2: ARM-eksplodering på PE")

df_arm_x = df_arm.copy()
print(f"  Før eksplodering: {len(df_arm_x):,} jern, vekt {vekt(df_arm_x):.2f} kg")

# Hvor mange jern har tom/NaN PE?
tom_pe_mask = df_arm_x[KEY].isna() | (df_arm_x[KEY].astype(str).str.strip() == "")
n_tom = tom_pe_mask.sum()
v_tom = df_arm_x.loc[tom_pe_mask, WEIGHT].fillna(0).sum()
print(f"  Jern uten PE:        {n_tom:,} jern, vekt {v_tom:.2f} kg")

# Eksplodér (samme som original)
df_arm_x[KEY] = df_arm_x[KEY].astype(str).fillna("").str.split(",")
df_arm_x = df_arm_x.explode(KEY)
df_arm_x[KEY] = df_arm_x[KEY].str.strip()
# Filtrer tomme PE-er
n_for_drop = len(df_arm_x)
v_for_drop = vekt(df_arm_x)
df_arm_x = df_arm_x[df_arm_x[KEY] != ""]
n_etter_drop = len(df_arm_x)
v_etter_drop = vekt(df_arm_x)
print(f"  Etter eksplodering:  {n_etter_drop:,} rader, vekt {v_etter_drop:.2f} kg")
print(f"  Tapt på tom PE:      {n_for_drop - n_etter_drop:,} rader, vekt {v_for_drop - v_etter_drop:.2f} kg")

# ------------------------------------------------------------------
# STEG 3 – Hvilke ARM-PE-er finnes i RIB?
# ------------------------------------------------------------------
hdr("STEG 3: ARM-PE som ikke har match i filtrert RIB")

df_arm_x["_match"] = df_arm_x[KEY].str.lower()
arm_match = df_arm_x["_match"]

har_match = arm_match.isin(pe_i_rib)
n_uten_match = (~har_match).sum()
v_uten_match = df_arm_x.loc[~har_match, WEIGHT].fillna(0).sum()
print(f"  ARM-rader UTEN matchende RIB-PE: {n_uten_match:,}, vekt {v_uten_match:.2f} kg")
print(f"  ARM-rader MED match:             {har_match.sum():,}, vekt {df_arm_x.loc[har_match, WEIGHT].fillna(0).sum():.2f} kg")

if n_uten_match:
    pe_uten_match = (df_arm_x.loc[~har_match]
                     .groupby(KEY)[WEIGHT]
                     .agg(lambda s: s.fillna(0).astype(float).sum())
                     .sort_values(ascending=False))
    print("\n  Topp 15 PE-er uten RIB-match:")
    for pe, w in pe_uten_match.head(15).items():
        print(f"    {pe!r:40s}  {w:>10.2f} kg")

# ------------------------------------------------------------------
# STEG 4 – RIB-rader uten Mengdelistepost og uten volum
# ------------------------------------------------------------------
hdr("STEG 4: RIB-rader som faller bort fordi de mangler Mengdelistepost / volum")

# Final-bygging: original beholder rib_rows for_final = rib_long med volum > 0,
# men forventer Mengdelistepost. Rader uten Mengdelistepost gir tom/NaN —
# disse blir likevel beholdt som rader i 'final' så lenge volum finnes.
n_meng_na = df_rib_med_vol[MENG].isna().sum()
print(f"  RIB-rader uten Mengdelistepost: {n_meng_na:,} av {len(df_rib_med_vol):,}")
# Hvor mye armeringsvekt sitter på PE-er som KUN har RIB-rader uten mengdelistepost?
pe_kun_uten_meng = set()
df_rib_eks = df_rib_med_vol.copy()
df_rib_eks[KEY] = df_rib_eks[KEY].astype(str).fillna("").str.split(",")
df_rib_eks = df_rib_eks.explode(KEY)
df_rib_eks[KEY] = df_rib_eks[KEY].str.strip()
df_rib_eks["_match"] = df_rib_eks[KEY].str.lower()
pe_med_meng = set(df_rib_eks.loc[df_rib_eks[MENG].notna(), "_match"].unique())
pe_alle_rib = set(df_rib_eks["_match"].unique())
pe_kun_uten_meng = pe_alle_rib - pe_med_meng
v_pe_kun_uten = df_arm_x.loc[df_arm_x["_match"].isin(pe_kun_uten_meng), WEIGHT].fillna(0).sum()
print(f"  PE-er der ALLE RIB-rader mangler mengdelistepost: {len(pe_kun_uten_meng):,}")
print(f"  Armeringsvekt på slike PE-er: {v_pe_kun_uten:.2f} kg")

if pe_kun_uten_meng:
    print("\n  Detaljer per PE (kandidat for 'forsvunnet' vekt):")
    for pe_low in sorted(pe_kun_uten_meng):
        # Hent original PE-streng (case)
        pe_orig = df_rib_eks.loc[df_rib_eks["_match"] == pe_low, KEY].iloc[0]
        # Vekt fra ARM
        arm_w = df_arm_x.loc[df_arm_x["_match"] == pe_low, WEIGHT].fillna(0).sum()
        arm_n = (df_arm_x["_match"] == pe_low).sum()
        # Volum og info fra RIB
        rib_rows = df_rib_eks[df_rib_eks["_match"] == pe_low]
        vol_sum = rib_rows[VOL].fillna(0).sum()
        types = rib_rows[TYPE].dropna().unique()
        fagk = rib_rows[FAGKAP].dropna().unique()
        print(f"\n    PE: {pe_orig}")
        print(f"      ARM: {arm_n:,} jern, vekt {arm_w:.2f} kg")
        print(f"      RIB: {len(rib_rows):,} rader, volum {vol_sum:.3f} m³, type(r) {list(types)[:3]}, fagkap {list(fagk)}")
        print(f"      Mengdelistepost: {rib_rows[MENG].dropna().unique().tolist() or 'INGEN'}")

# ------------------------------------------------------------------
# STEG 5 – Sluttfiltrering i Mengdeuttak_Armering: betong-krav per gruppe
# ------------------------------------------------------------------
hdr("STEG 5: Sluttfiltrering 'kun betong' – `groups_med_betong`")

# Mengdeuttak_Armering.py linje ~744:
#   if not groups_med_betong: continue
# Det filtrerer bort PE-rader der mengdelistepost ikke har match som betong-post i Omega
# OG fagkap heller ikke er betong. Dette tapet er det vanskelig å beregne uten
# Omega-koblingen, men vi kan estimere ved fagkap.
print("  Dette steget krever Omega-data — sjekkes via diff i hovedflyten.")

# ------------------------------------------------------------------
# OPPSUMMERING
# ------------------------------------------------------------------
hdr("OPPSUMMERING")
slutt_vekt = 713927.40  # fra hovedkjøringen (logget i forrige run)
print(f"  Rå ARM:                       {start_vekt:>12.2f} kg")
print(f"  Tapt på tom PE:               {(start_vekt - v_etter_drop):>12.2f} kg")
print(f"  Tapt på manglende RIB-match:  {v_uten_match:>12.2f} kg")
print(f"  Slutt-rapport (logget):       {slutt_vekt:>12.2f} kg")
diff = start_vekt - slutt_vekt
print(f"  Total differanse:             {diff:>12.2f} kg ({diff/start_vekt*100:.2f}%)")
forklart = (start_vekt - v_etter_drop) + v_uten_match
print(f"  Forklart av tom PE + RIB-mismatch: {forklart:>7.2f} kg")
uforklart = diff - forklart
print(f"  Uforklart rest:               {uforklart:>12.2f} kg")
