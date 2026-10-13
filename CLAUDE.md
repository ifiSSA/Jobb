# CLAUDE.md — `ssa/ifi`

Mengde- og prognoseverktøy for Scan Survey (Fornebubanen, K5A/K7/K10 — Lysaker/Vækerø). Skripta les BIM-uttrekk (selection-basket CSV), Omega-eksport (dsBOQ/BOQNT), ISY Linker (NS3420/prosesskode) og APM-målebrev (PDF), og skriv Excel-rapportar til FOB.

Mappa ligg i monorepoet `C:\Users\ifi\Documents\ssa` (git-rota), med venv i `..\.venv`.

## Arbeidsreglar (viktig)

1. **Les aldri datafiler utan beskjed.** Ikkje opne, les, parse eller inspiser Excel-ark (`.xlsx`), PDF-ar, CSV-ar eller andre inn-/utdatafiler som koden brukar (heller ikkje via Python-snuttar, pandas, openpyxl eller pdfplumber) med mindre brukaren eksplisitt ber om det. Dette gjeld både filer i mappa (alt i `excel/`, `output/`) og stiane i `*_innstillingar.json` / `*_innstillinger.json`.
2. **Brukaren gjev formatet.** Brukaren oppgjev som standard kolonnenamn, arknamn, oppsett og døme på korleis dataa ser ut. Skriv koden ut frå det brukaren har skildra — ikkje gå og sjekk dei faktiske filene for å verifisere. Manglar noko (t.d. eit kolonnenamn), spør brukaren i staden for å lese fila.
3. Kodelesing (`.py`, `.bat`, `.json`-innstillingar som konfig) er greitt.

## Kode-konvensjonar

- Norsk namngjeving og kommentarar (blanding av nynorsk og bokmål — følg stilen i fila du endrar): `les_*`, `bygg_*`, `skriv_*`, `lag_*`, `kjor_*`, `finn_*`. Private hjelparar har `_`-prefiks.
- `# -*- coding: utf-8 -*-`, `from __future__ import annotations`, modul-docstring som forklarer format, reglar og `Bruk:`.
- Seksjonsskiljarar som `# ── Tittel ───…` eller `# ====…`.
- Skriv alltid Excel via `lagring.vent_til_skrivbar(utfil)` rett før `wb.save()` / `ExcelWriter` (fila kan vere open i Excel).
- Excel-filer som vert lagra i kodemappa skal i `excel/` — bruk `lagring.excel_mappe()`. Mappa (og `*.xlsx`) står i `.gitignore`. Ingen hardkoda stiar til kodemappa: rekn ut frå `__file__`.
- Utfiler får tidsstempel: `<namn>_YYYYMMDD_HHMMSS.xlsx`. Excel-formatering med openpyxl (tabellar, fargar, kolonnebreidder, hyperlenker, kommentarar).
- Felles lesing av Omega og CSV ligg i `base_lesning.py` — gjenbruk det i nye analysar.
- Brukaren tek manuelle backupar som `<fil>.py.bak_YYYYMMDD` — ikkje rediger eller slett dei.
- Ingen testsuite for denne mappa. Verifiser med `python -c "import modul"` / `py_compile`, ikkje ved å køyre mot ekte datafiler utan løyve.
- Konsollutskrift må tole æøå (Windows cp1252) — sjå korleis modulane set UTF-8 på stdout.

## Dataflyt

```
selection-basket CSV ──► csv_til_armrib.kjor_mengdeuttak ──► Mengdeuttak_Armering.py (exec)
   (ARM + RIB)                                                 └─► resultat_mengdeuttak_*_LM.xlsx ("Sammendrag")
                                                                          │
Omega (dsBOQ/BOQNT) ─┐                                                    ▼
ISY Linker (NS3420) ─┼──────────────────────────────► mengde_oversikt.lag_oversikt ──► mengde_oversikt_*.xlsx
APM-PDF-ar ──► pdf_maalebrev_til_excel.extract_folder ─┘                                   │
Kabellister (PF-*-LC-*.xlsx) ──► kabel_prognose ─────────► prognose_ark.lag_prognose ◄─────┘
                                                              └─► prognose_*.xlsx (+ Omega-opplastingsmal)
```

## Filer

**Kjerne**
- [base_lesning.py](base_lesning.py) — `les_omega`, `bygg_omega_koder`, `finn_omega_ark`, `les_csv`, `split_arm_rib`, `les_og_split_csv`. ARM-modell = Path med `_RIB_ARM_`, RIB = `_RIB_` utan `_RIB_ARM_`.
- [Mengdeuttak_Armering.py](Mengdeuttak_Armering.py) — stort script (ingen `main`), køyrer på modulnivå. Inndata via globale variablar (`excel_file_armering`, `excel_file_RIB`, `excel_file_OMEGA`, …) som vert injiserte med `exec()` frå `csv_til_armrib.py` eller `Mengdeuttak_GUI.py`. Endringar her må halde på dette globals-kontraktet.
- [csv_til_armrib.py](csv_til_armrib.py) — `kjor_mengdeuttak(csv_path, omega_path, …)`: LP-filter (`-LP` i path), FOB_Status S5 = hovudrapport, S4 på eige ark, omdøyper kolonner (`/`→`.`, `ObjectType`→`Type`) og monkey-patchar `pd.read_excel` før exec.
- [mengde_oversikt.py](mengde_oversikt.py) — `lag_oversikt(...)`: ark `I_modell`, `Ikke_i_modell`, `ISY_Omega_avvik`, `BIM_utan_Omega`, betong/forskaling.
- [prognose_ark.py](prognose_ark.py) — `lag_prognose(...)`: FOB-rapport (Forside, NS3420, Prosesskoden, Revisjonsbeskrivelser, Alle poster, Kontroll) + Omega-opplasting. Prognoseregelen står i modul-docstringen.
- [kabel_prognose.py](kabel_prognose.py) — kabelmengder frå kabellister kopla mot NS3420 på kabeltype (`stk` = tal taggar, `m` = lengde).
- [pdf_maalebrev_til_excel.py](pdf_maalebrev_til_excel.py) — `extract_folder(mappe, ...)`: les APM-målebrev med pdfplumber (parallelt), MD5-dedup, `(N)`-suffiks = utgått, mykje prosesskode-reparasjon.
- [pdf_manglende_poster_til_excel.py](pdf_manglende_poster_til_excel.py) — rapporten «Poster i modeller som mangler i kontrakt…»; kolonnar delast på x-koordinat (ikkje `extract_tables`). CLI + tkinter-GUI.
- [lagring.py](lagring.py) — `vent_til_skrivbar`, `excel_mappe()`.
- [isy_txt_til_xlsx.py](isy_txt_til_xlsx.py) — ISY Linker `.txt` → `.xlsx` med kolonnene `les_isy` ventar.
- [sjekk_volum_kjede.py](sjekk_volum_kjede.py) — sporar ein post gjennom CSV → Sammendrag → oversikt → prognose.

**Brukargrensesnitt**
- [mengde_oversikt_gui.py](mengde_oversikt_gui.py) — hovud-GUI (tkinter, trådar, stdout til widget). Lagrar val i `mengde_oversikt_gui_innstillingar.json`, sorterer utfiler i månadsmapper. Start: `Mengde Oversikt.bat`.
- [mengde_oversikt_web.py](mengde_oversikt_web.py) / [pdf_maalebrev_web.py](pdf_maalebrev_web.py) — Streamlit (port 8501 / 8503), delast via ngrok (`Start Web.bat`, `Start PDF Web.bat`, `Start PDF Lokal.bat`).
- [Mengdeuttak_GUI.py](Mengdeuttak_GUI.py) — eldre GUI som `exec`-ar `Mengdeuttak_Armering.py` (`Start_Mengdeuttak_GUI.bat`).

**Anna**
- [tools/](tools/) — eldre/eingongs-analysar (Omfar, SKY-samanlikning, LARK, …).
- [verktoy_andre/](verktoy_andre/) — urelaterte målingsverktøy (bolter, vinkelavvik); `build/`/`dist/` er PyInstaller-output derifrå.
- [_archive/](_archive/) — gamle debug-/testscript. Ikkje bruk som referanse.
- [excel/](excel/) — alle Excel-filer (resultat, oversikter, grunnlag). Ikkje i git (sjå regel 1).
- [feil_poster.py](feil_poster.py) — feillista `feil_mengdelisteposter_*.xlsx` (éi rad per feil per post).

## Avhengigheiter

pandas, numpy, openpyxl, pdfplumber, streamlit (web), tkinter (GUI). Kjør med `python` i PATH eller `C:\Users\ifi\Documents\ssa\.venv\Scripts\python.exe`.
