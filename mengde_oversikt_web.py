# -*- coding: utf-8 -*-
"""
Webversjon av Mengde Oversikt GUI — kjoer med:
    streamlit run mengde_oversikt_web.py
Del lenka fraa ngrok med kollegaer.
"""

from __future__ import annotations

import io
import os
import shutil
import sys
import tempfile
import traceback

import streamlit as st

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ── stdout-omdirigering (same mønster som tkinter-GUI) ────────────────────────

class _BytesLogStream(io.RawIOBase):
    """Bytes-lag — trengst av codecs.getwriter('utf-8')(sys.stdout.buffer)."""
    def __init__(self, parent: "_LogStream"):
        self._p = parent
    def write(self, b: bytes | bytearray) -> int:
        self._p.write(b.decode("utf-8", errors="replace"))
        return len(b)
    def flush(self) -> None:
        pass


class _LogStream(io.TextIOBase):
    """Samlar all stdout/stderr i ein StringIO; har .buffer for codecs.getwriter."""
    def __init__(self):
        self._buf = io.StringIO()
        self.buffer = _BytesLogStream(self)
    def write(self, s: str) -> int:
        self._buf.write(s)
        return len(s)
    def flush(self) -> None:
        pass
    def getvalue(self) -> str:
        return self._buf.getvalue()
sys.path.insert(0, _SCRIPT_DIR)

# ── Sideoppsett ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Mengde Oversikt",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Mengde Oversikt")
st.caption("Last opp filer, vel alternativ og trykk Kjør.")

# ── Initialisér session state ─────────────────────────────────────────────────
for key, val in [("log", ""), ("resultater", {}), ("kjoert", False)]:
    if key not in st.session_state:
        st.session_state[key] = val

# ── Filopplasting ─────────────────────────────────────────────────────────────
st.subheader("Filer")

col_csv, col_excel = st.columns([1, 1])

with col_csv:
    csv_filer = st.file_uploader(
        "CSV-filer (selection-basket)",
        type=["csv"],
        accept_multiple_files=True,
        help="Ein eller fleire selection-basket CSV-ar frå Solibri.",
    )

with col_excel:
    omega_fil = st.file_uploader(
        "Omega-fil (.xlsx)",
        type=["xlsx", "xls"],
    )
    isy_filer = st.file_uploader(
        "ISY Linker-filer (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
    apm_fil = st.file_uploader(
        "APM Målebrev (Maalebrev_samlet.xlsx) — valfri",
        type=["xlsx", "xls"],
        help=(
            "Trekk inn `Maalebrev_samlet.xlsx` for å få APM-mengde og "
            "APM-kilde-kolonner i hovuddokumentet. Utan denne vert dei "
            "tomme. Hyperlenkjer i resultatfila peikar til same mappe — "
            "legg Maalebrev_samlet.xlsx ved sidan av oversikt-fila for at "
            "lenkjene skal virke."
        ),
    )

# ── Alternativ ────────────────────────────────────────────────────────────────
st.subheader("Alternativ")
lag_rib_arm  = st.checkbox(
    "Lag ARM/RIB-rapport  (Mengdeuttak_Armering)",
    value=True,
)
fob_mengde = st.checkbox(
    "Bruk FOB_Mengde for RIB-volum  (i staden for PGF_Kalkyle)",
    value=False,
)

# ── Køyr-knapp ────────────────────────────────────────────────────────────────
klar = bool(csv_filer and omega_fil and isy_filer)
kjor = st.button("▶  Kjør", type="primary", disabled=not klar)

if not klar:
    manglar = []
    if not csv_filer:
        manglar.append("CSV-fil")
    if not omega_fil:
        manglar.append("Omega-fil")
    if not isy_filer:
        manglar.append("ISY Linker-fil")
    if manglar:
        st.info(f"Manglar: {', '.join(manglar)}")

# ── Analyse ───────────────────────────────────────────────────────────────────
if kjor and klar:
    st.session_state.kjoert    = False
    st.session_state.log       = ""
    st.session_state.resultater = {}

    tmp = tempfile.mkdtemp()
    log_buf = _LogStream()
    feil: str | None = None
    arm_rib_excels: list[str] = []
    oversikt: str | None = None

    try:
        # Lagre opplasta filer til temp-mappe
        csv_paths: list[str] = []
        for f in csv_filer:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as fp:
                fp.write(f.getbuffer())
            csv_paths.append(p)

        omega_path = os.path.join(tmp, omega_fil.name)
        with open(omega_path, "wb") as fp:
            fp.write(omega_fil.getbuffer())

        isy_paths: list[str] = []
        for f in isy_filer:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as fp:
                fp.write(f.getbuffer())
            isy_paths.append(p)

        # APM Målebrev — valfri. Tom streng → gracefully ingen APM-kolonne.
        apm_path: str = ""
        if apm_fil is not None:
            apm_path = os.path.join(tmp, apm_fil.name)
            with open(apm_path, "wb") as fp:
                fp.write(apm_fil.getbuffer())

        # Køyr analyse med stdout omdirigert til buffer
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = log_buf

        with st.spinner("Køyrer analyse — dette kan ta nokre minutt ..."):
            try:
                from mengde_oversikt import kjor_arm_rib_per_omraade, lag_oversikt

                if lag_rib_arm:
                    print("=" * 60)
                    print("STEG 1 — Mengdeuttak ARM/RIB")
                    print("=" * 60)
                    arm_rib_excels = kjor_arm_rib_per_omraade(
                        csv_path=csv_paths,
                        omega_path=omega_path,
                        fob_mengde_mode=fob_mengde,
                    )
                    for p in arm_rib_excels:
                        print(f"\n[OK] ARM/RIB: {p}")

                print("\n" + "=" * 60)
                print("STEG 2 — Mengde Oversikt")
                print("=" * 60)
                # Web-modus: hyperlenkjer skal vere relative (berre filnamn)
                # slik at dei virkar viss brukar legg begge filer i same mappe.
                apm_link = os.path.basename(apm_path) if apm_path else None
                oversikt = lag_oversikt(
                    csv_path=csv_paths,
                    omega_path=omega_path,
                    isy_path=isy_paths,
                    apm_excel=apm_path,
                    apm_link_target=apm_link,
                    arm_rib_excels=arm_rib_excels if arm_rib_excels else None,
                )
                print(f"\n[OK] Oversikt: {oversikt}")
                print("\nFerdig!")
            except Exception as exc:
                feil = f"[FEIL] {exc}\n{traceback.format_exc()}"
                print(feil)
            finally:
                sys.stdout = old_out
                sys.stderr = old_err

        # Les resultatfiler inn i minnet foer temp-mappe vert sletta
        resultater: dict[str, tuple[str, bytes]] = {}

        if oversikt and os.path.exists(oversikt):
            with open(oversikt, "rb") as f:
                resultater["oversikt"] = (os.path.basename(oversikt), f.read())

        for excel in arm_rib_excels:
            if os.path.exists(excel):
                name = os.path.basename(excel)
                with open(excel, "rb") as f:
                    resultater[name] = (name, f.read())

        # Tilby APM-fila for nedlasting òg — slik at brukar kan lagre han ved
        # sidan av oversikt-fila og få hyperlenkjene til å virke.
        if apm_path and os.path.exists(apm_path):
            with open(apm_path, "rb") as f:
                resultater["apm"] = (os.path.basename(apm_path), f.read())

        st.session_state.resultater = resultater
        st.session_state.log        = log_buf.getvalue()
        st.session_state.kjoert     = True

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

# ── Vis logg og resultat (fraa session state) ─────────────────────────────────
if st.session_state.kjoert:
    st.divider()

    with st.expander("Logg", expanded=True):
        st.code(st.session_state.log or "(tom)", language=None)

    resultater = st.session_state.resultater
    if not resultater:
        st.error("Analysen feila eller produserte ingen filer. Sjaa loggen.")
    else:
        st.subheader("Last ned resultat")
        MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        # Oversikt-knapp
        if "oversikt" in resultater:
            name, data = resultater["oversikt"]
            st.download_button(
                label="📥  Last ned Oversikt",
                data=data,
                file_name=name,
                mime=MIME,
            )
            if "apm" in resultater:
                st.caption(
                    "💡 APM-hyperlenkjene peikar til *Maalebrev_samlet.xlsx* "
                    "i same mappe — last ned begge filene under og legg dei "
                    "i same mappe for at lenkjene skal opne riktig ark."
                )

        # ARM/RIB- og APM-knapper
        for key, (name, data) in resultater.items():
            if key == "oversikt":
                continue
            if key == "apm":
                label = "📥  Last ned APM Målebrev (for hyperlenkjer)"
            elif "_LYA" in name:
                label = "📥  Last ned ARM/RIB LYA"
            elif "_VEK" in name:
                label = "📥  Last ned ARM/RIB VEK"
            else:
                label = "📥  Last ned ARM/RIB-rapport"
            st.download_button(label=label, data=data, file_name=name, mime=MIME)
