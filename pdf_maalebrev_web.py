# -*- coding: utf-8 -*-
"""
Webversjon av pdf_maalebrev_til_excel — kjoer med:
    streamlit run pdf_maalebrev_web.py
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
sys.path.insert(0, _SCRIPT_DIR)


# ── stdout-omdirigering (same mønster som mengde_oversikt_web) ────────────────

class _BytesLogStream(io.RawIOBase):
    def __init__(self, parent: "_LogStream"):
        self._p = parent
    def write(self, b: bytes | bytearray) -> int:
        self._p.write(b.decode("utf-8", errors="replace"))
        return len(b)
    def flush(self) -> None:
        pass


class _LogStream(io.TextIOBase):
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


# ── Sideoppsett ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Målebrev til Excel",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Målebrev PDF → Excel")
st.caption("Last opp APM-PDF-filer og Omega-fil, trykk Kjør og last ned resultatet.")

# ── Initialisér session state ─────────────────────────────────────────────────
for key, val in [("log", ""), ("resultat", None), ("resultat_namn", ""), ("kjoert", False)]:
    if key not in st.session_state:
        st.session_state[key] = val

# ── Filveljing ────────────────────────────────────────────────────────────────
st.subheader("Filer")

fane_lokal, fane_last_opp = st.tabs(["📁  Lokal mappe", "⬆️  Last opp filer"])

pdf_filer   = []       # brukt i opplastings-modus
pdf_mappe_lokal = ""   # brukt i lokal-modus
lokal_modus = False

with fane_lokal:
    st.caption("Skriv inn stien til mappa med PDF-filene på denne maskinen.")
    sti_input = st.text_input(
        "Sti til PDF-mappe",
        placeholder=r"C:\Users\...\Målebrev\APM-PDFer",
    )
    if sti_input:
        sti_input = sti_input.strip().strip('"')
        if not os.path.isdir(sti_input):
            st.error(f"Mappa finst ikkje: {sti_input}")
        else:
            pdfer_i_mappe = [f for f in os.listdir(sti_input) if f.lower().endswith(".pdf")]
            if not pdfer_i_mappe:
                st.warning("Ingen PDF-filer funne i mappa.")
            else:
                st.success(f"{len(pdfer_i_mappe)} PDF-filer funne.")
                pdf_mappe_lokal = sti_input
                lokal_modus = True

with fane_last_opp:
    st.caption("Last opp enkeltfiler frå din PC (fungerer også for eksterne brukarar via lenke).")
    pdf_filer = st.file_uploader(
        "APM PDF-filer",
        type=["pdf"],
        accept_multiple_files=True,
    )

# Omega-fil (felles for begge modusane)
st.divider()
omega_fil = st.file_uploader(
    "Omega-fil (.xlsx)  — valfri",
    type=["xlsx", "xls"],
    help="BOQNT-arket vert brukt til å slå opp kontraktsmengder og rapporterte mengder.",
)

# ── Alternativ ────────────────────────────────────────────────────────────────
st.subheader("Alternativ")
test_modus = st.checkbox(
    "Testmodus  (behandlar berre ~15 % av filene — rask sjekk)",
    value=False,
)

# ── Køyr-knapp ────────────────────────────────────────────────────────────────
klar = bool(lokal_modus and pdf_mappe_lokal) or bool(pdf_filer)
kjor = st.button("▶  Kjør", type="primary", disabled=not klar)

if not klar:
    st.info("Vel ei lokal PDF-mappe eller last opp PDF-filer for å starte.")

# ── Analyse ───────────────────────────────────────────────────────────────────
if kjor and klar:
    st.session_state.kjoert      = False
    st.session_state.log         = ""
    st.session_state.resultat    = None
    st.session_state.resultat_namn = ""

    tmp = tempfile.mkdtemp()
    log_buf = _LogStream()
    feil: str | None = None
    output_fil: str | None = None

    try:
        if lokal_modus:
            # Lokal mappe — les direkte frå filsystemet
            pdf_mappe = pdf_mappe_lokal
            n_pdfer = len([f for f in os.listdir(pdf_mappe) if f.lower().endswith(".pdf")])
        else:
            # Opplastingsmodus — lagre til temp-mappe
            pdf_mappe = os.path.join(tmp, "pdfer")
            os.makedirs(pdf_mappe)
            for f in pdf_filer:
                p = os.path.join(pdf_mappe, f.name)
                with open(p, "wb") as fp:
                    fp.write(f.getbuffer())
            n_pdfer = len(pdf_filer)

        # Lagre Omega-fil viss oppgjeven
        omega_path: str | None = None
        if omega_fil:
            omega_path = os.path.join(tmp, omega_fil.name)
            with open(omega_path, "wb") as fp:
                fp.write(omega_fil.getbuffer())

        output_path = os.path.join(tmp, "Maalebrev_samlet.xlsx")

        # Køyr analyse
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = log_buf

        with st.spinner(f"Behandlar {n_pdfer} PDF-filer — dette kan ta nokre minutt ..."):
            try:
                from pdf_maalebrev_til_excel import extract_folder
                output_fil = extract_folder(
                    mappe=pdf_mappe,
                    output_path=output_path,
                    omega_path=omega_path,
                    test_modus=test_modus,
                )
            except Exception as exc:
                feil = f"[FEIL] {exc}\n{traceback.format_exc()}"
                print(feil)
            finally:
                sys.stdout = old_out
                sys.stderr = old_err

        # Les resultatfil inn i minnet
        if output_fil and os.path.exists(output_fil):
            with open(output_fil, "rb") as f:
                st.session_state.resultat      = f.read()
                st.session_state.resultat_namn = os.path.basename(output_fil)
        elif not feil:
            feil = "[FEIL] extract_folder returnerte ingen output-fil."

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    st.session_state.log    = log_buf.getvalue() + (feil or "")
    st.session_state.kjoert = True

# ── Vis logg og resultat ──────────────────────────────────────────────────────
if st.session_state.kjoert:
    st.divider()

    with st.expander("Logg", expanded=True):
        st.code(st.session_state.log or "(tom)", language=None)

    if st.session_state.resultat:
        st.success("Ferdig!")
        st.subheader("Last ned resultat")
        st.download_button(
            label="📥  Last ned Maalebrev_samlet.xlsx",
            data=st.session_state.resultat,
            file_name=st.session_state.resultat_namn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("Analysen feila. Sjå loggen.")
