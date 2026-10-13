# -*- coding: utf-8 -*-
"""
GUI for å kjøre mengde_oversikt og/eller Mengdeuttak_Armering.
Start med:  python mengde_oversikt_gui.py
"""

from __future__ import annotations

import os
import sys
import io
import json
import queue
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── Månadsmappe ───────────────────────────────────────────────────────────────
# Alle resultatfilene frå ei køyring skal samlast i mappa for den månaden
# køyringa gjeld. Namnet kjem frå Periode-feltet, elles frå dagens dato.
_NOR_MAANED = (
    "Januar", "Februar", "Mars", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Desember",
)
_UGYLDIGE_MAPPETEIKN = '<>:"/\\|?*'

# ── Standardstiar ─────────────────────────────────────────────────────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

# Hugsar siste filval mellom køyringar.
_INNSTILLINGAR_FIL = os.path.join(_SCRIPT_DIR, "mengde_oversikt_gui_innstillingar.json")

_DEFAULT_CSV = r"C:\Users\ifi\Downloads\selection-basket (34).csv"
_DEFAULT_OMEGA = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang\Omega_K7_April.xlsx"
)
_DEFAULT_ISY = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\Mengdelistpost NS3420 isylinker.xlsx"
)
# Rota for resultatfilene. Månadsmappa vert laga under denne.
# Tomt felt = mappa Omega-fila ligg i (slik det var før).
_DEFAULT_UTMAPPE = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\Målebrev\Prognose beregning"
)
_DEFAULT_APM_MAPPE = (
    r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
    r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon"
    r"\APM FOB Gjennomgang"
)
_MAPPE_FT = [("Alle filer", "*.*")]

_CSV_FT    = [("CSV-filer", "*.csv"), ("Alle filer", "*.*")]
_EXCEL_FT  = [("Excel-filer", "*.xlsx *.xls"), ("Alle filer", "*.*")]

# ── Stdout-omdirigering til tekst-widget ──────────────────────────────────────

class _BytesWidgetStream(io.RawIOBase):
    """Bytes-lag under _WidgetStream — trengst av codecs.getwriter."""

    def __init__(self, text_stream: "_WidgetStream"):
        self._t = text_stream

    def write(self, b: bytes | bytearray) -> int:  # type: ignore[override]
        self._t.write(b.decode("utf-8", errors="replace"))
        return len(b)

    def flush(self) -> None:
        pass


class _WidgetStream(io.TextIOBase):
    """Sender write()-kall til ein tkinter Text-widget.

    Analysen køyrer i ein arbeidstråd, medan Tk berre toler å bli rørt frå
    hovudtråden. Tidlegare gjekk kvart einaste write() rett i
    widget.after(...) — eit Tk-kall frå feil tråd, og eitt planlagt kall,
    ei innsetjing og ei rulling per tekstbit.

    No legg arbeidstråden teksten i ein kø utan å røre Tk i det heile.
    Hovudtråden tømmer køen med jamne mellomrom og set inn heile bolken i
    éin operasjon.
    """

    INTERVALL_MS = 120        # kor ofte hovudtråden tømmer køen
    MAKS_LINJER = 5000        # eldre linjer vert klipte bort

    def __init__(self, widget: tk.Text):
        self._w = widget
        self._ko: queue.SimpleQueue = queue.SimpleQueue()
        self._aktiv = True
        self.buffer = _BytesWidgetStream(self)   # for codecs.getwriter
        self._w.after(self.INTERVALL_MS, self._tom_ko)

    def write(self, s: str) -> int:
        self._ko.put(s)                # trådsikkert, ingen Tk-kall her
        return len(s)

    def _tom_ko(self) -> None:
        bitar = []
        try:
            while True:
                bitar.append(self._ko.get_nowait())
        except queue.Empty:
            pass

        if bitar:
            self._w.configure(state="normal")
            self._w.insert("end", "".join(bitar))
            linjer = int(self._w.index("end-1c").split(".")[0])
            if linjer > self.MAKS_LINJER:
                self._w.delete("1.0", f"{linjer - self.MAKS_LINJER}.0")
            self._w.see("end")
            self._w.configure(state="disabled")

        if self._aktiv:
            self._w.after(self.INTERVALL_MS, self._tom_ko)

    def stopp(self) -> None:
        """Tøm køen ein siste gong og slutt å planlegge nye tømmingar."""
        self._aktiv = False
        self._tom_ko()

    def flush(self) -> None:
        pass


# ── Hovud-GUI ────────────────────────────────────────────────────────────────

class MengdeApp(tk.Tk):
    PAD = 8

    def __init__(self):
        super().__init__()
        self.title("Mengde Oversikt")
        self.resizable(True, False)
        self.minsize(700, 0)

        # Last lagra val frå førre køyring (fall tilbake til standardar)
        lagra = self._les_innstillingar()

        self._csv_stiar: list[str] = lagra.get(
            "csv_stiar", [_DEFAULT_CSV] if os.path.exists(_DEFAULT_CSV) else []
        )
        self._isy_stiar: list[str] = lagra.get(
            "isy_stiar", [_DEFAULT_ISY] if os.path.exists(_DEFAULT_ISY) else []
        )
        self._omega_var     = tk.StringVar(value=lagra.get("omega", _DEFAULT_OMEGA))
        self._apm_mappe     = tk.StringVar(value=lagra.get("apm_mappe", _DEFAULT_APM_MAPPE))
        self._rib_arm       = tk.BooleanVar(value=lagra.get("rib_arm", True))
        self._fob_mengde    = tk.BooleanVar(value=lagra.get("fob_mengde", False))
        self._hopp_over_apm = tk.BooleanVar(value=lagra.get("hopp_over_apm", False))
        self._lag_prognose  = tk.BooleanVar(value=lagra.get("lag_prognose", False))
        self._periode       = tk.StringVar(value=lagra.get("periode", ""))
        self._omega_overstyrer = tk.BooleanVar(value=lagra.get("omega_overstyrer", True))
        self._omega_mal     = tk.StringVar(value=lagra.get("omega_mal", ""))
        self._utmappe_rot   = tk.StringVar(
            value=lagra.get("utmappe_rot", _DEFAULT_UTMAPPE))

        self._output_oversikt:  str | None  = None
        self._output_rib_arm:   list[str]  = []
        self._output_apm:       str | None  = None
        self._output_prognose:  str | None  = None
        self._output_omega_oppl: str | None = None
        self._running           = False

        self._build_ui()

        # Lagre val når vindauget vert lukka
        self.protocol("WM_DELETE_WINDOW", self._ved_lukking)

    # ── Lagring av filval mellom køyringar ──────────────────────────────────────

    def _les_innstillingar(self) -> dict:
        try:
            with open(_INNSTILLINGAR_FIL, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def _lagre_innstillingar(self) -> None:
        data = {
            "csv_stiar":     list(self._csv_lb.get(0, "end")),
            "isy_stiar":     list(self._isy_lb.get(0, "end")),
            "omega":         self._omega_var.get().strip(),
            "apm_mappe":     self._apm_mappe.get().strip(),
            "rib_arm":       bool(self._rib_arm.get()),
            "fob_mengde":    bool(self._fob_mengde.get()),
            "hopp_over_apm": bool(self._hopp_over_apm.get()),
            "lag_prognose":  bool(self._lag_prognose.get()),
            "periode":       self._periode.get().strip(),
            "omega_overstyrer": bool(self._omega_overstyrer.get()),
            "omega_mal":     self._omega_mal.get().strip(),
            "utmappe_rot":   self._utmappe_rot.get().strip(),
        }
        try:
            with open(_INNSTILLINGAR_FIL, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError:
            pass  # ikkje kritisk om lagring feilar

    def _ved_lukking(self) -> None:
        self._lagre_innstillingar()
        self.destroy()

    # ── UI-bygging ────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        p = self.PAD
        outer = ttk.Frame(self, padding=p)
        outer.pack(fill="both", expand=True)

        # Overskrift
        ttk.Label(
            outer,
            text="Mengde Oversikt — Filvalg",
            font=("Segoe UI", 13, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, p))

        # ── CSV-listeboks ────────────────────────────────────────────────────────
        ttk.Label(outer, text="CSV-filer:", anchor="e", width=22).grid(
            row=1, column=0, sticky="ne", padx=(0, 4), pady=2,
        )

        csv_frame = ttk.Frame(outer)
        csv_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=2)
        csv_frame.columnconfigure(0, weight=1)

        self._csv_lb = tk.Listbox(csv_frame, height=4, selectmode="extended",
                                   font=("Consolas", 9))
        sb_csv = ttk.Scrollbar(csv_frame, orient="vertical", command=self._csv_lb.yview)
        self._csv_lb.configure(yscrollcommand=sb_csv.set)
        self._csv_lb.grid(row=0, column=0, sticky="ew")
        sb_csv.grid(row=0, column=1, sticky="ns")

        csv_btn_frame = ttk.Frame(csv_frame)
        csv_btn_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Button(csv_btn_frame, text="Legg til …", width=10,
                   command=self._legg_til_csv).pack(side="left", padx=(0, 4))
        ttk.Button(csv_btn_frame, text="Fjern valgte", width=12,
                   command=self._fjern_csv).pack(side="left")

        for sti in self._csv_stiar:
            self._csv_lb.insert("end", sti)

        # ── Omega-fil ────────────────────────────────────────────────────────────
        self._fil_rad(outer, 2, "Omega-fil:", self._omega_var, _EXCEL_FT)

        # ── ISY Linker-listeboks ──────────────────────────────────────────────
        ttk.Label(outer, text="ISY Linker-filer:", anchor="e", width=22).grid(
            row=3, column=0, sticky="ne", padx=(0, 4), pady=2,
        )

        isy_frame = ttk.Frame(outer)
        isy_frame.grid(row=3, column=1, columnspan=2, sticky="ew", pady=2)
        isy_frame.columnconfigure(0, weight=1)

        self._isy_lb = tk.Listbox(isy_frame, height=3, selectmode="extended",
                                   font=("Consolas", 9))
        sb_isy = ttk.Scrollbar(isy_frame, orient="vertical", command=self._isy_lb.yview)
        self._isy_lb.configure(yscrollcommand=sb_isy.set)
        self._isy_lb.grid(row=0, column=0, sticky="ew")
        sb_isy.grid(row=0, column=1, sticky="ns")

        isy_btn_frame = ttk.Frame(isy_frame)
        isy_btn_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Button(isy_btn_frame, text="Legg til …", width=10,
                   command=self._legg_til_isy).pack(side="left", padx=(0, 4))
        ttk.Button(isy_btn_frame, text="Fjern valgte", width=12,
                   command=self._fjern_isy).pack(side="left")

        for sti in self._isy_stiar:
            self._isy_lb.insert("end", sti)

        # ── APM PDF-mappe ─────────────────────────────────────────────────────
        self._mappe_rad(outer, 4, "APM PDF-mappe:", self._apm_mappe)
        self._mappe_rad(outer, 5, "Lagre resultat i:", self._utmappe_rot)

        # Separator
        ttk.Separator(outer, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=p,
        )

        # Avkryssingsboksar
        ttk.Checkbutton(
            outer,
            text="Lag også fullstendig ARM/RIB-rapport  (Mengdeuttak_Armering)",
            variable=self._rib_arm,
        ).grid(row=6, column=0, columnspan=3, sticky="w")

        ttk.Checkbutton(
            outer,
            text="Bruk FOB_Mengde for RIB-volum  (i staden for PGF_Kalkyle)",
            variable=self._fob_mengde,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            outer,
            text="Hopp over APM-analyse  (pdf_maalebrev_til_excel)",
            variable=self._hopp_over_apm,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            outer,
            text="Lag prognoseark til FOB  (pris- og mengderapport med forside)",
            variable=self._lag_prognose,
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            outer,
            text="        La prognosen fra Omega gå foran den beregnede  "
                 "(poster satt til 0 eller med annen verdi enn kontraktsmengden)",
            variable=self._omega_overstyrer,
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Periode — vises på forsida i prognosearket
        per_frame = ttk.Frame(outer)
        per_frame.grid(row=11, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(per_frame, text="        Periode:").pack(side="left")
        ttk.Entry(per_frame, textvariable=self._periode, width=24).pack(
            side="left", padx=(4, 6))
        ttk.Label(per_frame, text='t.d. "September 2026"',
                  foreground="#808080").pack(side="left")

        # Omega-opplastingsark — valfritt. Vert kopiert med ny prognose i
        # ForecastQty-kolonnen. Tomt felt = hopp over.
        oppl_frame = ttk.Frame(outer)
        oppl_frame.grid(row=12, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        oppl_frame.columnconfigure(1, weight=1)
        ttk.Label(oppl_frame, text="        Omega-opplasting:").grid(
            row=0, column=0, sticky="w")
        ttk.Entry(oppl_frame, textvariable=self._omega_mal).grid(
            row=0, column=1, sticky="ew", padx=(4, 0))
        ttk.Button(
            oppl_frame, text="Bla …", width=6,
            command=lambda: self._bla(self._omega_mal, _EXCEL_FT),
        ).grid(row=0, column=2, sticky="w", padx=(4, 0))
        ttk.Button(
            oppl_frame, text="Tøm", width=6,
            command=lambda: self._omega_mal.set(""),
        ).grid(row=0, column=3, sticky="w", padx=(4, 0))

        # Knappepanel
        btn_frame = ttk.Frame(outer)
        btn_frame.grid(row=13, column=0, columnspan=3, sticky="e", pady=(p, 0))

        self._kjor_btn = ttk.Button(
            btn_frame, text="Kjør", width=12, command=self._start,
        )
        self._kjor_btn.pack(side="left", padx=(0, 4))

        ttk.Button(
            btn_frame, text="Avbryt", width=12, command=self._ved_lukking,
        ).pack(side="left")

        # Loggfelt
        ttk.Separator(outer, orient="horizontal").grid(
            row=13, column=0, columnspan=3, sticky="ew", pady=p,
        )

        log_frame = ttk.LabelFrame(outer, text="Logg", padding=4)
        log_frame.grid(row=14, column=0, columnspan=3, sticky="nsew")
        outer.rowconfigure(14, weight=1)
        outer.columnconfigure(1, weight=1)

        self._log = tk.Text(
            log_frame,
            height=18,
            state="disabled",
            wrap="word",
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
        )
        sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

        # Åpne-knapp (vises etter fullføring)
        self._opne_frame = ttk.Frame(outer)
        self._opne_frame.grid(row=15, column=0, columnspan=3, sticky="e", pady=(4, 0))

    def _fil_rad(
        self,
        parent: ttk.Frame,
        row: int,
        etikett: str,
        var: tk.StringVar,
        filetypes: list,
    ) -> None:
        p = self.PAD
        ttk.Label(parent, text=etikett, anchor="e", width=22).grid(
            row=row, column=0, sticky="e", padx=(0, 4), pady=2,
        )
        e = ttk.Entry(parent, textvariable=var, width=60)
        e.grid(row=row, column=1, sticky="ew", pady=2)
        ttk.Button(
            parent,
            text="Bla …",
            width=6,
            command=lambda v=var, ft=filetypes: self._bla(v, ft),
        ).grid(row=row, column=2, sticky="w", padx=(4, 0), pady=2)

    def _mappe_rad(
        self,
        parent: ttk.Frame,
        row: int,
        etikett: str,
        var: tk.StringVar,
    ) -> None:
        ttk.Label(parent, text=etikett, anchor="e", width=22).grid(
            row=row, column=0, sticky="e", padx=(0, 4), pady=2,
        )
        e = ttk.Entry(parent, textvariable=var, width=60)
        e.grid(row=row, column=1, sticky="ew", pady=2)
        ttk.Button(
            parent,
            text="Bla …",
            width=6,
            command=lambda v=var: self._bla_mappe(v),
        ).grid(row=row, column=2, sticky="w", padx=(4, 0), pady=2)

    # ── Filovalar ─────────────────────────────────────────────────────────────

    def _legg_til_csv(self) -> None:
        stiar = filedialog.askopenfilenames(
            initialdir=os.path.expanduser("~"),
            filetypes=_CSV_FT,
        )
        eksisterande = set(self._csv_lb.get(0, "end"))
        for sti in stiar:
            if sti and sti not in eksisterande:
                self._csv_lb.insert("end", sti)
                eksisterande.add(sti)

    def _fjern_csv(self) -> None:
        for idx in reversed(self._csv_lb.curselection()):
            self._csv_lb.delete(idx)

    def _legg_til_isy(self) -> None:
        stiar = filedialog.askopenfilenames(
            initialdir=os.path.expanduser("~"),
            filetypes=_EXCEL_FT,
        )
        eksisterande = set(self._isy_lb.get(0, "end"))
        for sti in stiar:
            if sti and sti not in eksisterande:
                self._isy_lb.insert("end", sti)
                eksisterande.add(sti)

    def _fjern_isy(self) -> None:
        for idx in reversed(self._isy_lb.curselection()):
            self._isy_lb.delete(idx)

    def _bla(self, var: tk.StringVar, filetypes: list) -> None:
        start = os.path.dirname(var.get()) if os.path.exists(os.path.dirname(var.get())) else os.path.expanduser("~")
        path = filedialog.askopenfilename(
            initialdir=start,
            filetypes=filetypes,
        )
        if path:
            var.set(path)

    def _bla_mappe(self, var: tk.StringVar) -> None:
        start = var.get() if os.path.isdir(var.get()) else os.path.expanduser("~")
        path = filedialog.askdirectory(initialdir=start)
        if path:
            var.set(path)

    # ── Køyring ───────────────────────────────────────────────────────────────

    def _log_linje(self, tekst: str) -> None:
        self._log.after(0, self._log_skriv, tekst)

    def _log_skriv(self, tekst: str) -> None:
        self._log.configure(state="normal")
        self._log.insert("end", tekst + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _start(self) -> None:
        if self._running:
            return

        csv_stiar = list(self._csv_lb.get(0, "end"))
        isy_stiar = list(self._isy_lb.get(0, "end"))
        omega     = self._omega_var.get().strip()

        if not csv_stiar:
            messagebox.showerror("Manglende fil", "Ingen CSV-fil er lagt til.")
            return
        if not isy_stiar:
            messagebox.showerror("Manglende fil", "Ingen ISY Linker-fil er lagt til.")
            return
        for sti in csv_stiar:
            if not os.path.exists(sti):
                messagebox.showerror("Manglende fil", f"Finner ikke CSV-fil:\n{sti}")
                return
        for sti in isy_stiar:
            if not os.path.exists(sti):
                messagebox.showerror("Manglende fil", f"Finner ikke ISY Linker-fil:\n{sti}")
                return
        if not omega or not os.path.exists(omega):
            messagebox.showerror("Manglende fil", f"Finner ikke Omega-fil:\n{omega}")
            return

        # Hugs filvala til neste køyring
        self._lagre_innstillingar()

        self._kjor_btn.configure(state="disabled")
        for w in self._opne_frame.winfo_children():
            w.destroy()
        self._output_oversikt = None
        self._output_rib_arm  = []
        self._output_apm      = None
        self._output_prognose = None
        self._output_omega_oppl = None
        self._output_feil = None
        self._running = True

        # Omdiriger stdout/stderr til loggfeltet
        stream = _WidgetStream(self._log)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = stream
        sys.stderr = stream

        apm_mappe    = self._apm_mappe.get().strip()
        hopp_over_apm = self._hopp_over_apm.get()
        lag_prognose  = self._lag_prognose.get()
        periode       = self._periode.get().strip()
        omega_overstyrer = self._omega_overstyrer.get()
        omega_mal        = self._omega_mal.get().strip()
        utmappe_rot      = self._utmappe_rot.get().strip()

        def _run():
            try:
                self._kjor_analyse(
                    csv_stiar, omega, isy_stiar,
                    self._rib_arm.get(), self._fob_mengde.get(),
                    apm_mappe, hopp_over_apm,
                    lag_prognose, periode, omega_overstyrer, omega_mal,
                    utmappe_rot,
                )
            except Exception as exc:
                import traceback
                print(f"\n[FEIL] {exc}")
                print(traceback.format_exc())
            finally:
                sys.stdout = old_out
                sys.stderr = old_err
                self._running = False
                # stopp() rører Tk, så han må gå på hovudtråden
                self.after(0, stream.stopp)
                self.after(0, self._kjor_btn.configure, {"state": "normal"})
                self.after(0, self._vis_opne_knapper)

        threading.Thread(target=_run, daemon=True).start()

    def _kjor_analyse(
        self, csv_path: list[str], omega_path: str, isy_path: list[str],
        lag_rib_arm: bool, fob_mengde: bool,
        apm_mappe: str = "", hopp_over_apm: bool = False,
        lag_prognose: bool = False, periode: str = "",
        omega_overstyrer: bool = True, omega_mal: str = "",
        utmappe_rot: str = "",
    ) -> None:
        """Køyr heile kjeda.

        APM-uttrekket kjem FØRST fordi mengde_oversikt les APM-mengdene frå
        Maalebrev_samlet.xlsx. Køyrde vi det til slutt (slik det var før),
        fekk oversikta APM-tal frå førre månad.
        """
        from mengde_oversikt import kjor_arm_rib_per_omraade, lag_oversikt
        from csv_til_armrib import les_csv

        # Alle resultatfilene frå køyringa vert samla i månadsmappa, under
        # rota frå "Lagre resultat i". Er feltet tomt, fell vi tilbake på
        # mappa Omega-fila ligg i — den gamle åtferda. Det er verdt å skilje
        # dei: Omega-fila vert gjerne henta frå Nedlastingar, og då skal
        # ikkje rapportane hamne der.
        base = (utmappe_rot or "").strip()
        if base and not os.path.isdir(base):
            try:
                os.makedirs(base, exist_ok=True)
                print(f"[INFO] Oppretta utmappa: {base}")
            except OSError as exc:
                print(f"[ADVARSEL] Klarte ikkje lage utmappa {base}: {exc}")
                base = ""
        if not base:
            base = os.path.dirname(os.path.abspath(omega_path))
            print(f"[INFO] Inga utmappe valt — brukar mappa til Omega-fila")
        if not os.path.isdir(base):
            from lagring import excel_mappe
            base = excel_mappe()
        utmappe = _maanadsmappe(base, periode)
        print(f"[INFO] Resultatfiler vert lagra i: {utmappe}")

        # Feil mengdelisteposter vert samla gjennom heile køyringa og skrivne
        # til slutt — sjå feil_poster.py
        import feil_poster
        feil_poster.nullstill()

        steg = 0
        df_csv = None      # CSV-ane vert lesne éin gong og delt av alle stega

        def overskrift(tekst: str) -> None:
            nonlocal steg
            steg += 1
            print()
            print("=" * 60)
            print(f"STEG {steg} — {tekst}")
            print("=" * 60)

        # ── APM-målebrev frå PDF ──────────────────────────────────────────
        apm_excel: str | None = None
        if not hopp_over_apm and apm_mappe and os.path.isdir(apm_mappe):
            overskrift("APM Målebrev (pdf_maalebrev_til_excel.py)")
            from pdf_maalebrev_til_excel import extract_folder
            apm_ut = extract_folder(
                mappe=apm_mappe,
                omega_path=omega_path,
                isy_path=isy_path if isy_path else None,
            )
            if apm_ut:
                # Kopi, ikkje flytting: originalen må stå att i APM-mappa,
                # for det er der han vert lesen frå når APM-steget vert
                # hoppa over i ei seinare køyring.
                apm_ut = _kopier_til(apm_ut, utmappe)
                self._output_apm = apm_ut
                apm_excel = apm_ut
                print(f"\n[OK] APM-rapport: {apm_ut}")
        elif not hopp_over_apm and apm_mappe:
            print(f"[ADVARSEL] APM-mappe finnes ikke: {apm_mappe}")
        elif hopp_over_apm:
            # Datasettet har ingen APM-ar. Tom sti = ingen APM-data, elles fall
            # oversikta tilbake på den faste APM_EXCEL-stien i mengde_oversikt.py
            # og henta APM-mengder frå eit anna prosjekt.
            apm_excel = ""
            print("[INFO] APM-steget hoppa over — ingen APM-mengder i denne køyringa")

        # ── ARM/RIB ───────────────────────────────────────────────────────
        # ── Les CSV-ane ───────────────────────────────────────────────────
        # Éin gong for heile køyringa. Før vart dei same filene lesne på nytt
        # i kvart steg — fire fulle gjennomlesingar av fleire hundre MB.
        overskrift("Leser CSV-ane")
        df_csv = les_csv(csv_path, lp_only=True)
        print(f"[OK] {len(df_csv):,} rader x {len(df_csv.columns)} kolonnar "
              f"— gjenbrukast i alle stega under")

        # Foreldrekodar og feil postnummer i modellen → feillista.
        # Feillista skal aldri velte sjølve køyringa.
        try:
            from csv_til_armrib import _les_omega_koder
            feil_poster.sjekk_modellposter(df_csv, _les_omega_koder(omega_path)[0])
        except Exception as exc:
            print(f"[ADVARSEL] Sjekk av modellpostar feila: {exc}")

        arm_rib_excels: list[str] = []
        if lag_rib_arm:
            overskrift("Mengdeuttak ARM/RIB (Mengdeuttak_Armering.py)")
            arm_rib_excels = kjor_arm_rib_per_omraade(
                csv_path=csv_path,
                omega_path=omega_path,
                fob_mengde_mode=fob_mengde,
                df=df_csv,
            )
            arm_rib_excels = [_flytt_til(p, utmappe) for p in arm_rib_excels]
            self._output_rib_arm = arm_rib_excels
            for p in arm_rib_excels:
                print(f"\n[OK] ARM/RIB-rapport: {p}")

        # ── Mengde Oversikt ───────────────────────────────────────────────
        overskrift("Mengde Oversikt")
        oversikt_args = dict(
            csv_path=csv_path,
            omega_path=omega_path,
            isy_path=isy_path,
            arm_rib_excels=arm_rib_excels if arm_rib_excels else None,
            df_csv=df_csv,
            utdir=utmappe,
        )
        if apm_excel is not None:
            oversikt_args["apm_excel"] = apm_excel
        if apm_excel:
            print(f"[INFO] APM-mengder frå: {apm_excel}")
        oversikt = lag_oversikt(**oversikt_args)
        self._output_oversikt = oversikt
        print(f"\n[OK] Oversikt: {oversikt}")

        # ── Prognoseark til FOB ───────────────────────────────────────────
        if lag_prognose:
            overskrift("Prognoseark (prognose_ark.py)")
            from prognose_ark import lag_prognose as _lag_prognose
            # Opplastingsarket får eit eksplisitt namn, så vi veit kvar det
            # hamnar og kan tilby ein opne-knapp etterpå.
            omega_ut = None
            if omega_mal:
                from datetime import datetime as _dt
                omega_ut = os.path.join(
                    utmappe, f"omega_opplastning_{_dt.now():%Y%m%d_%H%M%S}.xlsx")

            prognose = _lag_prognose(
                oversikt_excel=oversikt,
                omega_path=omega_path,
                isy_path=isy_path,
                periode=periode,
                omega_overstyrer=omega_overstyrer,
                utdir=utmappe,
                omega_mal=omega_mal or None,
                omega_ut=omega_ut,
            )
            self._output_prognose = prognose
            if omega_ut and os.path.exists(omega_ut):
                self._output_omega_oppl = omega_ut

        overskrift("Feil mengdelisteposter")
        try:
            self._output_feil = feil_poster.skriv_excel(utmappe)
        except Exception as exc:
            print(f"[ADVARSEL] Klarte ikkje skrive feillista: {exc}")

        print("\nFerdig!")

    def _vis_opne_knapper(self) -> None:
        if self._output_oversikt and os.path.exists(self._output_oversikt):
            ttk.Button(
                self._opne_frame,
                text="Åpne Oversikt",
                command=lambda: _opne_fil(self._output_oversikt),
            ).pack(side="left", padx=4, pady=4)

        for path in self._output_rib_arm:
            if not os.path.exists(path):
                continue
            name = os.path.basename(path)
            if "_LYA" in name:
                label = "ARM/RIB LYA"
            elif "_VEK" in name:
                label = "ARM/RIB VEK"
            else:
                label = "ARM/RIB-rapport"
            ttk.Button(
                self._opne_frame,
                text=f"Åpne {label}",
                command=lambda p=path: _opne_fil(p),
            ).pack(side="left", padx=4, pady=4)

        if self._output_apm and os.path.exists(self._output_apm):
            ttk.Button(
                self._opne_frame,
                text="Åpne APM-rapport",
                command=lambda: _opne_fil(self._output_apm),
            ).pack(side="left", padx=4, pady=4)

        if self._output_prognose and os.path.exists(self._output_prognose):
            ttk.Button(
                self._opne_frame,
                text="Åpne Prognose",
                command=lambda: _opne_fil(self._output_prognose),
            ).pack(side="left", padx=4, pady=4)

        if self._output_omega_oppl and os.path.exists(self._output_omega_oppl):
            ttk.Button(
                self._opne_frame,
                text="Åpne Omega-opplasting",
                command=lambda: _opne_fil(self._output_omega_oppl),
            ).pack(side="left", padx=4, pady=4)

        if self._output_feil and os.path.exists(self._output_feil):
            ttk.Button(
                self._opne_frame,
                text="Åpne feilliste",
                command=lambda: _opne_fil(self._output_feil),
            ).pack(side="left", padx=4, pady=4)


# ── Hjelpefunksjonar ─────────────────────────────────────────────────────────

def _flytt_til(sti: str, mappe: str) -> str:
    """Flytt ei produsert fil til utmappa. Returnerer den nye stien.

    Mengdeuttak_Armering skriv til si eiga skriptmappe. For at alle filene
    frå ei køyring skal hamne saman med grunnlaget, vert dei flytta hit.
    """
    if not sti or not os.path.exists(sti) or not mappe or not os.path.isdir(mappe):
        return sti
    maal = os.path.join(mappe, os.path.basename(sti))
    if os.path.abspath(sti) == os.path.abspath(maal):
        return sti
    try:
        os.replace(sti, maal)
        return maal
    except OSError as exc:
        print(f"[ADVARSEL] Klarte ikkje flytte {os.path.basename(sti)}: {exc}")
        return sti


def _maanadsnamn(dato=None) -> str:
    """Norsk månadsnamn med årstal, t.d. 'September 2026'."""
    from datetime import date
    d = dato or date.today()
    return f"{_NOR_MAANED[d.month - 1]} {d.year}"


def _maanadsmappe(base: str, periode: str) -> str:
    """Mappa denne månaden sine filer skal liggje i.

    Namnet kjem frå Periode-feltet, elles frå dagens dato. Ligg base alt i
    ei mappe med det namnet, vert han brukt som han er — vi lagar ikkje
    'September 2026/September 2026'. Feilar opprettinga, held vi oss til base
    i staden for å velte køyringa.
    """
    namn = "".join(c for c in (periode or "").strip()
                   if c not in _UGYLDIGE_MAPPETEIKN).strip()
    if not namn:
        namn = _maanadsnamn()

    if os.path.basename(os.path.normpath(base)).strip().lower() == namn.lower():
        return base

    maal = os.path.join(base, namn)
    try:
        os.makedirs(maal, exist_ok=True)
    except OSError as exc:
        print(f"[ADVARSEL] Klarte ikkje lage månadsmappa {maal}: {exc}")
        return base
    return maal


def _kopier_til(sti: str, mappe: str) -> str:
    """Kopier ei produsert fil til utmappa. Returnerer stien til kopien.

    Til skilnad frå _flytt_til vert originalen ståande. APM-rapporten vert
    lesen frå si eiga mappe når APM-steget vert hoppa over, så han kan ikkje
    flyttast bort derifrå.
    """
    if not sti or not os.path.exists(sti) or not mappe or not os.path.isdir(mappe):
        return sti
    maal = os.path.join(mappe, os.path.basename(sti))
    if os.path.abspath(sti) == os.path.abspath(maal):
        return sti
    try:
        shutil.copy2(sti, maal)
        return maal
    except OSError as exc:
        print(f"[ADVARSEL] Klarte ikkje kopiere {os.path.basename(sti)}: {exc}")
        return sti


def _opne_fil(sti: str) -> None:
    """Opna fil med standardprogram (Windows)."""
    try:
        os.startfile(sti)
    except Exception:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(sti)])


# ── Inngangspunkt ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = MengdeApp()
    app.mainloop()
