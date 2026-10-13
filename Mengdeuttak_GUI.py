import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import subprocess

# Mappa denne fila ligg i — skripta vert funne her uansett kvar mappa er flytta
_HER = os.path.dirname(os.path.abspath(__file__))

class MengdetakGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mengdeuttak - Filvalg og Validering")
        self.root.geometry("600x400")
        
        # Filvariabler
        self.armering_file = tk.StringVar(value="Velg Armering-fil...")
        self.rib_file = tk.StringVar(value="Velg RIB-fil...")
        self.omega_file = tk.StringVar(value="Velg OMEGA-fil...")
        self.last_output_file = None  # Siste genererte Excel-fil
        
        self.setup_ui()
    
    def setup_ui(self):
        """Opprett GUI-elementer"""
        
        # Tittel
        title = tk.Label(self.root, text="Velg Excel-filer for Mengdeuttak", font=("Arial", 14, "bold"))
        title.pack(pady=10)
        
        # Armering
        frame_arm = tk.Frame(self.root)
        frame_arm.pack(fill="x", padx=20, pady=5)
        tk.Label(frame_arm, text="Armering-fil:", width=15, anchor="w").pack(side="left")
        self.arm_label = tk.Label(frame_arm, textvariable=self.armering_file, bg="lightgray", anchor="w")
        self.arm_label.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame_arm, text="Bla gjennom", command=self.select_armering).pack(side="left")
        
        # RIB
        frame_rib = tk.Frame(self.root)
        frame_rib.pack(fill="x", padx=20, pady=5)
        tk.Label(frame_rib, text="RIB-fil:", width=15, anchor="w").pack(side="left")
        self.rib_label = tk.Label(frame_rib, textvariable=self.rib_file, bg="lightgray", anchor="w")
        self.rib_label.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame_rib, text="Bla gjennom", command=self.select_rib).pack(side="left")
        
        # OMEGA
        frame_omega = tk.Frame(self.root)
        frame_omega.pack(fill="x", padx=20, pady=5)
        tk.Label(frame_omega, text="OMEGA-fil:", width=15, anchor="w").pack(side="left")
        self.omega_label = tk.Label(frame_omega, textvariable=self.omega_file, bg="lightgray", anchor="w")
        self.omega_label.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame_omega, text="Bla gjennom", command=self.select_omega).pack(side="left")
        
        # Knapper
        frame_buttons = tk.Frame(self.root)
        frame_buttons.pack(pady=20)
        tk.Button(frame_buttons, text="Valider Filer", command=self.validate_files, bg="lightblue", width=15).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Kjør Mengdeuttak", command=self.run_mengdeuttak, bg="lightgreen", width=15).pack(side="left", padx=5)
        tk.Button(frame_buttons, text="Avslutt", command=self.avslutt, bg="lightcoral", width=15).pack(side="left", padx=5)
        
        # Status-område
        self.status_text = tk.Text(self.root, height=10, width=70, state="disabled")
        self.status_text.pack(padx=10, pady=10, fill="both", expand=True)
    
    def select_armering(self):
        """Velg Armering-fil"""
        file = filedialog.askopenfilename(
            title="Velg Armering Excel-fil",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.armering_file.set(file)
    
    def select_rib(self):
        """Velg RIB-fil"""
        file = filedialog.askopenfilename(
            title="Velg RIB Excel-fil",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.rib_file.set(file)
    
    def select_omega(self):
        """Velg OMEGA-fil"""
        file = filedialog.askopenfilename(
            title="Velg OMEGA Excel-fil",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.omega_file.set(file)
    
    def log_status(self, message):
        """Skriv status-melding"""
        self.status_text.config(state="normal")
        self.status_text.insert("end", message + "\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()
    
    def clear_status(self):
        """Tøm status-området"""
        self.status_text.config(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.config(state="disabled")
    
    def validate_files(self):
        """Validér filene"""
        self.clear_status()
        self.log_status("="*60)
        self.log_status("VALIDERING AV FILER")
        self.log_status("="*60 + "\n")
        
        arm_file = self.armering_file.get()
        rib_file = self.rib_file.get()
        omega_file = self.omega_file.get()
        
        # Sjekk om filene er valgt
        if "Velg" in arm_file or "Velg" in rib_file or "Velg" in omega_file:
            self.log_status("[FEIL] Alle filer må velges!\n")
            messagebox.showerror("Feil", "Alle filer må velges!")
            return False
        
        # Sjekk at filene finnes
        all_valid = True
        
        if not os.path.exists(arm_file):
            self.log_status(f"[FEIL] Armering-fil finnes ikke: {arm_file}\n")
            all_valid = False
        else:
            self.log_status(f"[OK] Armering-fil finnes: {os.path.basename(arm_file)}")
        
        if not os.path.exists(rib_file):
            self.log_status(f"[FEIL] RIB-fil finnes ikke: {rib_file}\n")
            all_valid = False
        else:
            self.log_status(f"[OK] RIB-fil finnes: {os.path.basename(rib_file)}")
        
        if not os.path.exists(omega_file):
            self.log_status(f"[FEIL] OMEGA-fil finnes ikke: {omega_file}\n")
            all_valid = False
        else:
            self.log_status(f"[OK] OMEGA-fil finnes: {os.path.basename(omega_file)}")
        
        if not all_valid:
            messagebox.showerror("Feil", "En eller flere filer finnes ikke!")
            return False
        
        self.log_status("\n" + "-"*60)
        self.log_status("Validerer ark og kolonner...")
        self.log_status("-"*60 + "\n")
        
        # Valider Armering
        try:
            xls = pd.ExcelFile(arm_file)
            if "Armering" not in xls.sheet_names:
                self.log_status(f"[FEIL] 'Armering' ark finnes ikke i {os.path.basename(arm_file)}\n")
                all_valid = False
            else:
                df = pd.read_excel(arm_file, sheet_name="Armering")
                required_cols = ["FOB_Info.FOB_Produksjonsenhet", "PGF_RIB_Info.Vekt"]
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    self.log_status(f"[FEIL] Armering mangler kolonner: {missing}\n")
                    all_valid = False
                else:
                    self.log_status(f"[OK] Armering-fil OK ({len(df)} rader, alle kolonner funnet)")
        except Exception as e:
            self.log_status(f"[FEIL] Feil ved lesing av Armering: {str(e)}\n")
            all_valid = False
        
        # Valider RIB
        try:
            xls = pd.ExcelFile(rib_file)
            if "RIB" not in xls.sheet_names:
                self.log_status(f"[FEIL] 'RIB' ark finnes ikke i {os.path.basename(rib_file)}\n")
                all_valid = False
            else:
                df = pd.read_excel(rib_file, sheet_name="RIB")
                required_cols = ["FOB_Info.FOB_Produksjonsenhet", "PGF_Kalkyle.PGF_Volum"]
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    self.log_status(f"[FEIL] RIB mangler kolonner: {missing}\n")
                    all_valid = False
                else:
                    self.log_status(f"[OK] RIB-fil OK ({len(df)} rader, alle kolonner funnet)")
        except Exception as e:
            self.log_status(f"[FEIL] Feil ved lesing av RIB: {str(e)}\n")
            all_valid = False
        
        # Valider OMEGA
        try:
            xls = pd.ExcelFile(omega_file)
            from base_lesning import finn_omega_ark
            omega_ark = finn_omega_ark(omega_file)
            if omega_ark not in xls.sheet_names:
                self.log_status(f"[FEIL] Fant ingen Omega-ark i {os.path.basename(omega_file)}\n")
                all_valid = False
            else:
                df = pd.read_excel(omega_file, sheet_name=omega_ark, skiprows=2)
                required_cols = ["Navn (Prosesskode)", "Tittel", "Enhet"]
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    self.log_status(f"[FEIL] OMEGA mangler kolonner: {missing}\n")
                    all_valid = False
                else:
                    self.log_status(f"[OK] OMEGA-fil OK ({len(df)} rader, alle kolonner funnet)")
        except Exception as e:
            self.log_status(f"[FEIL] Feil ved lesing av OMEGA: {str(e)}\n")
            all_valid = False
        
        self.log_status("\n" + "="*60)
        if all_valid:
            self.log_status("[OK] ALLE FILER VALIDERT SUCCESSFULLY!")
            messagebox.showinfo("Suksess", "Alle filer validert OK!")
            return True
        else:
            self.log_status("[FEIL] EN ELLER FLERE VALIDERINGSFEIL FUNNET")
            messagebox.showerror("Feil", "En eller flere valideringsfeil funnet!")
            return False
    
    def run_mengdeuttak(self):
        """Kjør hovedscriptet"""
        if not self.validate_files():
            return
        
        arm_file = self.armering_file.get()
        rib_file = self.rib_file.get()
        omega_file = self.omega_file.get()
        
        self.log_status("\n" + "="*60)
        self.log_status("STARTER MENGDEUTTAK-SCRIPT...")
        self.log_status("="*60 + "\n")
        
        # Kjør scriptet som subprocess med environment variables
        import sys
        import tempfile
        
        try:
            # Lag et wrapper-script som setter variablene først.
            # Wrapperen ligg i ei temp-mappe, så __file__ og sys.path vert sette
            # til mappa Mengdeuttak_Armering.py ligg i (same mappe som denne fila).
            armering_script = os.path.join(_HER, "Mengdeuttak_Armering.py")
            wrapper_script = f"""
import sys
sys.path.insert(0, r"{_HER}")
__file__ = r"{armering_script}"

# Sett filstiene direkte i dette scriptets globale scope
excel_file_armering = r"{arm_file}"
excel_file_RIB = r"{rib_file}"
excel_file_OMEGA = r"{omega_file}"

# Kjør hovedscriptet - exec() vil arve dette scriptets globale variabler
with open(__file__, "r", encoding="utf-8") as f:
    exec(f.read())
"""
            
            # Skriv wrapper-scriptet til temp-fil
            with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False, encoding='utf-8') as f:
                f.write(wrapper_script)
                temp_script = f.name
            
            # Kjør wrapper-scriptet
            result = subprocess.run(
                [sys.executable, temp_script],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            
            # Vis output
            if result.stdout:
                self.log_status(result.stdout)
                # Parse output-filstien fra "[SUCCESS] Fil eksportert: ..."
                for line in result.stdout.splitlines():
                    if line.strip().startswith("[SUCCESS] Fil eksportert:"):
                        self.last_output_file = line.strip().split("[SUCCESS] Fil eksportert:", 1)[1].strip()
                        break
            if result.stderr:
                self.log_status("[STDERR]\n" + result.stderr)
            
            # Slett temp-filen
            import os
            try:
                os.remove(temp_script)
            except:
                pass
            
            if result.returncode == 0:
                self.log_status("\n" + "="*60)
                self.log_status("[OK] MENGDEUTTAK KJORT SUCCESSFULLY!")
                if self.last_output_file:
                    self.log_status(f"[INFO] Fil: {self.last_output_file}")
                self.log_status("="*60)
                messagebox.showinfo("Suksess", "Mengdeuttak kjort OK!\nTrykk Avslutt for aa aapne Excel-filen.")
            else:
                self.log_status(f"\n[FEIL] Script returnerte kode: {result.returncode}")
                messagebox.showerror("Feil", f"Script feilet med kode: {result.returncode}")
        
        except Exception as e:
            self.log_status(f"\n[FEIL] {str(e)}")
            import traceback
            self.log_status("\n" + traceback.format_exc())
            messagebox.showerror("Feil", f"Feil ved kjøring: {str(e)}")

    def avslutt(self):
        """Lukk GUI og aapne Excel-filen hvis den finnes"""
        filsti = self.last_output_file
        self.root.destroy()
        if filsti and os.path.exists(filsti):
            os.startfile(filsti)


if __name__ == "__main__":
    root = tk.Tk()
    gui = MengdetakGUI(root)
    root.mainloop()
