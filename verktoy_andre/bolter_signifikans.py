from tkinter import messagebox
import tkinter as tk
import math
from scipy import stats
import webbrowser
import os

# Funksjon for å beregne nødvendig utvalgsstørrelse
def calculate_sample_size(population_size, error_margin):
    # Konstantverdier
    confidence_level = 0.95
    p = 0.50
    
    # Z-verdi for 95% konfidensnivå
    z_value = stats.norm.ppf(1 - (1 - confidence_level) / 2)
    
    # Beregn utvalgsstørrelse
    sample_size = (z_value**2 * p * (1 - p)) / error_margin**2
    
    # Juster utvalgsstørrelsen for endelig populasjon
    finite_sample_size = (population_size * sample_size) / ((sample_size - 1) + population_size)
    
    return math.ceil(finite_sample_size), z_value, sample_size, finite_sample_size

# Funksjon for å hente input og vise resultat i HTML
def calculate():
    try:
        population_size = int(entry_population_size.get())
        uncontrolled_bolts = int(entry_uncontrolled_bolts.get())
        error_margin = float(entry_error_margin.get()) / 100

        # Beregn nødvendig utvalgsstørrelse
        sample_size_needed, z_value, sample_size, finite_sample_size = calculate_sample_size(population_size, error_margin)
        prosent = round(sample_size_needed / population_size * 100, 2)

        # Lag HTML-innhold med MathJax for matematikkvisning
        html_content = f"""
        <html>
        <head>
            <title>Resultat</title>
            <script type="text/javascript" async
                src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-MML-AM_CHTML">
            </script>
        </head>
        <body>
            <h2>Hvor mange bolter må trekkes?</h2>
            <p>For å være innenfor kravet om at det er signifikant sannsynlig at 
            <b>{uncontrolled_bolts}</b> bolter som ikke kan kontrolleres også er i orden, må du kontrollere minst 
            <b>{sample_size_needed}</b> bolter. Dette er <b>{prosent}%</b> av totalen: <b>{population_size}</b>.</p>
            
            <h3>Matematisk beregning:</h3>
            <p>Utvalgsstørrelsen beregnes ved hjelp av følgende formel:</p>
            <p>
                <b>Formel for utvalgsstørrelse:</b><br>
                $$ n = \\frac{{Z^2 \\cdot p \\cdot (1 - p)}}{{E^2}} $$
            </p>
            <p>Der:</p>
            <ul>
                <li><b>n</b> = Utvalgsstørrelsen</li>
                <li><b>Z</b> = Z-verdi for konfidensnivået (for 95% konfidensnivå er Z = {z_value:.2f})</li>
                <li><b>p</b> = Antatt pålitelighet (p = 0.50)</li>
                <li><b>E</b> = Feilmarginen (i vårt tilfelle: {error_margin * 100}%)</li>
            </ul>
            <p>Etter beregning av utvalgsstørrelsen, justeres den for en endelig populasjon med følgende formel:</p>
            <p>
                <b>Justert utvalgsstørrelse:</b><br>
                $$ n_{{adj}} = \\frac{{N \\cdot n}}{{(n - 1) + N}} $$
            </p>
            <p>Der:</p>
            <ul>
                <li><b>n_adj</b> = Justert utvalgsstørrelse</li>
                <li><b>N</b> = Totalt antall bolter (populasjonsstørrelse)</li>
                <li><b>n</b> = Opprinnelig beregnet utvalgsstørrelse</li>
            </ul>
        </body>
        </html>
        """
        
        # Lagre HTML-fil
        file_path = "resultat.html"
        with open(file_path, "w") as file:
            file.write(html_content)
        
        # Åpne HTML-fil i nettleseren
        webbrowser.open(f"file://{os.path.abspath(file_path)}")
         # Lukk GUI-vinduet etter at HTML er åpnet
        root.quit()
    except Exception as e:
        messagebox.showerror("Feil", f"Ugyldig input: {e}")

# Opprett hovedvinduet
root = tk.Tk()
root.title("Utvalgsstørrelse Kalkulator")

# Input-felt
tk.Label(root, text="Totalt antall bolter:").grid(row=0, column=0, padx=10, pady=5)
entry_population_size = tk.Entry(root)
entry_population_size.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Antall bolter som ikke kan kontrolleres:").grid(row=1, column=0, padx=10, pady=5)
entry_uncontrolled_bolts = tk.Entry(root)
entry_uncontrolled_bolts.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Feilmargin (%):").grid(row=2, column=0, padx=10, pady=5)
entry_error_margin = tk.Entry(root)
entry_error_margin.grid(row=2, column=1, padx=10, pady=5)

# Beregn-knapp
calculate_button = tk.Button(root, text="Beregn utvalgsstørrelse", command=calculate)
calculate_button.grid(row=3, column=0, columnspan=2, pady=10)

# Start GUI-en
root.mainloop()