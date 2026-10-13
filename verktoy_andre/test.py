import sympy as sp
import numpy as np

# Definer symboler
x, y, z, p = sp.symbols('x y z p')

# Definer likningene
eq1 = sp.Eq(sp.sqrt(x**2 + y**2), 10)
eq2 = sp.Eq(sp.sqrt(x**2 + z**2), 10)
eq3 = sp.Eq(sp.sqrt(y**2 + p**2), 10)
eq4 = sp.Eq(sp.sqrt(p**2 + z**2), 10)

# Løs systemet av likninger
solutions = sp.solve((eq1, eq2, eq3, eq4), (x, y, z, p))

# Skriv ut løsningene
print("Løsninger:")
for sol in solutions:
    print(sol)

# Funksjon for å beregne feil
def calculate_error(sol, x_val, y_val, z_val, p_val):
    error1 = abs(sp.sqrt(x_val**2 + y_val**2) - 10)
    error2 = abs(sp.sqrt(x_val**2 + z_val**2) - 10)
    error3 = abs(sp.sqrt(y_val**2 + p_val**2) - 10)
    error4 = abs(sp.sqrt(p_val**2 + z_val**2) - 10)
    return (error1 + error2 + error3 + error4) / 4

# Generer 100 tilfeldige verdier for x, y, z, p
antall_ganger = 600
np.random.seed(0)  # For reproduserbarhet
x_vals = np.random.uniform(-10, 10, antall_ganger)
y_vals = np.random.uniform(-10, 10, antall_ganger)
z_vals = np.random.uniform(-10, 10, antall_ganger)
p_vals = np.random.uniform(-10, 10, antall_ganger)

# Beregn gjennomsnittsfeil for hver løsning

for i, sol in enumerate(solutions):
    total_error = 0
    for j in range(antall_ganger):
        x_val = x_vals[j]
        y_val = y_vals[j]
        z_val = z_vals[j]
        p_val = p_vals[j]
        total_error += calculate_error(sol, x_val, y_val, z_val, p_val)
    avg_error = total_error / antall_ganger
    print(f"Gjennomsnittsfeil for løsning {i+1}: {avg_error}")