import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import minimize

# Definer maksimal feil på ±10 mm
max_deviation = 0.010  # i mm

# Definer de fire hjørnene av rektangelet (opprinnelige boltplasseringer)
corners = np.array([
    [90, 10],    # Øvre venstre (Bolt 1)
    [90, 110],  # Øvre høyre (Bolt 2)
    [10, 10],     # Nedre venstre (Bolt 3)
    [10, 110]  # Nedre høyre (Bolt 4)
])

# Funksjon for å beregne avstanden mellom to punkter
def calculate_distance(p1, p2):
    return np.linalg.norm(p1 - p2)

# Målfunksjon for optimering (minimer total feil)
def objective_function(perturbed_corners_flat):
    perturbed_corners = perturbed_corners_flat.reshape(-1, 2)  # Omforme til 2D-matrise
    total_error = np.sum((corners - perturbed_corners) ** 2)
    return total_error

def constraint_function(perturbed_corners_flat):
    perturbed_corners = perturbed_corners_flat.reshape(-1, 2)
    constraints = []
    
    # Hver bolt skal ikke flyttes mer enn ±10 mm
    for i in range(len(corners)):
        bolt_error = calculate_distance(corners[i], perturbed_corners[i])
        constraints.append(max_deviation - bolt_error)
    
    # Begrensning på inbyrdes x- og y-avstander
    for i in range(len(corners)):
        for j in range(i + 1, len(corners)):
            x_error = abs(perturbed_corners[i][0] - perturbed_corners[j][0]) - abs(corners[i][0] - corners[j][0])
            y_error = abs(perturbed_corners[i][1] - perturbed_corners[j][1]) - abs(corners[i][1] - corners[j][1])
            constraints.append(max_deviation - abs(x_error))
            constraints.append(max_deviation - abs(y_error))
    
    # Begrensning på diagonaler
    diagonal_error_1_4 = abs(calculate_distance(perturbed_corners[0], perturbed_corners[3]) - calculate_distance(corners[0], corners[3]))
    diagonal_error_3_2 = abs(calculate_distance(perturbed_corners[2], perturbed_corners[1]) - calculate_distance(corners[2], corners[1]))
    constraints.append(max_deviation - diagonal_error_1_4)
    constraints.append(max_deviation - diagonal_error_3_2)
    
    return np.array(constraints)

# Startgjett for perturbed corners
initial_guess = corners.flatten()

bounds = []
for c in corners:
    bounds.append((c[0] - max_deviation, c[0] + max_deviation))  # x-koordinat
    bounds.append((c[1] - max_deviation, c[1] + max_deviation))  # y-koordinat

# Utfør optimering med strengere toleranse
solution = minimize(
    objective_function,
    initial_guess,
    method='SLSQP',
    bounds=bounds,
    constraints={'type': 'ineq', 'fun': constraint_function},
    options={'ftol': 1e-9, 'disp': True}  # Strengere toleranse for nøyaktighet
)

# Hent de optimerte perturbed corners
perturbed_corners = solution.x.reshape(-1, 2)

# Skriv ut resultater
print("\nOpprinnelige hjørner:")
print(corners)
print("\nForskjøvede hjørner:")
print(perturbed_corners)

# Beregn feil for hver bolt
bolt_errors = [calculate_distance(corners[i], perturbed_corners[i]) for i in range(len(corners))]
print("\nMaksimal feil for hver bolt:")
for i, error in enumerate(bolt_errors):
    print(f"Bolt {i + 1}: {error:.2f} mm")

# Beregn og skriv ut inbyrdes feil i x, y og diagonaler
print("\nInbyrdes feil:")
for i in range(len(corners)):
    for j in range(i + 1, len(corners)):
        x_error = abs(perturbed_corners[i][0] - perturbed_corners[j][0]) - abs(corners[i][0] - corners[j][0])
        y_error = abs(perturbed_corners[i][1] - perturbed_corners[j][1]) - abs(corners[i][1] - corners[j][1])
        diagonal_error = calculate_distance(perturbed_corners[i], perturbed_corners[j]) - calculate_distance(corners[i], corners[j])
        print(f"Bolt {i + 1} og Bolt {j + 1}:")
        print(f"  Feil i x-retning: {x_error:.2f} mm")
        print(f"  Feil i y-retning: {y_error:.2f} mm")
        print(f"  Feil diagonalt: {diagonal_error:.2f} mm")

# Plott resultatene
plt.figure(figsize=(6, 6))
plt.plot([perturbed_corners[i][0] for i in [0,1,2,3,0]],
         [perturbed_corners[i][1] for i in [0,1,2,3,0]], 'r--', label="Feiljustert ramme")
plt.scatter(corners[:, 0], corners[:, 1], color='black', label="Opprinnelige bolter", s=100)
plt.scatter(perturbed_corners[:, 0], perturbed_corners[:, 1], color='blue', label="Forskjøvede bolter", s=100)
plt.xlim(-10, 150)
plt.ylim(-10, 150)
plt.legend()
plt.title("Visualisering av maksimal feil")
plt.grid(True, linestyle="--", linewidth=0.5)
plt.show()
