import numpy as np
from scipy.optimize import minimize

# Input data
# Predicted positions (N, E, H) for rows 1 to 8
predicted_positions = np.array([
    [1215389.285, 111528.586, 45.52],
    [1215392.143, 111457.179, 41.848],
    [1215360.763, 111364.964, 37.145],
    [1215331.863, 111321.719, 34.783],
    [1215292.25, 111280.499, 33.335],
    [1215224.138, 111242.845, 30.578],
    [1215134.051, 111190.97, 27.875],  # Row 7 (no errors)
    [1215063.28, 111124.919, 25.898]   # Row 8 (no errors)
])

# Observed errors (dn, de, h) for rows 1 to 6
observed_errors = np.array([
    [50, -47, 1],
    [43, -48, -2],
    [28, -43, -2],
    [23, -40, -2],
    [19, -35, -2],
    [14, -28, -2]
])

# True positions for rows 1 to 6 (calculated as predicted + error)
true_positions = predicted_positions[:6] + observed_errors

# Function to compute the rotation matrix from Euler angles
def rotation_matrix(angles):
    rx, ry, rz = angles
    Rx = np.array([[1, 0, 0],
                   [0, np.cos(rx), -np.sin(rx)],
                   [0, np.sin(rx), np.cos(rx)]])
    Ry = np.array([[np.cos(ry), 0, np.sin(ry)],
                   [0, 1, 0],
                   [-np.sin(ry), 0, np.cos(ry)]])
    Rz = np.array([[np.cos(rz), -np.sin(rz), 0],
                   [np.sin(rz), np.cos(rz), 0],
                   [0, 0, 1]])
    return Rz @ Ry @ Rx

# Objective function to minimize
def objective(params):
    angles = params[:3]  # Euler angles for rotation
    translation = params[3:]  # Translation vector
    R = rotation_matrix(angles)
    errors = true_positions - (predicted_positions[:6] @ R.T + translation)
    return np.sum(errors**2)

# Initial guess for Euler angles and translation
initial_guess = np.zeros(6)  # [rx, ry, rz, tx, ty, tz]

# Solve the least-squares problem
result = minimize(objective, initial_guess, method='BFGS')
optimal_params = result.x

# Extract optimal rotation and translation
optimal_angles = optimal_params[:3]
optimal_translation = optimal_params[3:]
R_optimal = rotation_matrix(optimal_angles)

# Predict errors for rows 7 and 8
predicted_errors_7 = (predicted_positions[6] @ R_optimal.T + optimal_translation) - predicted_positions[6]
predicted_errors_8 = (predicted_positions[7] @ R_optimal.T + optimal_translation) - predicted_positions[7]

# Print results
print("Optimal rotation angles (radians):", optimal_angles)
print("Optimal translation vector:", optimal_translation)
print("Predicted errors for row 7 (dn, de, h):", predicted_errors_7)
print("Predicted errors for row 8 (dn, de, h):", predicted_errors_8)