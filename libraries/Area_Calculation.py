# Area_Calculation.py - Atomic concentration and weight percentage calculations

# Atomic masses dictionary for all elements
ATOMIC_MASSES = {
    'H': 1.008, 'He': 4.003, 'Li': 6.94, 'Be': 9.012, 'B': 10.81, 'C': 12.01, 'N': 14.01, 'O': 16.00,
    'F': 19.00, 'Ne': 20.18, 'Na': 22.99, 'Mg': 24.31, 'Al': 26.98, 'Si': 28.09, 'P': 30.97, 'S': 32.07,
    'Cl': 35.45, 'Ar': 39.95, 'K': 39.10, 'Ca': 40.08, 'Sc': 44.96, 'Ti': 47.87, 'V': 50.94, 'Cr': 52.00,
    'Mn': 54.94, 'Fe': 55.85, 'Co': 58.93, 'Ni': 58.69, 'Cu': 63.55, 'Zn': 65.38, 'Ga': 69.72, 'Ge': 72.63,
    'As': 74.92, 'Se': 78.96, 'Br': 79.90, 'Kr': 83.80, 'Rb': 85.47, 'Sr': 87.62, 'Y': 88.91, 'Zr': 91.22,
    'Nb': 92.91, 'Mo': 95.96, 'Tc': 98.00, 'Ru': 101.07, 'Rh': 102.91, 'Pd': 106.42, 'Ag': 107.87, 'Cd': 112.41,
    'In': 114.82, 'Sn': 118.71, 'Sb': 121.76, 'Te': 127.60, 'I': 126.90, 'Xe': 131.29, 'Cs': 132.91, 'Ba': 137.33,
    'La': 138.91, 'Ce': 140.12, 'Pr': 140.91, 'Nd': 144.24, 'Pm': 145.00, 'Sm': 150.36, 'Eu': 151.96, 'Gd': 157.25,
    'Tb': 158.93, 'Dy': 162.50, 'Ho': 164.93, 'Er': 167.26, 'Tm': 168.93, 'Yb': 173.05, 'Lu': 174.97, 'Hf': 178.49,
    'Ta': 180.95, 'W': 183.84, 'Re': 186.21, 'Os': 190.23, 'Ir': 192.22, 'Pt': 195.08, 'Au': 196.97, 'Hg': 200.59,
    'Tl': 204.38, 'Pb': 207.2, 'Bi': 208.98, 'Po': 209.00, 'At': 210.00, 'Rn': 222.00, 'Fr': 223.00, 'Ra': 226.00,
    'Ac': 227.00, 'Th': 232.04, 'Pa': 231.04, 'U': 238.03
}


def extract_element_symbol(peak_name):
    """Extract element symbol from peak name like 'C1s', 'Ti2p3/2', 'Fe2p1/2', etc."""
    import re

    # Remove spaces and convert to proper case
    clean_name = peak_name.replace(' ', '').strip()

    # Try to match element at the beginning of the string
    # Pattern: One uppercase letter followed by optional lowercase letter, then numbers/letters
    match = re.match(r'^([A-Z][a-z]?)', clean_name)
    if match:
        element = match.group(1)
        # Check if it's a valid element
        if element in ATOMIC_MASSES:
            return element

    # Fallback: try to find any valid element symbol in the string
    for element in ATOMIC_MASSES:
        if clean_name.upper().startswith(element.upper()):
            return element

    # Default fallback
    return 'C'  # Default to Carbon if no element found


def calculate_weight_percentages(atomic_percentages, peak_names):
    """
    Calculate weight percentages from atomic percentages and peak names.

    Args:
        atomic_percentages: List of atomic percentages
        peak_names: List of peak names to extract element symbols from

    Returns:
        List of weight percentages
    """
    if len(atomic_percentages) != len(peak_names):
        raise ValueError("Number of atomic percentages must match number of peak names")

    weight_data = []
    total_weight_sum = 0

    # Calculate weight contributions
    for atomic_percent, peak_name in zip(atomic_percentages, peak_names):
        element_symbol = extract_element_symbol(peak_name)
        atomic_mass = ATOMIC_MASSES.get(element_symbol, 12.01)  # Default to carbon mass
        weight_contribution = atomic_percent * atomic_mass
        total_weight_sum += weight_contribution
        weight_data.append(weight_contribution)

    # Calculate weight percentages
    weight_percentages = []
    for weight_contribution in weight_data:
        weight_percent = (weight_contribution / total_weight_sum) * 100 if total_weight_sum > 0 else 0
        weight_percentages.append(weight_percent)

    return weight_percentages


def get_atomic_mass(element_symbol):
    """Get atomic mass for a given element symbol."""
    return ATOMIC_MASSES.get(element_symbol.capitalize(), 12.01)  # Default to carbon mass