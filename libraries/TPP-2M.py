# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

import numpy as np


def calculate_imfp_tpp2m(kinetic_energy):
    """
    Calculate IMFP using TPP-2M formula with Briggs & Grant average matrix values

    Parameters:
        kinetic_energy: electron energy in eV

    Returns:
        imfp: Inelastic Mean Free Path in nanometers
    """
    # Average matrix values from Briggs & Grant
    N_v = 4.684  # Average number of valence electrons
    rho = 6.767  # Density in g/cm³
    M = 137.51  # Molecular weight
    E_g = 0  # Bandgap energy

    # Calculate plasmon energy
    E_p = 28.8 * np.sqrt((N_v * rho) / M)

    # Calculate U
    U = N_v * rho / M

    # Calculate parameters
    beta = -0.10 + 0.944 / (E_p ** 2 + E_g ** 2) ** 0.5 + 0.069 * rho ** 0.1
    gamma = 0.191 * rho ** (-0.5)
    C = 1.97 - 0.91 * U
    D = 53.4 - 20.8 * U

    # Calculate IMFP
    imfp = kinetic_energy / (E_p ** 2 * (beta * np.log(gamma * kinetic_energy) -
                                         (C / kinetic_energy) + (D / kinetic_energy ** 2))) / 10

    return imfp


def main():
    ke = float(input("Enter kinetic energy (eV): "))
    imfp = calculate_imfp_tpp2m(ke)
    print(f"IMFP = {imfp:.4f} nm")
    # for ke in range(100, 1401, 100):
    #     imfp = calculate_imfp_tpp2m(ke)
    #     print(f"{imfp:.3f}")

if __name__ == "__main__":
    main()