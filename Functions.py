# FUNCTIONS.PY-------------------------------------------------------------------

# LIBRARIES----------------------------------------------------------------------
import wx.grid

import numpy as np
import lmfit

import sys
from scipy.stats import linregress

from libraries.Save import refresh_sheets, create_plot_script_from_excel
from libraries.Peak_Functions import PeakFunctions, BackgroundCalculations
from libraries.Sheet_Operations import on_sheet_selected
from libraries.Save import save_results_table, save_all_sheets_with_plots
from libraries.Help import on_about
from libraries.Help import show_shortcuts, show_mini_game
from libraries.Save import undo, redo, save_state, update_undo_redo_state
from libraries.Open import update_recent_files, import_avantage_file, open_avg_file, import_multiple_avg_files
from libraries.Utilities import load_rsf_data
from libraries.Grid_Operations import populate_results_grid





# -------------------------------------------------------------------------------

def save_data_wrapper(window, data):
    from libraries.Save import save_data
    save_data(window, data)

def on_sheet_selected_wrapper(window, event):
    from libraries.Sheet_Operations import on_sheet_selected
    on_sheet_selected(window, event)



def safe_delete_rows(grid, pos, num_rows):
    try:
        if pos >= 0 and num_rows > 0 and (pos + num_rows) <= grid.GetNumberRows():
            grid.DeleteRows(pos, num_rows)
        else:
            wx.MessageBox("Invalid row indices for deletion.", "Information", wx.OK | wx.ICON_INFORMATION)
    except Exception as e:
        print("Error")





def remove_peak(window):
    save_state(window)
    num_rows = window.peak_params_grid.GetNumberRows()
    if num_rows > 0:
        sheet_name = window.sheet_combobox.GetValue()

        # Remove the last two rows from the peak_params_grid
        if num_rows >= 2:
            safe_delete_rows(window.peak_params_grid, num_rows - 2, 2)
        elif num_rows == 1:
            safe_delete_rows(window.peak_params_grid, num_rows - 1, 1)
        else:
            wx.MessageBox("No rows to delete.", "Information", wx.OK | wx.ICON_INFORMATION)
            return

        # Decrease the peak count
        window.peak_count = num_rows // 2 - 1  # Update peak count based on remaining rows

        # Remove the last peak from window.Data
        if 'Fitting' in window.Data['Core levels'][sheet_name] and 'Peaks' in window.Data['Core levels'][sheet_name][
            'Fitting']:
            peaks = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            if peaks:
                last_peak_key = f"p{window.peak_count + 1}"
                if last_peak_key in peaks:
                    del peaks[last_peak_key]
                else:
                    # If the key doesn't match the expected pattern, remove the last item
                    peaks.popitem()

        # Reset the selected peak index
        window.selected_peak_index = None

        # Call the method to clear and replot everything
        window.clear_and_replot()

        # Layout the updated panel
        window.panel.Layout()

        window.canvas.draw_idle()
    else:
        wx.MessageBox("No peaks to remove.", "Information", wx.OK | wx.ICON_INFORMATION)







def clear_plot(window):
    window.ax.clear()
    window.canvas.draw()

    # Reinitialize the background to raw data
    window.background = None







def update_sheet_names(window):
    if window.selected_files:
        file_path = window.selected_files[0]
        try:
            sheet_names = pd.ExcelFile(file_path).sheet_names
            window.sheet_combobox.Set(sheet_names)
            if sheet_names:
                window.sheet_combobox.SetSelection(0)  # Set the first sheet as selected
        except Exception as e:
            wx.MessageBox(f"Error reading sheet names: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
    else:
        wx.MessageBox("No files selected.", "Information", wx.OK | wx.ICON_INFORMATION)


def rename_sheet(window, new_sheet_name):
    """
    Rename a selected sheet in one or more Excel files.

    This function renames a specified sheet in selected Excel files. It iterates through
    the selected files, reads the content of the specified sheet, and writes it back
    with the new sheet name. Other sheets in the file remain unchanged.

    Args:
        window: The main application window object, containing necessary UI elements.
        new_sheet_name (str): The new name to be given to the selected sheet.

    Note:
        This function assumes the existence of certain attributes in the window object:
        - file_listbox: A listbox containing selected file names.
        - sheet_combobox: A combobox with the current sheet name.
        - entry: An entry field containing the directory path of the files.

    Raises:
        Exceptions are caught and displayed in a message box.
    """
    selected_indices = window.file_listbox.GetSelections()
    sheet_name = window.sheet_combobox.GetValue()

    if selected_indices:
        for i in selected_indices:
            filename = window.file_listbox.GetString(i)
            file_path = os.path.join(window.entry.GetValue(), filename)

            try:
                with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                    df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        for sheet in xls.sheet_names:
                            if sheet == sheet_name:
                                df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                            else:
                                pd.read_excel(xls, sheet_name=sheet, engine='openpyxl').to_excel(writer, sheet_name=sheet, index=False)
            except Exception as e:
                wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)


def on_exit(window, event):
    """Handles the Exit menu item."""
    window.Destroy()
    wx.GetApp().ExitMainLoop()



def toggle_plot(window):
    window.show_fit = not window.show_fit
    sheet_name = window.sheet_combobox.GetValue()
    if window.show_fit:
        window.clear_and_replot()
    else:
        window.plot_manager.plot_data(window)
    window.canvas.draw_idle()

import json
from libraries.ConfigFile import Init_Measurement_Data, add_core_level_Data



""""
def convert_from_serializable(obj):
    if isinstance(obj, list):
        return [convert_from_serializable(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: convert_from_serializable(v) for k, v in obj.items()}
    else:
        return obj
"""


import shutil
from vamas import Vamas
from openpyxl import Workbook







def on_save_plot(window):
    from libraries.Save import save_plot_as_png
    save_plot_as_png(window)


def on_save_plot_pdf(window):
    from libraries.Save import save_plot_as_pdf
    save_plot_as_pdf(window)


def on_save_plot_svg(window):
    from libraries.Save import save_plot_as_svg
    save_plot_as_svg(window)


def on_save(window):
    from libraries.Save import save_data, save_results_table
    data = window.get_data_for_save()
    save_data(window, data)
    # save_results_table(window)  # Add this line to also save results table


def on_save_all_sheets(window, event):
    from libraries.Save import save_all_sheets_with_plots
    save_all_sheets_with_plots(window)

def toggle_Col_1(window):
    # List of columns to toggle
    columns1 = [13, 14, 15, 16,17, 18]
    columns2 = [2,4,12,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28]


    # Check if the first column in the list is hidden or shown
    if window.peak_params_grid.IsColShown(columns1[0]):
        for col in columns1:
            window.peak_params_grid.HideCol(col)
        for col in columns2:
            window.results_grid.HideCol(col)
    else:
        for col in columns1:
            window.peak_params_grid.ShowCol(col)
        for col in columns2:
            window.results_grid.ShowCol(col)

    # print(window.Data)




def calculate_r2(y_true, y_pred):
    """Calculate the coefficient of determination (R²)"""
    ss_res = np.sum((y_true - y_pred) ** 2)
    ss_tot = np.sum((y_true - np.mean(y_true)) ** 2)
    return 1 - (ss_res / ss_tot)


def calculate_chi_square(y_true, y_pred):
    """Calculate the chi-square value"""
    return np.sum((y_true - y_pred) ** 2 / y_pred)



from matplotlib.ticker import ScalarFormatter


def fit_peaks(window, peak_params_grid, evaluate=False):
    """
    Perform peak fitting on the spectral data and update the peak parameters.
    """
    global fraction
    if peak_params_grid is None or peak_params_grid.GetNumberRows() == 0:
        wx.MessageBox("No peak parameters defined. Please add at least one peak before fitting.", "Error",
                      wx.OK | wx.ICON_ERROR)
        return None

    sheet_name = window.sheet_combobox.GetValue()

    if sheet_name not in window.plot_config.plot_limits:
        window.plot_config.update_plot_limits(window, sheet_name)
    limits = window.plot_config.plot_limits[sheet_name]

    if sheet_name not in window.Data['Core levels']:
        wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
        return None

    core_level_data = window.Data['Core levels'][sheet_name]
    x_values = np.array(core_level_data['B.E.'])
    y_values = np.array(core_level_data['Raw Data'])
    background = np.array(core_level_data['Background']['Bkg Y'])

    num_peaks = peak_params_grid.GetNumberRows() // 2

    bg_min_energy = core_level_data['Background'].get('Bkg Low')
    bg_max_energy = core_level_data['Background'].get('Bkg High')

    try:
        bg_min_energy = float(bg_min_energy)
        bg_max_energy = float(bg_max_energy)
    except (ValueError, TypeError):
        bg_min_energy = min(x_values)
        bg_max_energy = max(x_values)

    if bg_min_energy is not None and bg_max_energy is not None and bg_min_energy <= bg_max_energy:
        mask = (x_values >= bg_min_energy) & (x_values <= bg_max_energy)
        x_values_filtered = x_values[mask]
        y_values_filtered = y_values[mask]
        background_filtered = background[mask]

        if len(x_values_filtered) > 0 and len(y_values_filtered) > 0:
            y_values_subtracted = y_values_filtered - background_filtered

            model_choice = window.selected_fitting_method
            max_nfev = window.max_iterations

            model = None
            params = lmfit.Parameters()

            individual_peaks = []

            for i in range(num_peaks):
                row = i * 2
                prefix = f'peak{i}_'  # Define the prefix here

                center = float(peak_params_grid.GetCellValue(row, 2))
                height = float(peak_params_grid.GetCellValue(row, 3))
                fwhm = float(peak_params_grid.GetCellValue(row, 4))
                lg_ratio = float(peak_params_grid.GetCellValue(row, 5))
                try:
                    fwhm_g = float(peak_params_grid.GetCellValue(row, 9))
                    skew = float(peak_params_grid.GetCellValue(row, 9))
                except ValueError:
                    fwhm_g = 0.64
                    skew = 0.1
                try:
                    area = float(peak_params_grid.GetCellValue(row, 6))
                except ValueError:
                    area = 0  # Or any default value you prefer
                peak_model_choice = peak_params_grid.GetCellValue(row, 13)

                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                gamma = lg_ratio/100 * sigma

                center_min, center_max, center_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 2),
                                                                        center, peak_params_grid, i, "Position")
                height_min, height_max, height_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 3),
                                                                        height, peak_params_grid, i, "Height")
                fwhm_min, fwhm_max, fwhm_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 4),
                                                                  fwhm, peak_params_grid, i, "FWHM")
                lg_ratio_min, lg_ratio_max, lg_ratio_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 5),
                                                                              lg_ratio, peak_params_grid, i, "L/G")
                area_min, area_max, area_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 6),
                                                                  area, peak_params_grid, i, "area")

                center_min = evaluate_constraint(center_min, peak_params_grid, 'center', center)
                center_max = evaluate_constraint(center_max, peak_params_grid, 'center', center)
                height_min = evaluate_constraint(height_min, peak_params_grid, 'height', height)
                height_max = evaluate_constraint(height_max, peak_params_grid, 'height', height)
                area_min = evaluate_constraint(area_min, peak_params_grid, 'area', area)
                area_max = evaluate_constraint(area_max, peak_params_grid, 'area', area)
                if area_min == area_max:
                    area_max += 1e-6
                fwhm_min = evaluate_constraint(fwhm_min, peak_params_grid, 'fwhm', fwhm)
                fwhm_max = evaluate_constraint(fwhm_max, peak_params_grid, 'fwhm', fwhm)
                lg_ratio_min = evaluate_constraint(lg_ratio_min, peak_params_grid, 'lg_ratio', lg_ratio)
                lg_ratio_max = evaluate_constraint(lg_ratio_max, peak_params_grid, 'lg_ratio', lg_ratio)

                prefix = f'peak{i}_'
                if peak_model_choice == "Voigt (Area, L/G, \u03c3)":
                    try:
                        sigma = float(peak_params_grid.GetCellValue(row, 7)) / 2.355
                        fraction = float(peak_params_grid.GetCellValue(row, 5))  # L/G ratio
                    except ValueError:
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                        fraction = lg_ratio

                    # Parse constraints for sigma
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 7),
                                                                         sigma, peak_params_grid, i, "Sigma")
                    fraction_min, fraction_max, fraction_vary = parse_constraints(
                        peak_params_grid.GetCellValue(row + 1, 5),
                        fraction, peak_params_grid, i, "lg_ratio")

                    # Evaluate constraints
                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    fraction_min = evaluate_constraint(fraction_min, peak_params_grid, 'lg_ratio', fraction)
                    fraction_max = evaluate_constraint(fraction_max, peak_params_grid, 'lg_ratio', fraction)

                    # Calculate gamma, gamma_min, and gamma_max
                    def calc_gamma(f, s):
                        return (f * 2.355 * s) / (200 - 2 * f)

                    GAMMA_TOLERANCE = 1e-6  # Small tolerance value

                    # Gamma calculation section:
                    gamma = calc_gamma(fraction, sigma)
                    gamma_min = calc_gamma(fraction_min, sigma)
                    gamma_max = calc_gamma(fraction_max, sigma)

                    # Ensure gamma_min and gamma_max are different
                    if abs(gamma_max - gamma_min) < GAMMA_TOLERANCE:
                        gamma_min = max(0, gamma - GAMMA_TOLERANCE)
                        gamma_max = gamma + GAMMA_TOLERANCE

                    # Ensure gamma is within the range
                    gamma = max(gamma_min, min(gamma, gamma_max))

                    peak_model = lmfit.models.VoigtModel(prefix=prefix)
                    params.add(f'{prefix}area', value=area, min=area_min, max=area_max, vary=area_vary,
                               brute_step=area * 0.01)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary,
                               brute_step=0.1)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min/2.355, max=sigma_max/2.355,
                               vary=sigma_vary/2.355, brute_step=sigma * 0.01)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=fraction_vary,
                               brute_step=gamma * 0.01)

                    params.add(f'{prefix}amplitude', expr=f'{prefix}area')

                elif peak_model_choice == "Voigt (Area, L/G, \u03c3, S)":
                    try:
                        sigma = float(peak_params_grid.GetCellValue(row, 7)) / 2.355
                        fraction = float(peak_params_grid.GetCellValue(row, 5))  # L/G ratio
                        skew = float(peak_params_grid.GetCellValue(row, 9))
                    except ValueError:
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                        fraction = lg_ratio
                        skew = 0.0

                    # Parse constraints for sigma, fraction and skew
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 7),
                                                                         sigma, peak_params_grid, i, "Sigma")
                    fraction_min, fraction_max, fraction_vary = parse_constraints(
                        peak_params_grid.GetCellValue(row + 1, 5),
                        fraction, peak_params_grid, i, "lg_ratio")

                    skew_min, skew_max, skew_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 9),
                                                                      skew, peak_params_grid, i, "Skew")

                    # Evaluate constraints
                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    fraction_min = evaluate_constraint(fraction_min, peak_params_grid, 'lg_ratio', fraction)
                    fraction_max = evaluate_constraint(fraction_max, peak_params_grid, 'lg_ratio', fraction)
                    skew_min = evaluate_constraint(skew_min, peak_params_grid, 'skew', skew)
                    skew_max = evaluate_constraint(skew_max, peak_params_grid, 'skew', skew)

                    # Calculate gamma
                    def calc_gamma(f, s):
                        return (f * 2.355 * s) / (200 - 2 * f)

                    GAMMA_TOLERANCE = 1e-6

                    gamma = calc_gamma(fraction, sigma)
                    gamma_min = calc_gamma(fraction_min, sigma)
                    gamma_max = calc_gamma(fraction_max, sigma)

                    if abs(gamma_max - gamma_min) < GAMMA_TOLERANCE:
                        gamma_min = max(0, gamma - GAMMA_TOLERANCE)
                        gamma_max = gamma + GAMMA_TOLERANCE

                    gamma = max(gamma_min, min(gamma, gamma_max))
                    peak_model = lmfit.models.SkewedVoigtModel(prefix=prefix)
                    params.add(f'{prefix}area', value=area, min=area_min, max=area_max, vary=area_vary,
                               brute_step=area * 0.01)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary,
                               brute_step=0.1)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min / 2.355, max=sigma_max / 2.355,
                               vary=sigma_vary / 2.355, brute_step=sigma * 0.01)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=fraction_vary,
                               brute_step=gamma * 0.01)
                    params.add(f'{prefix}skew', value=skew, min=skew_min, max=skew_max, vary=skew_vary)#, brute_step)=skew * 0.001)
                    params.add(f'{prefix}amplitude', expr=f'{prefix}area')

                elif peak_model_choice == "DS (A, \u03c3, \u03b3)":
                    try:
                        # sigma = float(peak_params_grid.GetCellValue(row, 7)) / 2.355
                        sigma = float(peak_params_grid.GetCellValue(row, 7)) / 1
                        # gamma = float(peak_params_grid.GetCellValue(row, 8)) / 2
                        gamma = float(peak_params_grid.GetCellValue(row, 8)) / 1
                        skew = float(peak_params_grid.GetCellValue(row, 9))

                    except ValueError:
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                        gamma = lg_ratio / 100 * sigma
                        skew = 0.1

                    # Parse constraints
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 7),
                                                                         sigma, peak_params_grid, i, "Sigma")
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 8),
                                                                         gamma, peak_params_grid, i, "Gamma")
                    skew_min, skew_max, skew_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 9),
                                                                      skew, peak_params_grid, i, "Skew")

                    # Evaluate constraints
                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)
                    skew_min = evaluate_constraint(skew_min, peak_params_grid, 'skew', skew)
                    skew_max = evaluate_constraint(skew_max, peak_params_grid, 'skew', skew)

                    # Make sure skew is within reasonable bounds to avoid numerical issues
                    skew = max(0.01, min(skew, 0.99))
                    skew_min = max(0.01, skew_min)
                    skew_max = min(0.99, skew_max)

                    # After evaluating constraints for gamma
                    if gamma_min == gamma_max:
                        gamma_min = max(0, gamma_min - 0.0001)
                        gamma_max += 0.0001

                    peak_model = lmfit.models.DoniachModel(prefix=prefix)
                    params.add(f'{prefix}amplitude', value=area, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min, max=sigma_max, vary=sigma_vary)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=gamma_vary)
                    params.add(f'{prefix}asymmetry', value=skew, min=skew_min, max=skew_max, vary=skew_vary)

                elif peak_model_choice == "Voigt (Area, \u03c3, \u03b3)":
                    try:
                        sigma = float(peak_params_grid.GetCellValue(row, 7)) / 2.355
                        gamma = float(peak_params_grid.GetCellValue(row, 8)) / 2
                    except ValueError:
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))  # Default calculation if value is invalid
                        gamma = lg_ratio / 100 * sigma  # Default calculation if value is invalid

                    # Parse constraints for sigma and gamma
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1,
                                        7),sigma, peak_params_grid, i, "Sigma")
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1,
                                        8), gamma, peak_params_grid, i, "Gamma")

                    # Evaluate constraints
                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)

                    peak_model = lmfit.models.VoigtModel(prefix=prefix)
                    params.add(f'{prefix}area', value=area, min=area_min, max=area_max, vary=area_vary, brute_step=area * 0.01)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary, brute_step=0.1)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min/2.355, max=sigma_max/2.355,
                               vary=sigma_vary/2.355, brute_step=sigma * 0.01)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min/2, max=gamma_max/2, vary=gamma_vary/2,
                               brute_step=gamma*0.01)

                    params.add(f'{prefix}amplitude', expr=f'{prefix}area')


                elif peak_model_choice == "ExpGauss.(Area, \u03c3, \u03b3)":
                    try:
                        sigma = float(peak_params_grid.GetCellValue(row, 7)) / 1
                        gamma = float(peak_params_grid.GetCellValue(row, 8)) / 1
                    except ValueError:
                        print("ERROR CANNOT GET GAMMA")
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))  # Default calculation if value is invalid
                        gamma = lg_ratio / 100 * sigma  # Default calculation if value is invalid

                    # Parse constraints for sigma and gamma
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1,
                                                                                                       7),
                                                                         sigma, peak_params_grid, i, "Sigma")
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1,
                                                                                                           8),
                                                                         gamma, peak_params_grid, i, "Gamma")

                    # Evaluate constraints
                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)

                    peak_model = lmfit.models.ExponentialGaussianModel(prefix=prefix)
                    params.add(f'{prefix}amplitude', value=area, min=area_min, max=area_max, vary=area_vary, brute_step=area * 0.01)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary, brute_step=0.1)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min, max=sigma_max, vary=sigma_vary, brute_step=sigma * 0.01)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=gamma_vary, brute_step=gamma*0.01)

                elif peak_model_choice == "Pseudo-Voigt (Area)":
                    peak_model = lmfit.models.PseudoVoigtModel(prefix=prefix)
                    sigma = fwhm / 2.

                    params.add(f'{prefix}center', value=center,min=center_min, max=center_max,vary=center_vary, brute_step=0.1)
                    params.add(f'{prefix}area', value=area,min=area_min,max=area_max,vary=area_vary,brute_step=area * 0.01)
                    params.add(f'{prefix}sigma', value=sigma, min=fwhm_min / 2. if fwhm_min else None,
                        max=fwhm_max / 2. if fwhm_max else None, vary=fwhm_vary, brute_step=sigma * 0.01)
                    params.add(f'{prefix}fraction', value=lg_ratio / 100, min=lg_ratio_min / 100, max=lg_ratio_max / 100,
                               vary=lg_ratio_vary, brute_step=0.01)
                    params.add(f'{prefix}amplitude', expr=f'{prefix}area')


                elif peak_model_choice == "LA (Area, \u03c3, \u03b3)":
                    peak_model = lmfit.Model(PeakFunctions.LA, prefix=prefix)
                    amplitude = float(peak_params_grid.GetCellValue(row, 6))
                    fraction = float(peak_params_grid.GetCellValue(row, 5))  # L/G ratio
                    sigma = float(peak_params_grid.GetCellValue(row, 7))
                    gamma = float(peak_params_grid.GetCellValue(row, 8))

                    # Parse constraints
                    sigma_min, sigma_max, sigma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 7),
                                sigma, peak_params_grid, i, "Sigma")
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 8),
                                gamma, peak_params_grid, i, "Gamma")

                    sigma_min = evaluate_constraint(sigma_min, peak_params_grid, 'sigma', sigma)
                    sigma_max = evaluate_constraint(sigma_max, peak_params_grid, 'sigma', sigma)
                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)

                    params.add(f'{prefix}amplitude', value=amplitude, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=gamma_vary)
                    params.add(f'{prefix}sigma', value=sigma, min=sigma_min, max=sigma_max,vary=sigma_vary)
                elif peak_model_choice == "LA (Area, \u03c3/\u03b3, \u03b3)":
                    peak_model = lmfit.Model(PeakFunctions.LA, prefix=prefix)
                    amplitude = float(peak_params_grid.GetCellValue(row, 6))
                    fraction = float(peak_params_grid.GetCellValue(row, 5))  # L/G ratio
                    gamma = float(peak_params_grid.GetCellValue(row, 8))
                    sigma = (fraction / 100) * gamma / (1 - fraction / 100)  # Calculate sigma from L/G and gamma
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 8),
                                gamma, peak_params_grid, i, "Gamma")

                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)
                    params.add(f'{prefix}amplitude', value=amplitude, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=gamma_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max,vary=lg_ratio_vary)

                    # Add constraint to calculate sigma from L/G ratio and gamma
                    params.add(f'{prefix}sigma', expr=f'({prefix}fraction / 100) * {prefix}gamma / (1 -{prefix}fraction / 100)')
                elif peak_model_choice == "LA*G (Area, \u03c3/\u03b3, \u03b3)":
                    peak_model = lmfit.Model(PeakFunctions.LAxG, prefix=prefix)
                    amplitude = float(peak_params_grid.GetCellValue(row, 6))
                    fraction = float(peak_params_grid.GetCellValue(row, 5))  # L/G ratio
                    gamma = float(peak_params_grid.GetCellValue(row, 8))
                    fwhm_g = float(peak_params_grid.GetCellValue(row, 9))
                    sigma = (fraction / 100) * gamma / (1 - fraction / 100)  # Calculate sigma from L/G and gamma
                    gamma_min, gamma_max, gamma_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 8),
                                                                         gamma, peak_params_grid, i, "Gamma")
                    gamma_min = evaluate_constraint(gamma_min, peak_params_grid, 'gamma', gamma)
                    gamma_max = evaluate_constraint(gamma_max, peak_params_grid, 'gamma', gamma)

                    fwhm_g_min, fwhm_g_max, fwhm_g_vary = parse_constraints(peak_params_grid.GetCellValue(row + 1, 9),
                                                                         fwhm_g, peak_params_grid, i, "fwhm_g")
                    fwhm_g_min = evaluate_constraint(fwhm_g_min, peak_params_grid, 'fwhm_g', fwhm_g)
                    fwhm_g_max = evaluate_constraint(fwhm_g_max, peak_params_grid, 'fwhm_g', fwhm_g)

                    params.add(f'{prefix}amplitude', value=amplitude, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}gamma', value=gamma, min=gamma_min, max=gamma_max, vary=gamma_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max, vary=lg_ratio_vary)
                    params.add(f'{prefix}fwhm_g', value=fwhm_g, min=fwhm_g_min, max=fwhm_g_max,vary=True)

                    # Add constraint to calculate sigma from L/G ratio and gamma
                    params.add(f'{prefix}sigma',expr=f'({prefix}fraction / 100) * {prefix}gamma / (1 -{prefix}fraction / 100)')
                elif peak_model_choice == "GL (Area)":
                    peak_model = lmfit.Model(PeakFunctions.gauss_lorentz_Area, prefix=prefix)
                    params.add(f'{prefix}area', value=area, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max,
                               vary=lg_ratio_vary)
                elif peak_model_choice == "SGL (Area)":
                    peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz_Area, prefix=prefix)
                    params.add(f'{prefix}area', value=area, min=area_min, max=area_max, vary=area_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max,
                               vary=lg_ratio_vary)

                elif peak_model_choice == "GL (Height)":
                    peak_model = lmfit.Model(PeakFunctions.gauss_lorentz, prefix=prefix)
                    params.add(f'{prefix}amplitude', value=height, min=height_min, max=height_max, vary=height_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max,
                               vary=lg_ratio_vary)
                elif peak_model_choice == "SGL (Height)":
                    peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz, prefix=prefix)
                    params.add(f'{prefix}amplitude', value=height, min=height_min, max=height_max, vary=height_vary)
                    params.add(f'{prefix}center', value=center, min=center_min, max=center_max, vary=center_vary)
                    params.add(f'{prefix}fwhm', value=fwhm, min=fwhm_min, max=fwhm_max, vary=fwhm_vary)
                    params.add(f'{prefix}fraction', value=lg_ratio, min=lg_ratio_min, max=lg_ratio_max,
                               vary=lg_ratio_vary)
                elif peak_model_choice == "Unfitted":
                    return
                elif peak_model_choice == "D-parameter":
                    # Skip fitting for D-parameter
                    return
                else:
                    raise ValueError(f"Unknown fitting model: {peak_model_choice} for peak {i}")

                if model is None:
                    model = peak_model
                else:
                    model += peak_model

                individual_peaks.append(peak_model)

            optimization_method = window.fitting_window.get_optimization_method() if window.fitting_window else 'leastsq'
            # Define fit_kws only for methods that support it
            if optimization_method in ['leastsq', 'least_squares']:
                fit_kws = {'ftol': 1e-10, 'xtol': 1e-10}
            else:
                fit_kws = None  # Don't pass fit_kws for 'nelder', 'powell', or 'cobyla'

            if evaluate:
                # Use eval()
                result_eval = model.eval(params, x=x_values_filtered)
                residuals = y_values_subtracted - result_eval
                ss_res = np.sum(residuals ** 2)
                ss_tot = np.sum((y_values_subtracted - np.mean(y_values_subtracted)) ** 2)
                r_squared = 1 - (ss_res / ss_tot)
                window.r_squared = r_squared
                chi_square = ss_res
                red_chi_square = ss_res / (len(y_values_subtracted) - len(params))

                # Create a result object similar to fit() output
                result = type('Result', (), {
                    'best_fit': result_eval,
                    'params': params,
                    'chisqr': ss_res,
                    'redchi': ss_res / (len(y_values_subtracted) - len(params)),
                    'nfev': 1
                })

            else:
                # Use existing fit() code
                result = model.fit(
                    y_values_subtracted,
                    params,
                    x=x_values_filtered,
                    max_nfev=max_nfev,
                    method=optimization_method,
                    weights=np.ones(len(y_values_filtered)),
                    scale_covar=True,
                    nan_policy='omit',
                    verbose=True,
                    **({'fit_kws': fit_kws} if fit_kws else {})
                )
                residuals = y_values_subtracted - result.best_fit
                chi_square = result.chisqr
                red_chi_square = result.redchi
                ss_res = np.sum(residuals ** 2)
                ss_tot = np.sum((y_values_subtracted - np.mean(y_values_subtracted)) ** 2)
                r_squared = 1 - (ss_res / ss_tot)
                window.r_squared = r_squared


            if 'Fitting' not in window.Data['Core levels'][sheet_name]:
                window.Data['Core levels'][sheet_name]['Fitting'] = {}
            if 'Peaks' not in window.Data['Core levels'][sheet_name]['Fitting']:
                window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

            existing_peaks = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']

            for i in range(num_peaks):
                row = i * 2
                prefix = f'peak{i}_'
                peak_label = peak_params_grid.GetCellValue(row, 1)
                peak_model_choice = peak_params_grid.GetCellValue(row, 13)

                if peak_label in existing_peaks:
                    center = result.params[f'{prefix}center'].value
                    if peak_model_choice in["Voigt (Area, L/G, \u03c3)","Voigt (Area, \u03c3, \u03b3)"]: #, "Voigt (Area, L/G, \u03c3, S)"]:
                        amplitude = result.params[f'{prefix}amplitude'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        height = PeakFunctions.get_voigt_height(amplitude, sigma, gamma)
                        fwhm = PeakFunctions.voigt_fwhm(sigma, gamma)
                        fraction = (2*gamma) / (sigma*2.355 + 2*gamma) * 100
                        area = amplitude # * (sigma * np.sqrt(2 * np.pi))

                    elif peak_model_choice == "Voigt (Area, L/G, \u03c3, S)":
                        amplitude = result.params[f'{prefix}amplitude'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        skew  = result.params[f'{prefix}skew'].value
                        height = PeakFunctions.get_skewedvoigt_height(amplitude, sigma, gamma, skew)
                        peak_params_grid.SetCellValue(row, 3, f"{height:.2f}")  # Update height in grid
                        # amplitude = height / max_height
                        fwhm = PeakFunctions.skewed_voigt_fwhm(sigma, gamma, skew)
                        fraction = (2 * gamma) / (sigma * 2.355 + 2 * gamma) * 100
                        area = amplitude
                        # print(f'Into Voigt (Area, L/G, \u03c3, skew): H {height}  w {fwhm}')


                    elif peak_model_choice == "DS (A, \u03c3, \u03b3)":
                        amplitude = result.params[f'{prefix}amplitude'].value
                        center = result.params[f'{prefix}center'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        skew = result.params[f'{prefix}asymmetry'].value

                        # # Create DS model instance
                        # model = lmfit.models.DoniachModel()

                        # Get height directly from model
                        height = model.eval(x=[center], amplitude=amplitude, center=center,
                                            sigma=sigma, gamma=gamma, asymmetry=skew)[0]

                        # Calculate height numerically using the SAME x array
                        y_values = model.eval(x=x_values_filtered, amplitude=amplitude, center=center,
                                              sigma=sigma, gamma=gamma, asymmetry=skew)
                        # height = np.max(y_values)

                        # Estimate FWHM numerically
                        half_max = height / 2
                        indices = np.where(y_values >= half_max)[0]
                        if len(indices) >= 2:
                            fwhm = abs(x_values_filtered[indices[-1]] - x_values_filtered[indices[0]])
                        else:
                            fwhm = 2.355 * sigma  # Fallback to Gaussian FWHM
                        # sigma = round(float(sigma * 2.355), 2)
                        sigma = round(float(sigma * 1), 2)
                        # gamma = round(float(gamma * 2), 2)
                        gamma = round(float(gamma * 1), 2)
                        skew = round(float(skew), 2)
                        area = round(float(amplitude), 2)

                    elif peak_model_choice == "Pseudo-Voigt (Area)":
                        amplitude = result.params[f'{prefix}area'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        fraction = result.params[f'{prefix}fraction'].value * 100
                        fwhm = sigma * 2
                        height = PeakFunctions.get_pseudo_voigt_height(amplitude, sigma, fraction)
                        area = amplitude
                    elif peak_model_choice == "ExpGauss.(Area, \u03c3, \u03b3)":
                        amplitude = result.params[f'{prefix}amplitude'].value
                        center = result.params[f'{prefix}center'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        # Calculate height numerically
                        # Create the model
                        model = lmfit.models.ExponentialGaussianModel()
                        # Evaluate the model
                        y_values = model.eval(x=x_values, amplitude=amplitude, center=center, sigma=sigma, gamma=gamma)
                        height = np.max(y_values)
                        # Estimate FWHM numerically
                        half_max = height / 2
                        indices = np.where(y_values >= half_max)[0]
                        if len(indices) >= 2:
                            fwhm = abs(x_values[indices[-1]] - x_values[indices[0]])
                        else:
                            fwhm = None  # or some default value
                        fraction = gamma / (sigma + gamma) * 100
                        area = amplitude  # For area-based models, amplitude represents the area
                    elif peak_model_choice == "LA (Area, \u03c3, \u03b3)":
                        area = result.params[f'{prefix}amplitude'].value
                        center = result.params[f'{prefix}center'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value

                        # Calculate height numerically
                        y_values = PeakFunctions.LA(x_values, center, area, fwhm, sigma, gamma)
                        height = np.max(y_values)

                        # FOR RSD calc
                        existing_peaks[peak_label]['y_values'] = y_values

                        # No direct equivalent to 'fraction' for LA model
                        fraction = sigma / (sigma + gamma)
                    elif peak_model_choice in ["LA (Area, \u03c3/\u03b3, \u03b3)"]:
                        area = result.params[f'{prefix}amplitude'].value
                        center = result.params[f'{prefix}center'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        fraction = result.params[f'{prefix}fraction'].value /100
                        # area = window.calculate_peak_area(peak_model_choice, height, fwhm, fraction, sigma, gamma)

                        # Calculate height numerically
                        y_values = PeakFunctions.LA(x_values, center, area, fwhm, sigma, gamma)
                        height = np.max(y_values)

                        # FOR RSD calc
                        existing_peaks[peak_label]['y_values'] = y_values

                    elif peak_model_choice in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                        area = result.params[f'{prefix}amplitude'].value
                        center = result.params[f'{prefix}center'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        sigma = result.params[f'{prefix}sigma'].value
                        gamma = result.params[f'{prefix}gamma'].value
                        fraction = result.params[f'{prefix}fraction'].value /100
                        fwhm_g = result.params[f'{prefix}fwhm_g'].value

                        # Calculate height numerically
                        y_values = PeakFunctions.LAxG(x_values, center, area, fwhm, sigma, gamma, fwhm_g)
                        height = np.max(y_values)
                        existing_peaks[peak_label]['y_values'] = y_values

                    elif peak_model_choice in ["GL (Height)", "SGL (Height)"]:
                        height = result.params[f'{prefix}amplitude'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        fraction = result.params[f'{prefix}fraction'].value
                        area = height * fwhm * np.sqrt(np.pi / (4 * np.log(2)))
                    elif peak_model_choice in ["GL (Area)"]:
                        area = result.params[f'{prefix}area'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        fraction = result.params[f'{prefix}fraction'].value
                        height = area / (fwhm * np.sqrt(np.pi / (4 * np.log(2))))
                    elif peak_model_choice in ["SGL (Area)"]:
                        area = result.params[f'{prefix}area'].value
                        fwhm = result.params[f'{prefix}fwhm'].value
                        fraction = result.params[f'{prefix}fraction'].value
                        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                        gamma = fwhm / 2
                        height = area / ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (
                                    fraction / 100) * np.pi * gamma)
                    # elif peak_model_choice == "D-parameter":
                    else:
                        raise ValueError(f"Unknown fitting model: {peak_model_choice} for peak {peak_label}")

                    center = round(float(center), 2)
                    height = round(float(height), 2)
                    fwhm = round(float(fwhm), 2)
                    if peak_model_choice in ["ExpGauss.(Area, \u03c3, \u03b3)", "LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
                        # Exponential Gaussian doesn't use fraction
                        sigma = round(float(sigma * 1), 2)
                        gamma = round(float(gamma * 1), 2)
                        fraction = round(fraction * 100,2)
                        area = round(float(area), 2)
                    elif peak_model_choice in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                        # Exponential Gaussian doesn't use fraction
                        sigma = round(float(sigma * 1), 2)
                        gamma = round(float(gamma * 1), 2)
                        fraction = round(fraction * 100,2)
                        area = round(float(area), 2)
                        fwhm_g = round(float(fwhm_g), 2)
                    elif peak_model_choice in ["Voigt (Area, L/G, \u03c3, S)"]:
                        sigma = round(float(sigma * 2.355), 2)
                        gamma = round(float(gamma * 2), 2)
                        fraction = round(float(fraction), 2)
                        area = round(float(area), 2)
                    elif peak_model_choice == "DS (A, \u03c3, \u03b3)":
                        sigma = round(float(sigma * 1), 2)
                        gamma = round(float(gamma * 1), 2)
                        fraction = round(0.2 * 100, 2)
                        area = round(float(area), 2)
                    else:
                        sigma = round(float(sigma * 2.355), 2)
                        gamma = round(float(gamma * 2), 2)
                        fraction = round(float(fraction), 2)
                        area = round(float(area), 2)


                    peak_params_grid.SetCellValue(row, 2, f"{center:.2f}")
                    peak_params_grid.SetCellValue(row, 3, f"{height:.0f}")
                    peak_params_grid.SetCellValue(row, 4, f"{fwhm:.2f}")
                    peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                    peak_params_grid.SetCellValue(row, 6, f"{area:.0f}")
                    if peak_model_choice in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                             "Voigt (Area, L/G, \u03c3, S)",
                                             "ExpGauss.(Area, \u03c3, \u03b3)", "LA (Area, \u03c3, \u03b3)",
                                             "LA (Area, \u03c3/\u03b3, \u03b3)",
                                             "DS (A, \u03c3, \u03b3)"]:
                        peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                        peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                        if peak_model_choice in ["Voigt (Area, L/G, \u03c3, S)", "DS (A, \u03c3, \u03b3)"]:
                            peak_params_grid.SetCellValue(row, 9, f"{skew:.2f}")

                    elif peak_model_choice in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                        peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                        peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                        peak_params_grid.SetCellValue(row, 9, f"{fwhm_g:.2f}")
                    elif peak_model_choice == "D-parameter":
                        sigma = round(float(sigma * 1), 2)
                        gamma = round(float(gamma * 1), 2)
                        fraction = round(fraction,2)
                        fwhm_g = round(float(fwhm_g), 2)
                    else:
                        peak_params_grid.SetCellValue(row, 7, "")
                        peak_params_grid.SetCellValue(row, 8, "")
                        peak_params_grid.SetCellValue(row+1, 7, "")
                        peak_params_grid.SetCellValue(row+1, 8, "")
                    existing_peaks[peak_label].update({
                        'Position': center,
                        'Height': height,
                        'FWHM': fwhm,
                        'L/G': fraction,
                        'Area': area,
                        'Sigma': sigma,
                        'Gamma': gamma,
                        'fwhm_g':fwhm_g,
                        'Skew': skew,
                        'Fitting Model': peak_model_choice
                    })
                else:
                    print(f"Warning: Peak {peak_label} not found in existing data. Skipping update for this peak.")

            window.Data['Core levels'][sheet_name]['Fitting']['Model'] = model_choice

            # Calculate the RSD
            if any("LA" in peak_params_grid.GetCellValue(i * 2, 13) for i in range(num_peaks)):
                # For LA models, use stored y_values
                total_fit = np.zeros_like(x_values_filtered)
                for peak_label in existing_peaks:
                    if 'y_values' in existing_peaks[peak_label]:
                        total_fit += existing_peaks[peak_label]['y_values']
                rsd = round(PeakFunctions.calculate_rsd(y_values_filtered, total_fit + background_filtered), 3)
            else:
                # For other models
                rsd = round(PeakFunctions.calculate_rsd(y_values_filtered, result.best_fit + background_filtered), 3)

            window.fit_results = {
                'result': result,
                'rsd': rsd,
                'chi_square': chi_square,
                'red_chi_square': red_chi_square,
                'nfev': result.nfev,
                'fitted_peak': y_values.copy(),
                'mask': mask,
                'background_filtered': background_filtered,
                'y_values_subtracted': y_values_subtracted
            }
            window.fit_results['fitted_peak'][mask] = result.best_fit + background_filtered

            # Add text annotations with fit results
            std_value_int = int(window.noise_std_value) if hasattr(window, 'noise_std_value') else "N/A"

            window.update_ratios()
            window.clear_and_replot()

            # Fitting results --- THIS NEEDS TO BE SET AFTER CLEAR & REPLOT TO WORK
            window.plot_manager.set_fitting_results_text(f'Noise STD: {std_value_int}'
                                                         f' cps\nR²: {r_squared:.5f}\nChi²: {chi_square:.2f}\nRed. '
                                                         f'Chi²: {red_chi_square:.2f}\nIteration: {result.nfev}')

            return r_squared, rsd, red_chi_square

        else:
            raise ValueError("No data points found in the specified energy range for background subtraction")

    else:
        raise ValueError("Invalid background energy range")


def get_peak_value(peak_params_grid, peak_name, param_name):
    for i in range(peak_params_grid.GetNumberRows()):
        if peak_params_grid.GetCellValue(i, 0) == peak_name:
            fitting_model = peak_params_grid.GetCellValue(i, 12)
            if param_name == 'center':
                return float(peak_params_grid.GetCellValue(i, 2))
            elif param_name == 'height':
                return float(peak_params_grid.GetCellValue(i, 3))
            elif param_name == 'fwhm':
                return float(peak_params_grid.GetCellValue(i, 4))
            elif param_name == 'lg_ratio':
                return float(peak_params_grid.GetCellValue(i, 5))
            elif param_name == 'area':
                return float(peak_params_grid.GetCellValue(i, 6))
            elif param_name == 'sigma':
                value = float(peak_params_grid.GetCellValue(i, 7))
                return value / 2.355 if fitting_model in ["Voigt (Area, L/G, \u03c3)",
                        "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"] else value
            elif param_name == 'gamma':
                value = float(peak_params_grid.GetCellValue(i, 8))
                return value / 2 if fitting_model in ["Voigt (Area, L/G, \u03c3)",
                            "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"] else value
            elif param_name == 'fwhm_g':
                return float(peak_params_grid.GetCellValue(i, 9))
            elif param_name == 'skew':
                return float(peak_params_grid.GetCellValue(i, 9))

    return None


import re


def parse_constraints(constraint_str, current_value, peak_params_grid, peak_index, param_name):
    constraint_str = constraint_str.strip()
    small_error = 0.05

    # Pattern to match A+1.5#0.5 format
    pattern = r'^([A-P])([+\-*/])(\d+\.?\d*)#([\d\.]+)$'
    match = re.match(pattern, constraint_str)

    # Pattern for A+2 or A*2 or A/2 or A-2 format
    pattern_simple = r'^([A-P])([+\-*/])(\d+\.?\d*)$'
    match_simple = re.match(pattern_simple, constraint_str)

    if constraint_str in ['Fixed']:
        small_error3 = 0.001
        if param_name in ["L/G", "fraction"]:
            return current_value - 0.5, current_value + 0.5, False
        else:
            return current_value - small_error3, current_value + small_error3, False
    elif match:
        ref_peak, operator, value, delta = match.groups()
        value = float(value)
        delta = float(delta)
        if operator in ['+', '-']:
            return (f"{ref_peak}{operator}{value - delta}", f"{ref_peak}{operator}{value + delta}", True)
        elif operator in ['*', '/']:
            if param_name in ['POSITION','FWHM', 'L/G','fwhm_g']:
                delta_percent = delta
                # return (f"{ref_peak}{operator}{value} - {delta}", f"{ref_peak}{operator}{value} + {delta}", True)
                return (f"{ref_peak}{operator}{value - delta_percent}", f"{ref_peak}{operator}{value + delta_percent}",True)
            else:
                return (f"{ref_peak}{operator}{value-delta}", f"{ref_peak}{operator}{value+delta}", True)

    elif match_simple:
        ref_peak, operator, value = match_simple.groups()
        value = float(value)
        if operator in ['+', '-']:
            return f"{ref_peak}{operator}{value - small_error}", f"{ref_peak}{operator}{value + small_error}", True
        elif operator in ['*', '/']:
            if param_name == 'fwhm_g':
                small_error2 = 0.01
            if param_name == 'skew':
                small_error2 = 0.001
            else:
                small_error2 = 0.0001
            return f"{ref_peak}{operator}{value - small_error2}", f"{ref_peak}{operator}{value + small_error2}", True

    # If it's a simple number or range
    if ',' in constraint_str:
        min_val, max_val = map(float, constraint_str.split(','))
        return min_val, max_val, True
    if ':' in constraint_str:
        min_val, max_val = map(float, constraint_str.split(':'))
        return min_val, max_val, True

    try:
        value = float(constraint_str)
        return value - 0.1, value + 0.1, True
    except ValueError:
        pass

    # If we can't parse it, return the current value with a small range
    return current_value - 0.1, current_value + 0.1, True


def evaluate_constraint(constraint, peak_params_grid, param_name, current_value):
    if isinstance(constraint, (int, float)):
        return constraint
    if constraint is None:
        return None

    # Handle the case A+1.5 or A*1.5 or A/1.5 or A-1.5
    match = re.match(r'([A-J])([+\-*/])(-?\d+\.?\d*)', constraint)
    if match:
        peak, op, value = match.groups()
        peak_value = get_peak_value(peak_params_grid, peak, param_name)
        if peak_value is not None:
            value = float(value)
            if op == '+':
                return peak_value + value
            elif op == '-':
                return peak_value - value
            elif op == '*':
                return peak_value * value
            elif op == '/':
                return peak_value / value if value != 0 else current_value

    # Handle simple numeric constraints
    try:
        return float(constraint)
    except ValueError:
        return current_value



# WHERE IS IT USED???
def format_sheet_name2(sheet_name):
    # Regular expression to separate element and electron shell
    match = re.match(r'([A-Z][a-z]*)(\d+[spdfg])', sheet_name)
    if match:
        element, shell = match.groups()
        return f"{element} {shell}"
    else:
        return sheet_name  # Return original if it doesn't match the expected format










