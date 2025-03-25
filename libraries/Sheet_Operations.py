# sheet_operations.py

import wx
import wx.grid
import numpy as np
from libraries.Peak_Functions import OtherCalc

from libraries.Utilities import _clear_peak_params_grid

# In Sheet_Operations.py
def on_sheet_selected(window, event):
    # Extract the selected sheet name
    if isinstance(event, str):
        selected_sheet = event
    else:
        selected_sheet = window.sheet_combobox.GetValue()

    # Update BE correction from BEcorrections if available
    if 'BEcorrections' in window.Data:
        # Extract row number from sheet name
        import re
        match = re.search(r'(\d+)$', selected_sheet)
        if match:
            sample_row = match.group(1)
            if sample_row in window.Data['BEcorrections']:
                correction = window.Data['BEcorrections'][sample_row]
                if correction != window.be_correction:
                    window.be_correction = correction
                    window.be_correction_spinbox.SetValue(correction)
        else:
            # If no numeric suffix, check row 0
            if "0" in window.Data['BEcorrections']:
                correction = window.Data['BEcorrections']["0"]
                if correction != window.be_correction:
                    window.be_correction = correction
                    window.be_correction_spinbox.SetValue(correction)


    if selected_sheet:
        # Reinitialize peak count
        window.peak_count = 0
        window.bg_min_energy = None
        window.bg_max_energy = None

        window.remove_cross_from_peak()

        # Reset RSD value
        if 'fit_results' in window.__dict__:
            window.fit_results['rsd'] = None

        # Reset RSD text in PlotManager
        if hasattr(window.plot_manager, 'rsd_text') and window.plot_manager.rsd_text:
            window.plot_manager.rsd_text.remove()
            window.plot_manager.rsd_text = None

        # Check if there's data for the selected sheet in window.Data
        if selected_sheet in window.Data['Core levels']:
            core_level_data = window.Data['Core levels'][selected_sheet]

            # if 'Background' in window.Data['Core levels'][selected_sheet] and 'Bkg Y' in \
            #         window.Data['Core levels'][selected_sheet]['Background']:
            #     window.background = np.array(window.Data['Core levels'][selected_sheet]['Background']['Bkg Y'])
            # else:
            #     window.background = np.zeros_like(window.Data['Core levels'][selected_sheet]['Raw Data'])

            # # Apply BE correction when loading x-values
            # if hasattr(window, 'be_correction') and window.be_correction != 0:
            #     # Create a corrected copy of the x-values for display purposes
            #     window.x_values = np.array(core_level_data['B.E.']) + window.be_correction
            # else:
            #     window.x_values = np.array(core_level_data['B.E.'])
            #
            # window.y_values = np.array(core_level_data['Raw Data'])

            if 'Background' in window.Data['Core levels'][selected_sheet]:
                if 'Bkg Y' not in window.Data['Core levels'][selected_sheet]['Background']:
                    window.Data['Core levels'][selected_sheet]['Background']['Bkg Y'] = \
                    window.Data['Core levels'][selected_sheet]['Raw Data']
                window.background = np.array(window.Data['Core levels'][selected_sheet]['Background']['Bkg Y'])
            else:
                window.Data['Core levels'][selected_sheet]['Background'] = {
                    'Bkg Y': window.Data['Core levels'][selected_sheet]['Raw Data']}
                window.background = np.array(window.Data['Core levels'][selected_sheet]['Raw Data'])

            # Check if there's fitting data available
            if 'Fitting' in core_level_data and 'Peaks' in core_level_data['Fitting']:
                peaks = core_level_data['Fitting']['Peaks']

                # Update peak count
                window.peak_count = len(peaks)

                # Adjust the number of rows in the grid
                required_rows = window.peak_count * 2
                current_rows = window.peak_params_grid.GetNumberRows()

                if current_rows < required_rows:
                    window.peak_params_grid.AppendRows(required_rows - current_rows)
                elif current_rows > required_rows:
                    window.peak_params_grid.DeleteRows(required_rows, current_rows - required_rows)

                # Clear existing data in the peak_params_grid
                window.peak_params_grid.ClearGrid()

                # Populate the grid with peak data
                for i, (peak_label, peak_data) in enumerate(peaks.items()):
                    row = i * 2  # Each peak uses two rows

                    # Set peak data
                    window.peak_params_grid.SetCellValue(row, 0, chr(65 + i))  # A, B, C, etc.
                    window.peak_params_grid.SetCellValue(row, 1, peak_label)

                    # corrected_position = peak_data.get('Position', 0) + window.be_correction
                    # window.peak_params_grid.SetCellValue(row, 2, f"{corrected_position:.2f}")
                    # Use the position directly from peak_data, which is already corrected
                    window.peak_params_grid.SetCellValue(row, 2, f"{peak_data.get('Position', 'N/A'):.2f}")

                    window.peak_params_grid.SetCellValue(row, 3, f"{peak_data.get('Height', '0')}")
                    window.peak_params_grid.SetCellValue(row, 4, f"{peak_data.get('FWHM', '0')}")
                    window.peak_params_grid.SetCellValue(row, 5, f"{peak_data.get('L/G', '0')}")
                    window.peak_params_grid.SetCellValue(row, 6, f"{peak_data.get('Area', '0')}")
                    window.peak_params_grid.SetCellValue(row, 7, f"{peak_data.get('Sigma', '0')}")
                    window.peak_params_grid.SetCellValue(row, 8, f"{peak_data.get('Gamma', '0')}")
                    window.peak_params_grid.SetCellValue(row, 9, f"{peak_data.get('Skew', '0')}")
                    window.peak_params_grid.SetCellValue(row, 13, f"{peak_data.get('Fitting Model', '0')}")
                    window.peak_params_grid.SetCellValue(row, 14, f"{peak_data.get('Bkg Type', '0')}")
                    window.peak_params_grid.SetCellValue(row, 15, f"{peak_data.get('Bkg Low', '0')}")
                    window.peak_params_grid.SetCellValue(row, 16, f"{peak_data.get('Bkg High', '0')}")
                    window.peak_params_grid.SetCellValue(row, 17, f"{peak_data.get('Bkg Offset Low', '0')}")
                    window.peak_params_grid.SetCellValue(row, 18, f"{peak_data.get('Bkg Offset High', '0')}")

                    # Set constraints if available
                    if 'Constraints' in peak_data:
                        constraints = peak_data['Constraints']
                        window.peak_params_grid.SetCellValue(row + 1, 2, str(constraints.get('Position', '1:1200')))
                        window.peak_params_grid.SetCellValue(row + 1, 3, str(constraints.get('Height', '1:1e7')))
                        window.peak_params_grid.SetCellValue(row + 1, 4, str(constraints.get('FWHM', '0.4:3')))
                        window.peak_params_grid.SetCellValue(row + 1, 5, str(constraints.get('L/G', '10:90')))
                        window.peak_params_grid.SetCellValue(row + 1, 6, str(constraints.get('Area', '1:1e7')))
                        window.peak_params_grid.SetCellValue(row + 1, 7, str(constraints.get('Sigma', '0.01:1')))
                        window.peak_params_grid.SetCellValue(row + 1, 8, str(constraints.get('Gamma', '0.01:1')))
                        window.peak_params_grid.SetCellValue(row + 1, 9, str(constraints.get('Skew', '0.01:2')))

                    # Set background color for constraint rows
                    for col in range(window.peak_params_grid.GetNumberCols()+1):
                        window.peak_params_grid.SetCellBackgroundColour(row + 1, col-1, wx.Colour(200,245,228))
                        window.peak_params_grid.SetCellBackgroundColour(row, col - 1, wx.WHITE)

                    for col in [10,11,12]:  # Columns for Area, sigma and gamma
                        window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(27, 140, 60))
                    for col in [0,1,2]:  # Columns for Area, sigma and gamma
                        window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                        window.peak_params_grid.SetCellTextColour(row+1, col, wx.Colour(0, 0, 0))

                    window.selected_fitting_method = window.peak_params_grid.GetCellValue(row, 13)
                    # Set background color for Height, FWHM, and L/G ratio cells if Voigt function
                    if window.selected_fitting_method == "Voigt (Area, L/G, \u03c3)":
                        for col in [3,4,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [5,6,7]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                    elif window.selected_fitting_method == "Voigt (Area, L/G, \u03c3, S)":
                        for col in [3,4,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [5,6,7,9]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                    elif window.selected_fitting_method in ["Voigt (Area, \u03c3, \u03B3)",
                                        "ExpGauss.(Area, \u03c3, \u03b3)"]:
                        for col in [3,4,5]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [6,7,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                    elif window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)"]:
                        for col in [3,5]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [4,6,7,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                    elif window.selected_fitting_method in ["LA (Area, \u03c3/\u03b3, \u03b3)"]: # LA (Area, \u03c3/\u03b3, \u03b3)
                        for col in [3,7]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [4,5,6,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                    elif window.selected_fitting_method in ["Pseudo-Voigt (Area)", "GL (Area)", "SGL (Area)"]:
                        for col in [3]:  # Height
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [7, 8]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [4,5,6]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                    elif window.selected_fitting_method == "D-parameter":
                        for col in [2]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [4,5,6,8]:  # Columns for Height, FWHM
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                    else:
                        print("Fitting method not recognized")
                        for col in [6]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [7, 8]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                        for col in [3,4,5]:  # Columns for Height, FWHM, L/G ratio
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                        for col in [9]:  # Columns for Area, sigma and gamma
                            window.peak_params_grid.SetCellValue(row, col, "0")
                            window.peak_params_grid.SetCellValue(row + 1, col, "0")
                            window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                            window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))

                # # Update background information if available
                # if 'Background' in core_level_data:
                #     bg_data = core_level_data['Background']
                #     window.bg_min_energy = bg_data.get('Bkg Low', None)
                #     window.bg_max_energy = bg_data.get('Bkg High', None)
                #     window.background_method = bg_data.get('Bkg Type', 'N/A')
                #     window.offset_l = bg_data.get('Bkg Offset Low', 0)
                #     window.offset_h = bg_data.get('Bkg Offset High', 0)
                # Update background information if available
                if 'Background' in core_level_data:
                    bg_data = core_level_data['Background']
                    window.bg_min_energy = bg_data.get('Bkg Low', '')
                    window.bg_max_energy = bg_data.get('Bkg High', '')
                    window.background_method = bg_data.get('Bkg Type', '')  # Change 'N/A' to empty string
                    window.offset_l = bg_data.get('Bkg Offset Low', '')
                    window.offset_h = bg_data.get('Bkg Offset High', '')

                    # Ensure background fields have empty strings when not defined
                    if 'Bkg Type' not in bg_data:
                        core_level_data['Background']['Bkg Type'] = ""
                    if 'Bkg Low' not in bg_data:
                        core_level_data['Background']['Bkg Low'] = ""
                    if 'Bkg High' not in bg_data:
                        core_level_data['Background']['Bkg High'] = ""
                    if 'Bkg Offset Low' not in bg_data:
                        core_level_data['Background']['Bkg Offset Low'] = ""
                    if 'Bkg Offset High' not in bg_data:
                        core_level_data['Background']['Bkg Offset High'] = ""

            else:
                # If no fitting data, ensure the grid is empty
                _clear_peak_params_grid(window)
        else:
            # If no data for this sheet, ensure the grid is empty
            _clear_peak_params_grid(window)

        # Apply choice editors to the fitting model column
        window.set_model_choice_editors(window)

        # Refresh the grid display
        window.peak_params_grid.ForceRefresh()

        # Update the plot
        window.plot_manager.plot_data(window)  # Always plot raw data first
        if window.show_fit and window.peak_params_grid.GetNumberRows() > 0:
            window.clear_and_replot()  # Add fit and residuals if show_fit is True

        window.plot_config.update_plot_limits(window, selected_sheet)
        window.plot_manager.update_legend(window)
        window.update_ratios()
        # window.update_checkbox_visuals()

        # print(f"Selected sheet: {selected_sheet}, Peak count: {window.peak_count}, Show fit: {window.show_fit}")

    # Update the combobox selection if a string was passed directly
    if isinstance(event, str):
        window.sheet_combobox.SetValue(selected_sheet)

    window.update_checkboxes_from_data()

    # Update FileManager cell highlight if open
    if hasattr(window, 'file_manager') and window.file_manager is not None:
        try:
            selected_sheet = window.sheet_combobox.GetValue()
            window.file_manager.highlight_current_sheet(selected_sheet)
        except (RuntimeError, Exception):
            pass


def on_grid_left_click(window, event):
    if event.GetCol() == 7:  # Checkbox column
        row = event.GetRow()
        current_value = window.results_grid.GetCellValue(row, 7)
        new_value = '1' if current_value == '0' else '0'

        # Update grid
        window.results_grid.SetCellValue(row, 7, new_value)

        # Update Data structure
        peak_label = chr(65 + row)  # A, B, C, ...
        if 'Results' in window.Data and 'Peak' in window.Data['Results'] and peak_label in window.Data['Results'][
            'Peak']:
            window.Data['Results']['Peak'][peak_label]['Checkbox'] = new_value

        # window.results_grid.RefreshCell(row, 7)
        window.results_grid.ForceRefresh()
        window.update_atomic_percentages()

    event.Skip()


class CheckboxRenderer(wx.grid.GridCellRenderer):
    def __init__(self):
        wx.grid.GridCellRenderer.__init__(self)

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        dc.SetBrush(wx.Brush(wx.WHITE, wx.SOLID))
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.DrawRectangle(rect)

        value = grid.GetCellValue(row, col)
        if value == '1':
            flag = wx.CONTROL_CHECKED
        else:
            flag = 0

        wx.RendererNative.Get().DrawCheckBox(grid, dc, rect, flag)

    def GetBestSize(self, grid, attr, dc, row, col):
        return wx.Size(20, 20)

    def Clone(self):
        return CheckboxRenderer()