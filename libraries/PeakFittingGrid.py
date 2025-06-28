# peakFittingGrid.py
import numpy as np
from libraries.Save import save_state
import wx.grid
import wx.adv
import re

from libraries.Peak_Functions import PeakFunctions
from libraries.Sheet_Operations import on_sheet_selected


# Peak parameter methods
class PeakFittingGrid:
    def __init__(self, main_window):
        self.window = main_window

    # Peak parameter methods
    def add_peak_params(self):
        if hasattr(self.window, 'fitting_window'):
            self.window.selected_fitting_method = self.window.fitting_window.model_combobox.GetValue()
            print(f'Fitting method: {self.window.selected_fitting_method}')
        save_state(self.window)
        sheet_name = self.window.sheet_combobox.GetValue()

        if self.window.bg_min_energy is None or self.window.bg_max_energy is None:
            wx.MessageBox("Please create a background first.", "No Background", wx.OK | wx.ICON_WARNING)
            return None

        num_peaks = self.window.peak_params_grid.GetNumberRows() // 2

        # Update bg_min_energy and bg_max_energy from window.Data
        if sheet_name in self.window.Data['Core levels'] and 'Background' in self.window.Data['Core levels'][sheet_name]:
            background_data = self.window.Data['Core levels'][sheet_name]['Background']
            self.window.bg_min_energy = background_data.get('Bkg Low')
            self.window.bg_max_energy = background_data.get('Bkg High')

        # Ensure bg_min_energy and bg_max_energy are not None
        if self.window.bg_min_energy is None or self.window.bg_max_energy is None:
            wx.MessageBox("Background range is not set. Please set the background first.", "Warning",
                          wx.OK | wx.ICON_WARNING)
            return

        if num_peaks == 0:
            residual = self.window.y_values - np.array(self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
            peak_y = residual[np.argmax(residual)]
            peak_x = self.window.x_values[np.argmax(residual)]
        else:

            # Call update_overall_fit_and_residuals to get the residuals
            residual = self.window.plot_manager.update_overall_fit_and_residuals(self.window)

            if residual is not None:
                peak_y = residual.max()
                peak_x = self.window.x_values[np.argmax(residual)]
            else:
                # Fallback if residuals couldn't be calculated
                wx.MessageBox("Unable to calculate residuals. Using default peak position.", "Warning",
                              wx.OK | wx.ICON_WARNING)
                peak_y = self.window.y_values.max()
                peak_x = self.window.x_values[np.argmax(self.window.y_values)]

        self.window.peak_count += 1

        # Add new rows to the grid
        self.window.peak_params_grid.AppendRows(2)
        self.window.add_choice_editor_to_new_row(self.window.peak_params_grid, self.window.peak_params_grid.GetNumberRows() - 2)
        row = self.window.peak_params_grid.GetNumberRows() - 2

        # Assign letter IDs
        letter_id = chr(64 + self.window.peak_count)


        # Set values in the grid
        self.window.peak_params_grid.SetCellValue(row, 0, letter_id)
        self.window.peak_params_grid.SetReadOnly(row, 0)
        self.window.peak_params_grid.SetCellValue(row, 1, f"{sheet_name} p{self.window.peak_count}")
        self.window.peak_params_grid.SetCellValue(row, 2, f"{peak_x:.2f}")
        self.window.peak_params_grid.SetCellValue(row, 3, f"{peak_y:.2f}")
        self.window.peak_params_grid.SetCellValue(row, 4, "1.6")
        self.window.peak_params_grid.SetCellValue(row, 5, "20")
        if self.window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.window.peak_params_grid.SetCellValue(row, 6, f"{peak_y * 1.6 * 1.064:.2f}")
        else:
            self.window.peak_params_grid.SetCellValue(row, 6, f"{peak_y * 1.6 * 1.064:.2f}")
        if self.window.selected_fitting_method == "ExpGauss.(Area, \u03c3, \u03b3)":
            self.window.peak_params_grid.SetCellValue(row, 7, "0.3")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '1.2')  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, '0.64')  # skew
        elif self.window.selected_fitting_method in ["LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.window.peak_params_grid.SetCellValue(row, 5, "50")
            self.window.peak_params_grid.SetCellValue(row, 7, "2.7")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '2.7')  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, '0.64')  # skew
        elif self.window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)"]:
            self.window.peak_params_grid.SetCellValue(row, 5, "50")
            self.window.peak_params_grid.SetCellValue(row, 7, "2.7")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '2.7')  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, '0')  # skew
        elif self.window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            self.window.peak_params_grid.SetCellValue(row, 5, "20")
            self.window.peak_params_grid.SetCellValue(row, 7, "1.2")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '0.4')  #
            self.window.peak_params_grid.SetCellValue(row, 9, '0.01')  # skew
        elif self.window.selected_fitting_method == "DS (A, \u03c3, \u03b3)":
            self.window.peak_params_grid.SetCellValue(row, 7, "0.5")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, "0.0")  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, "0.0")  # skew
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "-0.1:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, "-0.2:0.2")
        elif self.window.selected_fitting_method == "DS*G (A, \u03c3, \u03b3, S)":
            # Calculate proper area based on peak shape
            x_range = np.linspace(-10, 10, 1000)  # Temporary x range
            y_values = PeakFunctions.DS_G(x_range, 0, 1.0, 0.4, 0.0, 0.8)  # Use unit amplitude
            max_height = np.max(y_values)
            area = peak_y / max_height  # Scale area appropriately
            self.window.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")
            self.window.peak_params_grid.SetCellValue(row, 7, "0.8")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, "0.4")  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, "0.0")  # skew
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.1:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0:0.2")
        elif self.window.selected_fitting_method in ["D-parameter"]:
            self.window.peak_params_grid.SetCellValue(row, 5, "2")
            self.window.peak_params_grid.SetCellValue(row, 7, "1")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '1')  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, '7')  # skew
        else:
            self.window.peak_params_grid.SetCellValue(row, 7, "1")  # sigma
            self.window.peak_params_grid.SetCellValue(row, 8, '0.15')  # gamma
            self.window.peak_params_grid.SetCellValue(row, 9, "0.64")  # Default value for skewed
        self.window.peak_params_grid.SetCellValue(row, 10, '')
        self.window.peak_params_grid.SetCellValue(row, 11, '')
        self.window.peak_params_grid.SetCellValue(row, 12, '')  # Split, initially empty
        self.window.peak_params_grid.SetCellValue(row, 13, self.window.selected_fitting_method)  # Fitting Model
        self.window.peak_params_grid.SetCellValue(row, 14, self.window.background_method)  # Bkg Type
        self.window.peak_params_grid.SetCellValue(row, 15,
                                           f"{self.window.bg_min_energy:.2f}" if self.window.bg_min_energy is not None else "")  # Bkg Low
        self.window.peak_params_grid.SetCellValue(row, 16,
                                           f"{self.window.bg_max_energy:.2f}" if self.window.bg_max_energy is not None else "")  # Bkg High
        self.window.peak_params_grid.SetCellValue(row, 17, f"{float(self.window.offset_l):.2f}")  # Bkg Offset Low
        self.window.peak_params_grid.SetCellValue(row, 18, f"{self.window.offset_h:.2f}")  # Bkg Offset High

        # Set position constraint to background range
        position_constraint = f"{self.window.bg_min_energy:.2f},{self.window.bg_max_energy:.2f}"
        self.window.peak_params_grid.SetCellValue(row + 1, 2, position_constraint)
        self.window.peak_params_grid.SetCellValue(row + 1, 3, "1:1e7")
        self.window.peak_params_grid.SetCellValue(row + 1, 4, "0.3:3.5")
        self.window.peak_params_grid.SetCellValue(row + 1, 5, "2:80")
        self.window.peak_params_grid.SetCellValue(row + 1, 6, "1:1e7")
        self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:3")
        self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.3:3")
        self.window.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')
        if self.window.selected_fitting_method == "ExpGauss.(Area, \u03c3, \u03b3)":
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.01:1")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.01:3")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        elif self.window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.01:10")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.01:10")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        elif self.window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            self.window.peak_params_grid.SetCellValue(row + 1, 5, "15:85")
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.2:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.2:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, '0.01:0.7')  # skew
        elif self.window.selected_fitting_method == "DS (A, \u03c3, \u03b3)":
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "-0.1:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, "-0.2:0.2")
        elif self.window.selected_fitting_method == "DS*G (A, \u03c3, \u03b3, S)":
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.1:1.5")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0:0.2")
        else:
            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:3")
            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.3:3")
            self.window.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        self.window.peak_params_grid.ForceRefresh()

        # Set constraint values
        self.window.peak_params_grid.SetReadOnly(row + 1, 0)
        for col in range(self.window.peak_params_grid.GetNumberCols()+1):  # Assuming you have 15 columns in total
            # self.window.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(230, 230, 230))
            self.window.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200,245,228))

        for col in [10, 11, 12]:  # Columns for Area, sigma and gamma
            self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(27, 140, 60))

        # Set background color for Height, FWHM, and L/G ratio cells if Voigt function
        if self.window.selected_fitting_method == "Voigt (Area, L/G, \u03c3)":
            for col in [3, 4, 8]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [5, 6, 7]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.window.peak_params_grid.SetCellValue(row, col, "0.1")
                self.window.peak_params_grid.SetCellValue(row + 1, col, "0.1:1")
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.window.selected_fitting_method == "Voigt (Area, L/G, \u03c3, S)":
            for col in [3, 4, 8]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [5, 6, 7, 9]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.window.selected_fitting_method in ["Voigt (Area, \u03c3, \u03b3)", "ExpGauss.(Area, \u03c3, \u03b3)"]:
            for col in [3,4, 5, 9]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [6, 7, 8]:  # Columns for Height, FWHM
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.window.peak_params_grid.SetCellValue(row, col, "0.1")
                self.window.peak_params_grid.SetCellValue(row + 1, col, "0.1:1")
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.window.selected_fitting_method == "DS (A, \u03c3, \u03b3)":
            for col in [3, 4, 5]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
            for col in [6, 7, 8, 9]:  # Columns for Area, sigma, gamma, skew
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.window.selected_fitting_method == "DS*G (A, \u03c3, \u03b3, S)":
            for col in [3, 4, 5]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
            for col in [6, 7, 8, 9]:  # Columns for Area, sigma, gamma, skew
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)" ]:
            for col in [3, 5]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 6, 7, 8]:  # Columns for Height, FWHM
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.window.peak_params_grid.SetCellValue(row, col, "0.1")
                self.window.peak_params_grid.SetCellValue(row + 1, col, "0.1:1")
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.window.selected_fitting_method in ["LA (Area, \u03c3/\u03b3, \u03b3)"]:
            for col in [3,7]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6, 8]:  # Columns for Height, FWHM
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.window.peak_params_grid.SetCellValue(row, col, "0.1")
                self.window.peak_params_grid.SetCellValue(row + 1, col, "0.1:1")
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.window.selected_fitting_method in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            for col in [3,7]:  # Columns for Height, FWHM, L/G ratio
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6, 8, 9]:  # Columns for Height, FWHM
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.window.selected_fitting_method in  ["Pseudo-Voigt (Area)", "GL (Area)", "SGL (Area)"]:
            for col in [3]:  # Height
                # self.window.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(200,245,228))
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [7,8, 9]:  # Columns for Area, sigma and gamma
                # self.window.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                # self.window.peak_params_grid.SetCellValue(row , col, "0")
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(255,255,255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        else:
            for col in [6]:  # Columns for Area, sigma and gamma
                # self.window.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [7,8, 9]:  # Columns for Area, sigma and gamma
                # self.window.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                # self.window.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.window.peak_params_grid.SetCellTextColour(row , col, wx.Colour(255,255,255))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [3, 4, 5]:  # Columns for Height, FWHM, L/G ratio
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))

        # Set selected_peak_index to the index of the new peak
        self.window.selected_peak_index = num_peaks

        # Update the Data structure with the new peak information
        if 'Fitting' not in self.window.Data['Core levels'][sheet_name]:
            self.window.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.window.Data['Core levels'][sheet_name]['Fitting']:
            self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}



        if self.window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': peak_y * 1.6 * 1.064,
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.01,
                'Fitting Model': self.window.selected_fitting_method,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "2:80",
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:2"
                }
            }
        elif self.window.selected_fitting_method in ["DS (A, \u03c3, \u03b3)"]:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.0,
                'L/G': 20,
                'Area': peak_y * 1.0 * 1.0,
                'Sigma': 0.5,
                'Gamma': 0.0,
                'Skew': 0.0,
                'Fitting Model': self.window.selected_fitting_method,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",
                    'Area': '1:1e7',
                    'Sigma': "0.3:1.5",
                    'Gamma': "0.1:1.5",
                    'Skew': "-0.2:0.2"
                }
            }
        elif self.window.selected_fitting_method in ["DS*G (A, \u03c3, \u03b3, S)"]:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.0,
                'L/G': 20,
                'Area': peak_y * 1.0 * 1.0,
                'Sigma': 0.5,
                'Gamma': 0.5,
                'Skew': 0.0,
                'Fitting Model': self.window.selected_fitting_method,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",
                    'Area': '1:1e7',
                    'Sigma': "0.3:1.5",
                    'Gamma': "0.1:1.5",
                    'Skew': "0:0.2"
                }
            }
        else:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': peak_y * 1.6 * 1.064,
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.64,
                'Fitting Model': self.window.selected_fitting_method,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "2:80",
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.00:2"
                }
            }

        if self.window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)","LA (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_data.update({
                'L/G': 50,  # Default L/G ratio for LA model
                'Sigma': 2.75,
                'Gamma': 2.75,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",  # Wider range for L/G in LA model
                    'Area': '1:1e7',
                    'Sigma': "0.01:10",
                    'Gamma': "0.01:10"
                }
            })
        elif self.window.selected_fitting_method in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_data.update({
                'L/G': 50,  # Default L/G ratio for LA model
                'Sigma': 2.75,
                'Gamma': 2.75,
                'Skew': 0.64,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",  # Wider range for L/G in LA model
                    'Area': '1:1e7',
                    'Sigma': "0.01:4",
                    'Gamma': "0.01:4",
                    'Skew': "0.01:2"
                }
            })
        elif self.window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
            peak_data.update({
                'Sigma': 1,
                'Gamma': 0.5,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3"
                }
            })
        elif self.window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            peak_data.update({
                'Sigma': 1,
                'Gamma': 0.5,
                'skew': 0.01,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:0.7"
                }
            })
        elif self.window.selected_fitting_method in ["SGL (Area)"]:
            # Value required to calculate area
            fwhm = 1.6
            fraction = 20
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = fwhm / 2
            sgl_area = peak_y * ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (fraction / 100) * np.pi * gamma)
            peak_data.update({
                'Area': sgl_area,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:0.7"
                }
            })

        self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks'][sheet_name + f" p{self.window.peak_count}"] = peak_data
        # print(self.window.Data)
        self.window.show_hide_vlines()

        # Call the method to clear and replot everything
        self.window.clear_and_replot()

        return self.window.peak_count - 1  # Return the index of the added peak

    def on_peak_params_cell_changed(self, event):
        row = event.GetRow()
        col = event.GetCol()
        new_value = self.window.peak_params_grid.GetCellValue(row, col)
        sheet_name = self.window.sheet_combobox.GetValue()
        peak_index = row // 2

        # # Define default constraint values
        # sheet_name = self.window.sheet_combobox.GetValue()
        # x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
        # new_value = f"{min(x_values):.2f}:{max(x_values):.2f}"
        default_constraints = {
            2: '1:1000',  # Position
            3: '1:1e7',  # Height
            4: '0.3:3.5',  # FWHM
            5: '5:80',  # L/G
            6: '1:1e7',  # Area
            7: '0.3:3',  # Sigma
            8: '0.3:3',  # Gamma
            9: '0.01:2' # Skew
        }
        # Check each constraint cell
        for col_idx in range(2, 10):
            if not self.window.peak_params_grid.GetCellValue(row, col_idx).strip():
                # If empty, set default constraint
                self.window.peak_params_grid.SetCellValue(row, col_idx, default_constraints[col_idx])

        # Also update constraints in Data structure
        peak_index = row // 2
        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name]:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            if peaks and peak_index < len(list(peaks.keys())):
                peak_label = list(peaks.keys())[peak_index]
                if 'Constraints' not in peaks[peak_label]:
                    peaks[peak_label]['Constraints'] = {}

                # Map column indices to constraint names
                constraint_names = {
                    2: 'Position', 3: 'Height', 4: 'FWHM', 5: 'L/G',
                    6: 'Area', 7: 'Sigma', 8: 'Gamma', 9: 'Skew'
                }

                # Update any empty constraints
                for col_idx in range(2, 10):
                    constraint_name = constraint_names[col_idx]
                    if not peaks[peak_label]['Constraints'].get(constraint_name, ''):
                        peaks[peak_label]['Constraints'][constraint_name] = default_constraints[col_idx]



        if col == 1:  # Peak label column
            # Check for duplicate names
            existing_names = []
            for i in range(0, self.window.peak_params_grid.GetNumberRows(), 2):
                if i != row:  # Skip current row
                    existing_names.append(self.window.peak_params_grid.GetCellValue(i, 1))

            if new_value in existing_names:
                wx.MessageBox(f"Peak name '{new_value}' already exists. Cannot have duplicate peak names.",
                              "Duplicate Peak Name", wx.OK | wx.ICON_ERROR)
                event.Veto()
                return
        elif col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1:  # Constraint rows
            print('It is a constraint row')

            # Check if this is a constraint row and the value contains "="
            if row % 2 == 1 and col in [2, 3, 4, 5, 6, 7, 8, 9] and "=" in new_value:
                print('It is a constraint row with "="')
                # Remove the "=" from the cell
                self.window.peak_params_grid.SetCellValue(row, col, new_value.replace("=", ""))

                # Import and call the propagate_constraint function
                from libraries.Utilities import propagate_constraint
                propagate_constraint(self.window, row, col)
                return  # Skip the rest of the function
            elif new_value.lower() in ['fi', 'fix', 'fixe', 'fixed']:
                print('It is a constraint row with "Fixed"')
                new_value = 'Fixed'
                # Make sure the "Fixed" value gets saved right away to the Data structure
                constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
                constraint_key = constraint_keys[col - 2]

                # Save to the Data structure immediately
                if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name]:
                    peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                    correct_peak_key = list(peaks.keys())[peak_index]

                    if 'Constraints' not in peaks[correct_peak_key]:
                        peaks[correct_peak_key]['Constraints'] = {}

                    # Save the "Fixed" value directly
                    peaks[correct_peak_key]['Constraints'][constraint_key] = new_value

                # Update the cell value
                self.window.peak_params_grid.SetCellValue(row, col, new_value)
                event.Skip()
                return
            elif new_value == 'F':
                new_value = 'F*1'
                self.window.peak_params_grid.SetCellValue(row, col, new_value)
                return
            elif new_value.startswith('#'):
                # Check if '#' is not followed by at least one digit
                if len(new_value) == 1 or not new_value[1:].replace('.', '', 1).isdigit():
                    wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()  # Veto the event if '#' is on its own or not followed by a valid number
                    return

                # If '#' is followed by a valid number, proceed with the calculation
                peak_value = float(self.window.peak_params_grid.GetCellValue(row - 1, col))
                new_value = str(round(peak_value - float(new_value[1:]), 2)) + ':' + str(
                    round(peak_value + float(new_value[1:]), 2))
                self.window.peak_params_grid.SetCellValue(row, col, new_value)

                # Save to Data structure before returning
                sheet_name = self.window.sheet_combobox.GetValue()
                peak_index = row // 2
                if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name]:
                    peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                    if peaks and peak_index < len(list(peaks.keys())):
                        peak_label = list(peaks.keys())[peak_index]
                        if 'Constraints' not in peaks[peak_label]:
                            peaks[peak_label]['Constraints'] = {}

                        # Map column indices to constraint names
                        constraint_names = {
                            2: 'Position', 3: 'Height', 4: 'FWHM', 5: 'L/G',
                            6: 'Area', 7: 'Sigma', 8: 'Gamma', 9: 'Skew'
                        }

                        peaks[peak_label]['Constraints'][constraint_names[col]] = new_value

                return
            elif new_value.startswith(('+', '/', '*')):
                wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                              wx.OK | wx.ICON_ERROR)
                event.Veto()
                return
            # Add validation for expressions starting with '-' and not containing ':'
            elif new_value.startswith('-') and ':' not in new_value:
                wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                              wx.OK | wx.ICON_ERROR)
                event.Veto()
                return

            elif ':' in new_value:
                parts = new_value.split(':')
                if len(parts) != 2:
                    wx.MessageBox("Constraint with ':' must have exactly one colon", "Wrong Value",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                try:
                    # Check if both parts are valid numbers
                    float(parts[0].strip())
                    float(parts[1].strip())
                except ValueError:
                    wx.MessageBox("Constraint with ':' must have numbers before and after the colon", "Wrong Value",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

            # Pattern to match all possible formats
            pattern = r'^([A-P])([+\-*/])(\d+\.?\d*)(?:#(\d+\.?\d*))?$'
            match = re.match(pattern, new_value)
            print(f'Checking if it is an empty string: {new_value}')
            if not new_value:  # If empty string
                print(f'It is an empty string: {new_value}')
                if col == 2:
                    sheet_name = self.window.sheet_combobox.GetValue()
                    x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
                    new_value = f"{min(x_values):.2f}:{max(x_values):.2f}"
                elif col == 4:
                    new_value = "0.3:3.5"
                elif col == 5:
                    new_value = "1:80"
                elif col == 6:
                    new_value = "1:1e7"
                elif col == 7:
                    new_value = "0.2:3"
                elif col == 8:
                    new_value = "0.2:3"
                elif col == 9:
                    new_value = "0.1:1"
                self.window.peak_params_grid.SetCellValue(row, col, new_value)

            # Save ALL constraint changes to Data structure
            sheet_name = self.window.sheet_combobox.GetValue()
            peak_index = row // 2
            if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][
                sheet_name]:
                peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                if peaks and peak_index < len(list(peaks.keys())):
                    peak_label = list(peaks.keys())[peak_index]
                    if 'Constraints' not in peaks[peak_label]:
                        peaks[peak_label]['Constraints'] = {}

                    # Map column indices to constraint names
                    constraint_names = {
                        2: 'Position', 3: 'Height', 4: 'FWHM', 5: 'L/G',
                        6: 'Area', 7: 'Sigma', 8: 'Gamma', 9: 'Skew'
                    }

                    # Save the constraint value
                    if col in constraint_names:
                        peaks[peak_label]['Constraints'][constraint_names[col]] = new_value
                        print(f"Saved constraint {constraint_names[col]} = {new_value} for peak {peak_label}")

            if match:
                referenced_peak = match.group(1)
                letter_index = ord(referenced_peak) - 65
                current_peak = row // 2

                if letter_index == current_peak:
                    wx.MessageBox(f"Cannot reference the same peak ({referenced_peak}).", "Invalid self.window Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                if letter_index * 2 >= self.window.peak_params_grid.GetNumberRows():
                    wx.MessageBox(f"Peak {referenced_peak} does not exist.", "Invalid Peak Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

            elif new_value.upper() in 'ABCDEFGHIJKLMNOP':
                letter_index = ord(new_value.upper()) - 65
                current_peak = row // 2

                if letter_index == current_peak:
                    wx.MessageBox(f"Cannot reference the same peak ({new_value.upper()}).", "Invalid self.window Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                if letter_index * 2 >= self.window.peak_params_grid.GetNumberRows():
                    wx.MessageBox(f"Peak {new_value.upper()} does not exist.", "Invalid Peak Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return
        elif col in [0, 10, 11, 12]:
            event.Veto()
            return
        elif col not in [13, 14] and row % 2 == 1:  # Constraint row
            if not new_value:  # If the cell is empty
                new_value = default_constraints.get(col, '')
                self.window.peak_params_grid.SetCellValue(row, col, new_value)
        # Allow only numeric input for specific columns in non-constraint rows
        elif col not in [1, 13, 14] and row % 2 == 0:
            try:
                float(new_value)  # This will handle integers, floats, and scientific notation
            except ValueError:
                event.Veto()
                return


        if col == 2 and new_value.upper() in 'ABCDEFGHIJKLMNOP':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                new_value = "0:1000"  # Default value
            else:
                current_split = float(self.window.peak_params_grid.GetCellValue(row-1, 12))
                print(f'Current split: {current_split}')
                ref_split = float(self.window.peak_params_grid.GetCellValue(letter_index * 2, 12))
                split_diff = current_split - ref_split
                if split_diff != 0:
                    new_value = f"{new_value.upper()}+{split_diff:.2f}#0.1"
                else:
                    new_value = new_value.upper() + '*1'
        elif col == 6 and new_value.upper() in 'ABCDEFGHIJKLMNOP':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                new_value = "1:1e7"  # Default value
            else:
                current_ratio = float(self.window.peak_params_grid.GetCellValue(row - 1, 11))
                ref_ratio = float(self.window.peak_params_grid.GetCellValue(letter_index * 2, 11))
                ratio = current_ratio / ref_ratio
                if ratio != 1:
                    new_value = f"{new_value.upper()}*{ratio:.2f}#0.01"
                else:
                    new_value = new_value.upper() + '*1'
        elif new_value.lower() in 'abcdefghijklmnop':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                if col == 2:
                    sheet_name = self.window.sheet_combobox.GetValue()
                    x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
                    new_value = f"{min(x_values):.2f}:{max(x_values):.2f}"
                elif col ==4:
                    new_value = "0.3:3.5"
                elif col ==5:
                    new_value = "1:80"
                elif col ==6:
                    new_value = "1:1e7"
                elif col ==7:
                    new_value = "0.2:3"
                elif col ==8:
                    new_value = "0.2:3"
                elif col ==9:
                    new_value = "0.1:1"
            else:
                new_value = new_value.upper() + '*1'
            self.window.peak_params_grid.SetCellValue(row, col, new_value)

        # Convert lowercase to uppercase in expressions like a*0.5
        if '*' in new_value or '+' in new_value:
            parts = new_value.split('*' if '*' in new_value else '+')
            if len(parts) == 2 and parts[0].lower() in 'abcdefghij':
                parts[0] = parts[0].upper()
                new_value = ('*' if '*' in new_value else '+').join(parts)
                self.window.peak_params_grid.SetCellValue(row, col, new_value)

        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.window.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            peak_keys = list(peaks.keys())

            if peak_index < len(peak_keys):
                correct_peak_key = peak_keys[peak_index]

                if row % 2 == 0:  # Main parameter row
                    if col == 1:  # Label
                        # Update the label while preserving order
                        new_peaks = {}
                        for i, (key, value) in enumerate(peaks.items()):
                            if i == peak_index:
                                new_peaks[new_value] = value
                            else:
                                new_peaks[key] = value
                        self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks
                    elif col == 2:  # Position
                        peaks[correct_peak_key]['Position'] = float(new_value)
                    elif col in [3, 4, 5, 6, 7, 8,9]:  # Height, FWHM, L/G, Area, Sigma, Gamma changed
                        def try_float(value, default=0.0):
                            try:
                                return float(value)
                            except (ValueError, TypeError):
                                return default


                        model = peaks[correct_peak_key]['Fitting Model']
                        height = float(self.window.peak_params_grid.GetCellValue(row, 3))
                        fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                        fraction = float(self.window.peak_params_grid.GetCellValue(row, 5))
                        area = float(self.window.peak_params_grid.GetCellValue(row, 6))
                        sigma = try_float(self.window.peak_params_grid.GetCellValue(row, 7), 0.0)
                        gamma = try_float(self.window.peak_params_grid.GetCellValue(row, 8), 0.0)
                        skew = try_float(self.window.peak_params_grid.GetCellValue(row, 9))

                        if model in ["LA (Area, \u03c3/\u03b3, \u03b3)"]:
                            if col == 5:  # L/G ratio changed
                                gamma = float(self.window.peak_params_grid.GetCellValue(row, 8))
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                            elif col == 7:  # Sigma changed
                                fraction = 100 * sigma / (sigma + gamma)
                                self.window.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                            elif col == 8:  # Gamma changed
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.3f}")
                            elif col == 9:
                                pass
                        elif model in ["LA (Area, \u03c3, \u03b3)"]:
                            if col == 5:  # L/G ratio changed
                                pass
                            elif col == 7:  # Sigma changed
                                fraction = 100 * sigma / (sigma + gamma)
                                self.window.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                            elif col == 8:  # Gamma changed
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.3f}")
                            elif col == 9:
                                pass
                        elif model in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                            if col == 5:  # L/G ratio changed
                                gamma = float(self.window.peak_params_grid.GetCellValue(row, 8))
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                            elif col == 7:  # Sigma changed
                                fraction = 100 * sigma / (sigma + gamma)
                                self.window.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                            elif col == 8:  # Gamma changed
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.3f}")
                            elif col == 9:
                                skew = float(self.window.peak_params_grid.GetCellValue(row, 9))
                                peaks[correct_peak_key]['Skew'] = float(new_value)
                        elif model in ["Voigt (Area, L/G, \u03c3)"]:
                            if col == 5: # L/G ratio changed
                                sigma = float(self.window.peak_params_grid.GetCellValue(row, 7))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 8, f"{gamma:.3f}")
                            elif col == 7:  # Sigma changed
                                fraction = float(self.window.peak_params_grid.GetCellValue(row, 5))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 8, f"{gamma:.3f}")
                            elif col == 8:
                                pass
                        elif model in ["Voigt (Area, L/G, \u03c3, S)"]:
                            if col == 5: # L/G ratio changed
                                sigma = float(self.window.peak_params_grid.GetCellValue(row, 7))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 8, f"{gamma:.3f}")
                            elif col == 7:  # Sigma changed
                                fraction = float(self.window.peak_params_grid.GetCellValue(row, 5))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.window.peak_params_grid.SetCellValue(row, 8, f"{gamma:.3f}")
                            elif col == 8:
                                pass
                            elif col == 9:
                                value = float(self.window.peak_params_grid.GetCellValue(row, 9))
                                if value == 0:
                                    skew = 0.1
                                    self.window.peak_params_grid.SetCellValue(row, 9, f"{skew:.3f}")
                        elif model in ["DS (A, \u03c3, \u03b3)", "DS*G (A, \u03c3, \u03b3, S)"]:
                            # For DS model, sigma and gamma are independent parameters
                            # No need to update other parameters when one changes
                            if col == 5:
                                # L/G ratio isn't relevant for DS model, ignore changes
                                pass
                            elif col == 7:
                                # Sigma changed, no automatic updates needed
                                pass
                            elif col == 8:
                                # Gamma changed, no automatic updates needed
                                pass
                            elif col == 9:
                                value = float(self.window.peak_params_grid.GetCellValue(row, 9))
                                if value == 0:
                                    skew = 0.0
                                    self.window.peak_params_grid.SetCellValue(row, 9, f"{skew:.3f}")
                        elif model in ["GL (Area)"]:
                                # For Gaussian-Lorentzian area-based model
                                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                                height = area / (sigma * np.sqrt(2 * np.pi))
                                self.window.peak_params_grid.SetCellValue(row, 3, f"{height:.2f}")
                        elif model in ["SGL (Area)"]:
                                # For Sum of Gaussian-Lorentzian area-based model
                                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                                gamma = fwhm / 2
                                height = area / ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (
                                            fraction / 100) * np.pi * gamma)
                                self.window.peak_params_grid.SetCellValue(row, 3, f"{height:.2f}")
                        elif model == "D-parameter":
                            return
                        else:
                            # Recalculate area
                            area = self.window.calculate_peak_area(model, height, fwhm, fraction, sigma, gamma,skew)
                            self.window.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")

                        self.window.update_ratios()
                        # Update grid and data
                        peaks[correct_peak_key].update({
                            'Height': round(height, 2),
                            'FWHM': round(fwhm, 2),
                            'L/G': round(fraction, 2),
                            'Area': round(area, 2),
                            'Sigma': round(sigma, 2),
                            'Gamma': round(gamma, 2),
                            'Skew': round(skew, 3)
                        })
                    elif col == 13:  # Fitting Model changed
                        peaks[correct_peak_key]['Fitting Model'] = new_value

                        # Default values based on model type
                        if new_value in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
                            # LA models
                            fraction = 50.0
                            sigma = 2.7
                            gamma = 2.7
                            skew = 0.64
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.01:10")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.01:10")  # Gamma constraint
                        elif new_value in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                            # LA*G model
                            fraction = 50.0
                            sigma = 2.7
                            gamma = 2.7
                            skew = 0.64
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.01:4")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.01:4")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0.01:2")  # Skew constraint
                        elif new_value == "Pseudo-Voigt (Area)":
                            # Pseudo-Voigt
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.15
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "5:80")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "Fixed")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "Fixed")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "Fixed")  # Skew constraint
                        elif new_value in ["Voigt (Area, L/G, \u03c3)"]:
                            # Voigt models with L/G
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.5
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "15:85")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:3")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.3:3")  # Gamma constraint
                        elif new_value in ["Voigt (Area, \u03c3, \u03b3)"]:
                            # Voigt models with separate sigma/gamma
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.5
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:3")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.3:3")  # Gamma constraint
                        elif new_value in ["Voigt (Area, L/G, \u03c3, S)"]:
                            # Skewed Voigt
                            fraction = 20.0
                            sigma = 1.2
                            gamma = 0.4
                            skew = 0.01
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "15:85")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.2:1.5")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.2:1.5")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0.01:0.7")  # Skew constraint
                        elif new_value in ["DS (A, \u03c3, \u03b3)"]:
                            # Doniach-Sunjic model
                            fraction = 0.0  # Not used in DS model
                            sigma = 0
                            gamma = 0.5
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint (not used)
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.1:1.5")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "-0.2:0.2")  # Skew/asymmetry constraint
                        elif new_value in ["DS*G (A, \u03c3, \u03b3, S)"]:
                            # Doniach-Sunjic model
                            fraction = 20  # Not used in DS model
                            sigma = 0.5
                            gamma = 0.5
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint (not used)
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.3:1.5")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.1:1.5")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0:0.2")  # Skew/asymmetry constraint
                        elif new_value == "ExpGauss.(Area, \u03c3, \u03b3)":
                            # Exponential Gaussian
                            fraction = 20.0
                            sigma = 0.3
                            gamma = 1.2
                            skew = 0.64
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "0.01:1")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "0.01:3")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0.01:2")  # Skew constraint
                        elif new_value == "GL (Area)":
                            # GL Area based
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.15
                            skew = 0.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "5:80")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "Fixed")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "Fixed")  # Gamma constraint
                        elif new_value == "SGL (Area)":
                            # SGL Area based
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.15
                            skew = 0.1
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "5:80")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "Fixed")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "Fixed")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "0.01:1")  # Skew constraint
                        elif new_value == "D-parameter":
                            # D-parameter
                            fraction = 2.0
                            sigma = 1.0
                            gamma = 1.0
                            skew = 7.0
                            # Set constraints
                            self.window.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 7, "Fixed")  # Sigma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 8, "Fixed")  # Gamma constraint
                            self.window.peak_params_grid.SetCellValue(row + 1, 9, "5:9")  # Skew constraint
                        else:
                            # Default values for any other model
                            fraction = 20.0
                            sigma = 1.0
                            gamma = 0.15
                            skew = 0.1

                        # Get height and FWHM from grid
                        height = float(
                            self.window.peak_params_grid.GetCellValue(row, 3)) if self.window.peak_params_grid.GetCellValue(row,
                                                                                                              3) else 1000.0
                        fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4)) if self.window.peak_params_grid.GetCellValue(
                            row, 4) else 1.6

                        # Update grid values with model-specific defaults
                        self.window.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                        self.window.peak_params_grid.SetCellValue(row, 7, f"{sigma:.3f}")
                        self.window.peak_params_grid.SetCellValue(row, 8, f"{gamma:.3f}")
                        self.window.peak_params_grid.SetCellValue(row, 9, f"{skew:.3f}")

                        # Recalculate area with new model parameters
                        area = self.window.calculate_peak_area(new_value, height, fwhm, fraction, sigma, gamma, skew)
                        self.window.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")

                        # Update peak data in memory
                        peaks[correct_peak_key].update({
                            'L/G': fraction,
                            'Sigma': sigma,
                            'Gamma': gamma,
                            'Skew': skew,
                            'Area': area
                        })

                        # Update constraints in data structure
                        if 'Constraints' not in peaks[correct_peak_key]:
                            peaks[correct_peak_key]['Constraints'] = {}

                        # Get all constraint values from grid
                        for c_idx, c_key in enumerate(
                                ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew'], 2):
                            constraint_value = self.window.peak_params_grid.GetCellValue(row + 1, c_idx)
                            peaks[correct_peak_key]['Constraints'][c_key] = constraint_value

                        # Refresh the display
                        on_sheet_selected(self.window, sheet_name)
                elif row % 2 == 1:  # Constraint rowelse:  # Constraint row
                    if col in [2, 3, 4, 5, 6, 7, 8, 9]:
                        constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
                        column_to_constraint = {2: 0, 3: 1, 4: 2, 5: 3, 6: 4, 7: 5, 8: 6, 9:7}
                        constraint_key = constraint_keys[column_to_constraint[col]]

                        if 'Constraints' not in peaks[correct_peak_key]:
                            peaks[correct_peak_key]['Constraints'] = {}

                        # Always save the constraint value exactly as entered
                        peaks[correct_peak_key]['Constraints'][constraint_key] = new_value

            # Ensure numeric values are displayed with 2 decimal places
        if col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 0:  # Only for main parameter rows, not constraint rows
            try:
                formatted_value = f"{float(new_value):.2f}"
                self.window.peak_params_grid.SetCellValue(row, col, formatted_value)
            except ValueError:
                pass

        event.Skip()

        # Update all data in window.Data for all peaks
        for i in range(self.window.peak_params_grid.GetNumberRows() // 2):
            row_data = i * 2  # Data row
            row_constraint = i * 2 + 1  # Constraint row
            peak_label = self.window.peak_params_grid.GetCellValue(row_data, 1)

            # Update main parameters
            peaks[peak_label] = {
                'Position': float(self.window.peak_params_grid.GetCellValue(row_data, 2)),
                'Height': float(self.window.peak_params_grid.GetCellValue(row_data, 3)),
                'FWHM': float(self.window.peak_params_grid.GetCellValue(row_data, 4)),
                'L/G': float(self.window.peak_params_grid.GetCellValue(row_data, 5)),
                'Area': float(self.window.peak_params_grid.GetCellValue(row_data, 6)),
                'Sigma': float(self.window.peak_params_grid.GetCellValue(row_data, 7)) if self.window.peak_params_grid.GetCellValue(row_data, 7) else 0.0,
                'Gamma': float(self.window.peak_params_grid.GetCellValue(row_data, 8)) if self.window.peak_params_grid.GetCellValue(row_data, 8) else 0.0,
                'Skew': float(self.window.peak_params_grid.GetCellValue(row_data, 9)),
                'Fitting Model': self.window.peak_params_grid.GetCellValue(row_data, 13)
            }

            # Update constraints
            if 'Constraints' not in peaks[peak_label]:
                peaks[peak_label]['Constraints'] = {}

            constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
            for col_idx, key in enumerate(constraint_keys, start=2):
                value = self.window.peak_params_grid.GetCellValue(row_constraint, col_idx)

                # If value is empty, use defaults
                if not value:
                    print(f'Setting default for {key} in {peak_label}')
                    # Use appropriate default based on column
                    if key == 'Position':
                        # Get min/max from current sheet's data
                        x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
                        min_pos = min(x_values)
                        max_pos = max(x_values)
                        value = f"{min_pos:.2f}:{max_pos:.2f}"
                    elif key == 'Height':
                        value = '1:1e7'
                    elif key == 'FWHM':
                        value = '0.3:3.5'
                    elif key == 'L/G':
                        value = '5:80'
                    elif key == 'Area':
                        value = '1:1e7'
                    elif key == 'Sigma':
                        value = '0.3:3'
                    elif key == 'Gamma':
                        value = '0.3:3'
                    elif key == 'Skew':
                        value = '0.01:2'

                    # Update grid with default
                    self.window.peak_params_grid.SetCellValue(row_constraint, col_idx, value)

                peaks[peak_label]['Constraints'][key] = value

        # Refresh the grid to ensure it reflects the current state of self.window.Data
        self.refresh_peak_params_grid()

        # Replot the peaks with updated parameters
        self.window.clear_and_replot()

        save_state(self.window)

    def refresh_peak_params_grid(self):
        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.window.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            for i, (peak_label, peak_data) in enumerate(peaks.items()):
                row = i * 2
                self.window.peak_params_grid.SetCellValue(row, 1, peak_label)
                self.window.peak_params_grid.SetCellValue(row, 2, f"{peak_data['Position']:.2f}")
                self.window.peak_params_grid.SetCellValue(row, 3, f"{peak_data['Height']:.2f}")
                self.window.peak_params_grid.SetCellValue(row, 4, f"{peak_data['FWHM']:.2f}")
                self.window.peak_params_grid.SetCellValue(row, 5, f"{peak_data['L/G']:.2f}")
                try:
                    area_value = float(peak_data['Area'])
                    self.window.peak_params_grid.SetCellValue(row, 6, f"{area_value:.2f}")
                except (ValueError, KeyError):
                    self.window.peak_params_grid.SetCellValue(row, 6, "ER! REFRESH PEAK")
                # self.window.peak_params_grid.SetCellValue(row, 6, f"{float(peak_data['Area']):.2f}")
                # self.window.peak_params_grid.SetCellValue(row, 6, f"{peak_data['Area']:.2f}")
                self.window.peak_params_grid.SetCellValue(row, 7, f"{peak_data['Sigma']:.3f}")
                self.window.peak_params_grid.SetCellValue(row, 8, f"{peak_data['Gamma']:.3f}")
                self.window.peak_params_grid.SetCellValue(row, 9, f"{peak_data.get('Skew', 0.1):.3f}")
                if 'Constraints' in peak_data:
                    # Get min/max from current sheet's data
                    sheet_name = self.window.sheet_combobox.GetValue()
                    x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
                    min_pos = min(x_values)
                    max_pos = max(x_values)
                    position_constraint = f"{min_pos:.2f}:{max_pos:.2f}"

                    default_constraints = {
                        'Position': position_constraint,
                        'Height': '1:1e7',
                        'FWHM': '0.3:3.5',
                        'L/G': '5:80',
                        'Area': '1:1e7',
                        'Sigma': '0.3:3',
                        'Gamma': '0.3:3',
                        'Skew': '0.01:2'
                    }

                    for col_idx, key in enumerate(
                            ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew'], 2):
                        constraint_value = peak_data['Constraints'].get(key, '')
                        if not constraint_value:
                            constraint_value = default_constraints[key]
                            peak_data['Constraints'][key] = constraint_value
                        self.window.peak_params_grid.SetCellValue(row + 1, col_idx, str(constraint_value))
                else:
                    # Create default constraints
                    sheet_name = self.window.sheet_combobox.GetValue()
                    x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
                    min_pos = min(x_values)
                    max_pos = max(x_values)
                    position_constraint = f"{min_pos:.2f}:{max_pos:.2f}"

                    default_constraints = {
                        'Position': position_constraint,
                        'Height': '1:1e7',
                        'FWHM': '0.3:3.5',
                        'L/G': '5:80',
                        'Area': '1:1e7',
                        'Sigma': '0.3:3',
                        'Gamma': '0.3:3',
                        'Skew': '0.01:2'
                    }
                    peak_data['Constraints'] = default_constraints.copy()
                    for col_idx, (key, value) in enumerate(default_constraints.items(), 2):
                        self.window.peak_params_grid.SetCellValue(row + 1, col_idx, str(value))
        self.window.peak_params_grid.ForceRefresh()

    def refresh_peak_params_grid_release(self):
        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.window.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            for i, (peak_label, peak_data) in enumerate(peaks.items()):
                row = i * 2
                self.window.peak_params_grid.SetCellValue(row, 2, f"{peak_data['Position']:.2f}")
                self.window.peak_params_grid.SetCellValue(row, 3, f"{peak_data['Height']:.2f}")
                try:
                    area_value = float(peak_data['Area'])
                    self.window.peak_params_grid.SetCellValue(row, 6, f"{area_value:.2f}")
                except (ValueError, KeyError):
                    self.window.peak_params_grid.SetCellValue(row, 6, "ER! REFRESH PEAK")
            self.window.peak_params_grid.ForceRefresh()




    



