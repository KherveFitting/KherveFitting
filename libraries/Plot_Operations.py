# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

# plot_operations.py

import re
import wx
import os
import lmfit
import numpy as np
import numpy.ma as ma
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator, ScalarFormatter
from itertools import cycle
from PIL import Image
from scipy.ndimage import gaussian_filter

from libraries.Peak_Functions import PeakFunctions, BackgroundCalculations, OtherCalc

from libraries.PeakManipulation import PeakManipulation


class PlotManager:
    def __init__(self, ax, canvas, window=None):

        self.peak_manipulation = PeakManipulation(window)

        self.ax = ax
        self.canvas = canvas
        self.figure = ax.figure
        self.window = window
        self.parent = window  # Add parent attribute for backward compatibility

        if window:
            self.peak_manipulation = PeakManipulation(window)
        else:
            self.peak_manipulation = None

        self.cross = None
        self.peak_letter = None
        self.peak_info = None
        self.peak_letter_t = None
        self.peak_info_t = None
        self.peak_fill_enabled = True
        self.fitting_results_text = None
        self.fitting_results_visible = False

        self.residuals_state = 2  # Default to subplot
        self.residuals_subplot = None  # Initialize to None
        self.legend_visible = 1  # Default to full legend
        self.y_axis_state = 0  # Default to full y-axis

        # init for preference window
        self.plot_style = "scatter"
        self.scatter_size = 4
        self.line_width = 1
        self.line_alpha = 0.7
        self.scatter_color = "#000000"
        self.line_color = "#000000"
        self.scatter_marker = "o"

        self.background_color = "#808080"
        self.background_alpha = 0.5
        self.background_linestyle = "--"
        self.background_thickness = 1
        self.envelope_color = "#0000FF"
        self.envelope_alpha = 0.6
        self.envelope_linestyle = "-"
        self.envelope_thickness = 1
        self.residual_color = "#00FF00"
        self.residual_alpha = 0.4
        self.residual_linestyle = "-"
        self.raw_data_linestyle = "-"
        self.residual_thickness = 1

        self.peak_colors = []
        self.peak_alpha = 0.3

        self.energy_scale = 'BE'

        self.residuals_visible = True
        self.legend_visible = True
        self.rsd_text = None

        self.y_axis_visible = True

        self.survey_table_state = getattr(window, 'survey_table_state', 0)
        self.survey_table_text = []


    def toggle_y_axis_OLD(self):
        if not hasattr(self, 'y_axis_state'):
            self.y_axis_state = 0

        self.y_axis_state = (self.y_axis_state + 1) % 3

        if self.y_axis_state == 0:
            # Show label and values
            self.ax.yaxis.set_visible(True)
            self.ax.set_ylabel("Intensity (CPS)")
            self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
            self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(True)
                self.residuals_subplot.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
                self.residuals_subplot.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
        elif self.y_axis_state == 1:
            # Hide everything
            self.ax.yaxis.set_visible(False)
            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(False)
        else:
            # Show "Intensity (a.u.)" but hide values
            self.ax.yaxis.set_visible(True)
            self.ax.set_ylabel("Intensity (a.u.)")
            self.ax.yaxis.set_ticklabels([])
            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(True)
                self.residuals_subplot.yaxis.set_ticklabels([])

        self.canvas.draw_idle()

    def toggle_y_axis(self):
        if not hasattr(self, 'y_axis_state'):
            self.y_axis_state = 0

        self.y_axis_state = (self.y_axis_state + 1) % 3

        if self.y_axis_state == 0:
            # Show label and values
            self.ax.yaxis.set_visible(True)
            for label in self.ax.get_yticklabels():
                label.set_visible(True)
            self.ax.set_ylabel("Intensity (CPS)")
            self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
            self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(True)
                for label in self.residuals_subplot.get_yticklabels():
                    label.set_visible(True)
                self.residuals_subplot.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
                self.residuals_subplot.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        elif self.y_axis_state == 1:
            # Hide everything
            self.ax.yaxis.set_visible(False)

            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(False)

        else:  # y_axis_state == 2
            # Show "Intensity (a.u.)" but hide values
            self.ax.yaxis.set_visible(True)
            self.ax.set_ylabel("Intensity (a.u.)")

            # Hide just the tick labels
            for label in self.ax.get_yticklabels():
                label.set_visible(False)

            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.yaxis.set_visible(True)
                for label in self.residuals_subplot.get_yticklabels():
                    label.set_visible(False)

        self.canvas.draw_idle()

    def toggle_energy_scale(self, window):
        window.energy_scale = 'KE' if window.energy_scale == 'BE' else 'BE'
        self.clear_and_replot(window)

    def load_and_process_image(self, blur_sigma=4):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(current_dir, "Images", "SplashScreen4.png")

        if not os.path.exists(image_path):
            print(f"Image not found: {image_path}")
            return None

        # Load the PNG image
        img = Image.open(image_path) #.convert('L')  # Convert to grayscale

        # Convert to numpy array
        img_array = np.array(img)

        # Apply Gaussian blur
        blurred_img =  gaussian_filter(img_array, sigma=blur_sigma)

        # return blurred_img
        return img_array

    def plot_initial_logo(self):
        img_array = self.load_and_process_image()
        if img_array is None:
            return

        # Clear the current axis
        self.ax.clear()

        # Display the image
        self.ax.imshow(img_array, aspect='auto', alpha = 0.07, extent=[1350, 0, 0, 1000000])
        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')

        # Set axis limits
        self.ax.set_xlim(1350, 0)
        self.ax.set_ylim(0, 1000000)

        # Enable scientific notation for the Y-axis
        self.ax.ticklabel_format(axis='y', style='sci', scilimits=(0, 0))

        # Draw the canvas
        self.canvas.draw()

    def restore_green_vline_if_active(self, window):
        """Restore green vertical line if it was active before plot clearing"""
        if not (hasattr(window, 'green_vline_active') and window.green_vline_active):
            return

        # Remove any existing green line objects first
        if hasattr(window, 'green_vline') and window.green_vline is not None:
            try:
                window.green_vline.remove()
            except:
                pass

        if hasattr(window, 'green_vline_text') and window.green_vline_text is not None:
            try:
                window.green_vline_text.remove()
            except:
                pass

        # Always place at center of current plot
        xlim = window.ax.get_xlim()
        green_x = (xlim[0] + xlim[1]) / 2

        # Create fresh line on current axes
        window.green_vline = window.ax.axvline(green_x, color='green', linestyle='-',
                                               linewidth=1, alpha=0.7)

        # Create text label
        from libraries.Widgets_Toolbars import add_green_vline_text_label
        add_green_vline_text_label(window)
    def apply_text_settings(self, window):
        # Apply font settings
        plt.rcParams['font.family'] = window.plot_font

        # Apply axis title size
        self.ax.set_xlabel(self.ax.get_xlabel(), fontsize=window.axis_title_size)
        self.ax.set_ylabel(self.ax.get_ylabel(), fontsize=window.axis_title_size)

        # Apply axis number size
        self.ax.tick_params(axis='both', labelsize=window.axis_number_size)

        # Apply sublines (minor ticks)
        if window.x_sublines > 0:
            self.ax.xaxis.set_minor_locator(AutoMinorLocator(window.x_sublines + 1))
        if window.y_sublines > 0:
            self.ax.yaxis.set_minor_locator(AutoMinorLocator(window.y_sublines + 1))

        # Apply legend font size
        if self.ax.get_legend():
            plt.setp(self.ax.get_legend().get_texts(), fontsize=window.legend_font_size)

        for text in self.ax.texts:
            if getattr(text, 'sheet_name_text', False):
                text.set_fontsize(window.core_level_text_size)
                text.set_fontfamily([window.plot_font])

    def plot_peak(self, x_values, background, peak_params, sheet_name, window, color=None, alpha=0.3):
        row = peak_params['row']
        fwhm = peak_params['fwhm']
        lg_ratio = peak_params['lg_ratio']
        x = peak_params['position']
        y = peak_params['height']
        peak_label = peak_params['label']
        area = peak_params.get('area', 0)  # Get area if available
        # print('Plot_peak Area = '+str(area))

        formatted_label = re.sub(r'(\d+/\d+)', r'$_{\1}$', peak_label)

        fitting_model = peak_params.get('fitting_model', "GL (Height)")

        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
        bkg_y = background[np.argmin(np.abs(x_values - x))]
        if fitting_model == "Unfitted":
            return
        elif fitting_model == "SurveyID":
            return
        elif fitting_model == "D-parameter":
            # Get parameters for D-parameter calculation
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            skew = float(window.peak_params_grid.GetCellValue(row, 9))
            lg_ratio = float(window.peak_params_grid.GetCellValue(row, 5))

            # Calculate derivative
            normalized_deriv = OtherCalc.smooth_and_differentiate(
                x_values,
                window.y_values,
                skew,  # smooth_width
                sigma,  # pre_smooth
                lg_ratio,  # diff_width
                gamma  # post_smooth
            )

            # Plot derivative
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - x_values, normalized_deriv, '-', color=color, label=peak_label)
            else:
                self.ax.plot(x_values, normalized_deriv, '-', color=color, label=peak_label)

            return background  # Return background unchanged
        elif fitting_model == "Fermi":
            # Access Fermi data from window.Data structure
            sheet_name = window.sheet_combobox.GetValue()
            if sheet_name in window.Data['Core levels']:
                fitting_data = window.Data['Core levels'][sheet_name].get('Fitting', {})
                peaks_data = fitting_data.get('Peaks', {})

                # Find Fermi peak by fitting model instead of fixed name
                fermi_data = {}
                for peak_name, peak_info in peaks_data.items():
                    if peak_info.get('Fitting Model') == 'Fermi':
                        fermi_data = peak_info
                        break

                if 'Fitted_X' in fermi_data and 'Fitted_Y' in fermi_data:
                    fitted_x = np.array(fermi_data['Fitted_X'])
                    fitted_y = np.array(fermi_data['Fitted_Y'])

                    if window.energy_scale == 'KE':
                        plot_x = window.photons - fitted_x
                    else:
                        plot_x = fitted_x

                    # Plot Fermi fit
                    if window.energy_scale == 'KE':
                        self.ax.plot(plot_x, fitted_y, '-', color=color, linewidth=1, label=peak_label)
                    else:
                        self.ax.plot(plot_x, fitted_y, '-', color=color, linewidth=1, label=peak_label)

                    # Add center line with position in legend
                    center_pos = fermi_data.get('Fermi_Center', fermi_data.get('Position', 0))
                    if window.energy_scale == 'KE':
                        center_display = window.photons - center_pos
                    else:
                        center_display = center_pos

                    # self.ax.axvline(center_display, color='green', linestyle='--',
                    #                 linewidth=1, alpha=0.7,
                    #                 label=f'Center: {center_pos:.3f} eV')

                    # Add 16%-84% edge lines with width in meV in legend
                    width_16_84 = fermi_data.get('Fermi_16_84_Width', fermi_data.get('FWHM', 0))
                    edge_left = center_pos - width_16_84 / 2
                    edge_right = center_pos + width_16_84 / 2

                    if window.energy_scale == 'KE':
                        edge_left_display = window.photons - edge_left
                        edge_right_display = window.photons - edge_right
                    else:
                        edge_left_display = edge_left
                        edge_right_display = edge_right

                    # Show width in meV in legend
                    self.ax.axvline(edge_left_display, color='orange', linestyle=':',
                                    alpha=0.6,
                                    label=f'16%-84%: {width_16_84 * 1000:.1f} meV')
                    self.ax.axvline(edge_right_display, color='orange', linestyle=':', alpha=0.6)

                    self.ax.axvline(center_display, color='green', linestyle='--',
                                    linewidth=1, alpha=0.7,
                                    label=f'Center: {center_pos:.3f} eV')

            return background  # Return background unchanged
        elif fitting_model == "VBM":
            # Get VBM data from the Data structure (like Fermi does)
            sheet_name = window.sheet_combobox.GetValue()
            if ('Fitting' in window.Data['Core levels'][sheet_name] and
                    'Peaks' in window.Data['Core levels'][sheet_name]['Fitting']):

                peaks_data = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                vbm_data = None

                # Find VBM peak data
                for peak_name, peak_info in peaks_data.items():
                    if peak_info.get('Fitting Model') == 'VBM':
                        vbm_data = peak_info
                        break

                if vbm_data:
                    # Get VBM parameters from stored data
                    vbm_position = vbm_data.get('Position', 0)
                    edge_center = vbm_data.get('VBM_Edge_Center', 0)
                    bg_center = vbm_data.get('VBM_BG_Center', 0)
                    use_bg = bool(vbm_data.get('VBM_Use_BG', False))

                    # Get extrapolation data
                    signal_coef = vbm_data.get('Signal_Coef')
                    bg_coef = vbm_data.get('BG_Coef')
                    x_signal_fit = vbm_data.get('X_Signal_Fit')
                    y_signal_fit = vbm_data.get('Y_Signal_Fit')
                    x_bg_fit = vbm_data.get('X_BG_Fit')
                    y_bg_fit = vbm_data.get('Y_BG_Fit')

                    if window.energy_scale == 'KE':
                        vbm_display = window.photons - vbm_position
                        edge_display = window.photons - edge_center
                        bg_display = window.photons - bg_center
                    else:
                        vbm_display = vbm_position
                        edge_display = edge_center
                        bg_display = bg_center

                    # VBM position line (purple dashed, linewidth 1) - WITH LEGEND
                    window.ax.axvline(vbm_display, color='Green', linestyle='--',
                                      linewidth=1, label=f'VBM = {vbm_position:.2f} eV')


                    # Draw signal extrapolation line (red dashed, linewidth 1, alpha 0.6) - WITH LEGEND
                    if signal_coef and len(signal_coef) == 2:
                        x_extrap = np.linspace(min(x_values), max(x_values), 100)
                        if window.energy_scale == 'KE':
                            y_extrap = signal_coef[0] * (window.photons - x_extrap) + signal_coef[1]
                            window.ax.plot(x_extrap, y_extrap, 'r--', linewidth=1, alpha=0.6,
                                           label='Linear Extrapolation')
                        else:
                            y_extrap = signal_coef[0] * x_extrap + signal_coef[1]
                            window.ax.plot(x_extrap, y_extrap, 'r--', linewidth=1, alpha=0.6,
                                           label='Linear Extrapolation')

                    # Draw signal fit points (blue circles, markersize 2) - NO LEGEND
                    if x_signal_fit and y_signal_fit:
                        x_sig = np.array(x_signal_fit)
                        y_sig = np.array(y_signal_fit)
                        if window.energy_scale == 'KE':
                            window.ax.plot(window.photons - x_sig, y_sig, 'bo', markersize=2)
                        else:
                            window.ax.plot(x_sig, y_sig, 'bo', markersize=2)

                    # Draw background extrapolation and points if used
                    if use_bg and bg_coef and len(bg_coef) == 2:
                        # Background extrapolation line (blue dashed, linewidth 1) - WITH LEGEND
                        x_bg_extrap = np.linspace(min(x_values), max(x_values), 100)
                        if window.energy_scale == 'KE':
                            y_bg_extrap = bg_coef[0] * (window.photons - x_bg_extrap) + bg_coef[1]
                            window.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                           label='Background Fit')
                        else:
                            y_bg_extrap = bg_coef[0] * x_bg_extrap + bg_coef[1]
                            window.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                           label='Background Fit')

                        # Draw background fit points (blue circles, markersize 2) - NO LEGEND
                        if x_bg_fit and y_bg_fit:
                            x_bg = np.array(x_bg_fit)
                            y_bg = np.array(y_bg_fit)
                            if window.energy_scale == 'KE':
                                window.ax.plot(window.photons - x_bg, y_bg, 'bo', markersize=2)
                            else:
                                window.ax.plot(x_bg, y_bg, 'bo', markersize=2)

            # CRITICAL: Return early to avoid peak_model.eval() call
            return
        elif fitting_model == "Cut-Off":
            # Get Cut-Off data from the Data structure (like Fermi does)
            sheet_name = window.sheet_combobox.GetValue()
            if ('Fitting' in window.Data['Core levels'][sheet_name] and
                    'Peaks' in window.Data['Core levels'][sheet_name]['Fitting']):

                peaks_data = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                cutoff_data = None

                # Find Cut-Off peak data
                for peak_name, peak_info in peaks_data.items():
                    if peak_info.get('Fitting Model') == 'Cut-Off':
                        cutoff_data = peak_info
                        break

                if cutoff_data:
                    # Get Cut-Off parameters from stored data
                    cutoff_position = cutoff_data.get('Position', 0)
                    edge_center = cutoff_data.get('Cut-Off_Edge_Center', 0)
                    bg_center = cutoff_data.get('Cut-Off_BG_Center', 0)
                    use_bg = bool(cutoff_data.get('Cut-Off_Use_BG', False))

                    # Get extrapolation data
                    signal_coef = cutoff_data.get('Signal_Coef')
                    bg_coef = cutoff_data.get('BG_Coef')
                    x_signal_fit = cutoff_data.get('X_Signal_Fit')
                    y_signal_fit = cutoff_data.get('Y_Signal_Fit')
                    x_bg_fit = cutoff_data.get('X_BG_Fit')
                    y_bg_fit = cutoff_data.get('Y_BG_Fit')

                    if window.energy_scale == 'KE':
                        cutoff_display = window.photons - cutoff_position
                        edge_display = window.photons - edge_center
                        bg_display = window.photons - bg_center
                    else:
                        cutoff_display = cutoff_position
                        edge_display = edge_center
                        bg_display = bg_center

                    # Cut-Off position line (orange dashed, linewidth 1) - WITH LEGEND
                    window.ax.axvline(cutoff_display, color='Orange', linestyle='--',
                                      linewidth=1, label=f'Cut-Off = {cutoff_position:.2f} eV')

                    # Draw signal extrapolation line (red dashed, linewidth 1, alpha 0.6) - WITH LEGEND
                    if signal_coef and len(signal_coef) == 2:
                        x_extrap = np.linspace(min(x_values), max(x_values), 100)
                        if window.energy_scale == 'KE':
                            y_extrap = signal_coef[0] * (window.photons - x_extrap) + signal_coef[1]
                            window.ax.plot(x_extrap, y_extrap, 'r--', linewidth=1, alpha=0.6,
                                           label='Linear Extrapolation')
                        else:
                            y_extrap = signal_coef[0] * x_extrap + signal_coef[1]
                            window.ax.plot(x_extrap, y_extrap, 'r--', linewidth=1, alpha=0.6,
                                           label='Linear Extrapolation')

                    # Draw signal fit points (blue circles, markersize 2) - NO LEGEND
                    if x_signal_fit and y_signal_fit:
                        x_sig = np.array(x_signal_fit)
                        y_sig = np.array(y_signal_fit)
                        if window.energy_scale == 'KE':
                            window.ax.plot(window.photons - x_sig, y_sig, 'bo', markersize=2)
                        else:
                            window.ax.plot(x_sig, y_sig, 'bo', markersize=2)

                    # Draw background extrapolation and points if used
                    if use_bg and bg_coef and len(bg_coef) == 2:
                        # Background extrapolation line (blue dashed, linewidth 1) - WITH LEGEND
                        x_bg_extrap = np.linspace(min(x_values), max(x_values), 100)
                        if window.energy_scale == 'KE':
                            y_bg_extrap = bg_coef[0] * (window.photons - x_bg_extrap) + bg_coef[1]
                            window.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                           label='Background Fit')
                        else:
                            y_bg_extrap = bg_coef[0] * x_bg_extrap + bg_coef[1]
                            window.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                           label='Background Fit')

                        # Draw background fit points (blue circles, markersize 2) - NO LEGEND
                        if x_bg_fit and y_bg_fit:
                            x_bg = np.array(x_bg_fit)
                            y_bg = np.array(y_bg_fit)
                            if window.energy_scale == 'KE':
                                window.ax.plot(window.photons - x_bg, y_bg, 'bo', markersize=2)
                            else:
                                window.ax.plot(x_bg, y_bg, 'bo', markersize=2)

            # CRITICAL: Return early to avoid peak_model.eval() call
            return

        elif fitting_model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
            peak_model = lmfit.models.VoigtModel()
            sigma = float(peak_params.get('sigma', 1.2)) / 2.355
            gamma = float(peak_params.get('gamma', 0.06)) / 2
            amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0)
            params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma)
        elif fitting_model in ["Voigt (Area, L/G, \u03c3, S)"]:
            peak_model = lmfit.models.SkewedVoigtModel()
            amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7)) / 2.355
            gamma = float(window.peak_params_grid.GetCellValue(row, 8)) / 2
            skew =  float(window.peak_params_grid.GetCellValue(row, 9))
            params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma, skew=skew)
        elif fitting_model == "DS (A, \u03c3, \u03b3)":
            peak_model = lmfit.models.DoniachModel()
            height = float(window.peak_params_grid.GetCellValue(row, 3))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            skew = float(window.peak_params_grid.GetCellValue(row, 9))
            amplitude = PeakFunctions.doniach_sunjic_height_to_amplitude(height, sigma, gamma, skew)
            params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma, asymmetry=skew)
        elif fitting_model == "DS*G (A, \u03c3, \u03b3, S)":
            peak_model = lmfit.Model(PeakFunctions.DS_G)
            amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            skew = float(window.peak_params_grid.GetCellValue(row, 9))
            params = peak_model.make_params(center=x, amplitude=amplitude, gamma=gamma,
                                            skew=skew, sigma=sigma)
        elif fitting_model == "ExpGauss.(Area, \u03c3, \u03b3)":
            peak_model = lmfit.models.ExponentialGaussianModel()
            area = float(window.peak_params_grid.GetCellValue(row, 6))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            params = peak_model.make_params(center=x, amplitude=area, sigma=sigma, gamma=gamma)
            peak_y = peak_model.eval(params, x=x_values)
            # Calculate height numerically
            y_values = peak_model.eval(params, x=x_values)
            height = np.max(y_values)
            # Estimate FWHM numerically
            half_max = height / 2
            indices = np.where(y_values >= half_max)[0]
            if len(indices) >= 2:
                fwhm = abs(x_values[indices[-1]] - x_values[indices[0]])
            else:
                fwhm = None  # or some default value
            fraction = gamma / (sigma + gamma) * 100
        elif fitting_model == "Pseudo-Voigt (Area)":
            sigma = fwhm / 2
            peak_model = lmfit.models.PseudoVoigtModel()
            amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, fraction=lg_ratio / 100, x=0)
            params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, fraction=lg_ratio / 100)
        elif fitting_model in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_model = lmfit.Model(PeakFunctions.LA)
            amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            params = peak_model.make_params(center=x,amplitude=amplitude,fwhm=fwhm,sigma=sigma,gamma=gamma)

            # No direct equivalent to 'fraction' for LA model
            fraction = (sigma + gamma) / 2  # You could define it differently if needed
        elif fitting_model in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_model = lmfit.Model(PeakFunctions.LAxG)
            amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
            sigma = float(window.peak_params_grid.GetCellValue(row, 7))
            gamma = float(window.peak_params_grid.GetCellValue(row, 8))
            fwhm_g = float(window.peak_params_grid.GetCellValue(row, 9))
            params = peak_model.make_params(center=x,amplitude=amplitude,fwhm=fwhm,sigma=sigma,gamma=gamma,
                                            fwhm_g=fwhm_g)

            # No direct equivalent to 'fraction' for LA model
            fraction = (sigma + gamma) / 2  # You could define it differently if needed
        elif fitting_model == "GL (Height)":
            peak_model = lmfit.Model(PeakFunctions.gauss_lorentz)
            params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)
        elif fitting_model == "SGL (Height)":
            peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz)
            params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)
        elif fitting_model == "GL (Area)":
            peak_model = lmfit.Model(PeakFunctions.gauss_lorentz_Area)
            area = y * (fwhm * np.sqrt(np.pi / (4 * np.log(2))))
            params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, area=area)
        elif fitting_model == "SGL (Area)":
            peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz_Area)
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = fwhm / 2
            area = y * ((1 - lg_ratio / 100) * sigma * np.sqrt(2 * np.pi) + (lg_ratio / 100) * np.pi * gamma)
            params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, area=area)
        elif fitting_model == "SingleEntity_OLD":
            # Handle SingleEntity envelope - get stored data and interpolate
            if ('Fitting' in window.Data['Core levels'][sheet_name] and
                    'Peaks' in window.Data['Core levels'][sheet_name]['Fitting']):

                peaks_dict = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']

                # Find the peak data with SingleEntity model
                peak_data = None
                for peak_name, data in peaks_dict.items():
                    if data.get('Fitting Model') == 'SingleEntity':
                        # Check if this is the right peak by position match (original position)
                        original_pos = data.get('Position', 0)
                        if abs(original_pos - x) < 0.01:
                            peak_data = data
                            break

                if peak_data and 'x_data' in peak_data and 'y_data' in peak_data:
                    # Get stored envelope data
                    x_env = np.array(peak_data['x_data'])
                    y_env = np.array(peak_data['y_data'])

                    # Get shift and scale from sigma/gamma columns
                    position_shift = float(window.peak_params_grid.GetCellValue(row, 7))  # Sigma = shift
                    area_scale = float(window.peak_params_grid.GetCellValue(row, 8))  # Gamma = scale

                    # Create interpolator from original envelope
                    from scipy.interpolate import interp1d
                    interpolator = interp1d(x_env, y_env, kind='cubic',
                                            bounds_error=False, fill_value=0.0)

                    # Apply shift and scale to envelope
                    x_shifted = x_values - position_shift  # Shift envelope position
                    y_interpolated = interpolator(x_shifted)
                    peak_y = y_interpolated * area_scale  # Scale envelope height

                    # Create interpolator
                    from scipy.interpolate import interp1d
                    interpolator = interp1d(x_env, y_env, kind='cubic',
                                            bounds_error=False, fill_value=0.0)

                    # Shift x and interpolate
                    x_shifted = x_values - position_shift
                    y_interpolated = interpolator(x_shifted)

                    # Scale by area and add background
                    peak_y = y_interpolated * area_scale + background

                    # Plot the SingleEntity
                    if color is None:
                        color = self.peak_colors[len(self.ax.lines) % len(self.peak_colors)]
                    if alpha is None:
                        alpha = self.peak_alpha

                    if window.energy_scale == 'KE':
                        self.ax.fill_between(window.photons - x_values, background, peak_y,
                                             color=color, alpha=alpha, label=formatted_label)
                    else:
                        self.ax.fill_between(x_values, background, peak_y,
                                             color=color, alpha=alpha, label=formatted_label)

                    return background
                else:
                    # If no data found, skip this peak
                    return background
            else:
                return background
        elif fitting_model == "SingleEntity":
            # Handle SingleEntity envelope - get stored data and interpolate
            if ('Fitting' in window.Data['Core levels'][sheet_name] and
                    'Peaks' in window.Data['Core levels'][sheet_name]['Fitting']):

                peaks_dict = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']

                # Find the peak data with SingleEntity model
                peak_data = None
                for peak_name, data in peaks_dict.items():
                    if data.get('Fitting Model') == 'SingleEntity':
                        # Check if this is the right peak by position match (original position)
                        original_pos = data.get('Position', 0)
                        if abs(original_pos - x) < 0.01:
                            peak_data = data
                            break

                if peak_data and 'x_data' in peak_data and 'y_data' in peak_data:
                    # Get stored envelope data
                    x_env = np.array(peak_data['x_data'])
                    y_env = np.array(peak_data['y_data'])

                    # Get shift and scale from sigma/gamma columns
                    position_shift = float(window.peak_params_grid.GetCellValue(row, 7))  # Sigma = shift
                    area_scale = float(window.peak_params_grid.GetCellValue(row, 8))  # Gamma = scale

                    # Create interpolator from original envelope
                    from scipy.interpolate import interp1d
                    interpolator = interp1d(x_env, y_env, kind='cubic',
                                            bounds_error=False, fill_value=0.0)

                    # Apply shift and scale to envelope
                    x_shifted = x_values - position_shift  # Shift envelope position
                    y_interpolated = interpolator(x_shifted)
                    peak_y = y_interpolated * area_scale  # Scale envelope height

                    # Create interpolator
                    from scipy.interpolate import interp1d
                    interpolator = interp1d(x_env, y_env, kind='cubic',
                                            bounds_error=False, fill_value=0.0)

                    # Shift x and interpolate
                    x_shifted = x_values - position_shift
                    y_interpolated = interpolator(x_shifted)

                    # Scale by area and add background
                    peak_y = y_interpolated * area_scale + background

                    # Plot the SingleEntity - same as other peaks
                    if color is None:
                        color = self.peak_colors[len(self.ax.lines) % len(self.peak_colors)]
                    if alpha is None:
                        alpha = self.peak_alpha

                    # Get background regions for masking
                    bg_regions = self.get_background_regions(window)

                    # Apply background region masking using NaN
                    if bg_regions is not None and len(bg_regions) > 0:
                        # Create mask for all background regions
                        region_mask = np.zeros(len(x_values), dtype=bool)
                        for bg_start, bg_end in bg_regions:
                            region_mask |= (x_values >= bg_start) & (x_values <= bg_end)

                        # Work with copies to avoid affecting original data
                        peak_y_masked = peak_y.copy()
                        peak_y_masked = peak_y_masked.astype(float)
                        peak_y_masked[~region_mask] = np.nan

                        background_masked = background.copy()
                        background_masked = background_masked.astype(float)
                        background_masked[~region_mask] = np.nan
                    else:
                        # No masking - but still use copies
                        peak_y_masked = peak_y.copy()
                        background_masked = background.copy()

                    # Get peak index and fill type for consistency with other peaks
                    num_peaks = window.peak_params_grid.GetNumberRows() // 2
                    peak_index = 0
                    for i in range(num_peaks):
                        if window.peak_params_grid.GetCellValue(i * 2, 1) == peak_label:
                            peak_index = i
                            break

                    # Determine fill type
                    fill_type = "Solid Fill"
                    if hasattr(window, 'peak_fill_types') and len(window.peak_fill_types) > 0:
                        fill_type = window.peak_fill_types[peak_index % len(window.peak_fill_types)]

                    if fill_type == "Solid Fill":
                        # Plot fill_between with solid fill
                        if bg_regions is not None and len(bg_regions) > 0:
                            if window.energy_scale == 'KE':
                                self.ax.fill_between(window.photons - x_values, background_masked, peak_y_masked,
                                                     color=color, alpha=alpha, edgecolor='none', label=formatted_label)
                            else:
                                self.ax.fill_between(x_values, background_masked, peak_y_masked,
                                                     color=color, alpha=alpha, edgecolor='none', label=formatted_label)
                        else:
                            if window.energy_scale == 'KE':
                                self.ax.fill_between(window.photons - x_values, background, peak_y,
                                                     color=color, alpha=alpha, edgecolor='none', label=formatted_label)
                            else:
                                self.ax.fill_between(x_values, background, peak_y,
                                                     color=color, alpha=alpha, edgecolor='none', label=formatted_label)

                    elif fill_type == "Hatch":
                        # Use fill_between with hatch pattern
                        hatch_pattern = "/"
                        if hasattr(window, 'peak_hatch_patterns') and len(window.peak_hatch_patterns) > 0:
                            hatch_pattern = window.peak_hatch_patterns[peak_index % len(window.peak_hatch_patterns)]
                        hatch_density = getattr(window, 'hatch_density', 2)
                        final_hatch = hatch_pattern * hatch_density

                        if bg_regions is not None and len(bg_regions) > 0:
                            if window.energy_scale == 'KE':
                                self.ax.fill_between(window.photons - x_values, background_masked, peak_y_masked,
                                                     color='none', hatch=final_hatch,
                                                     linewidth=getattr(window, 'peak_line_thickness', 1),
                                                     edgecolor=color, alpha=alpha, label=formatted_label)
                            else:
                                self.ax.fill_between(x_values, background_masked, peak_y_masked,
                                                     color='none', hatch=final_hatch,
                                                     linewidth=getattr(window, 'peak_line_thickness', 1),
                                                     edgecolor=color, alpha=alpha, label=formatted_label)
                        else:
                            if window.energy_scale == 'KE':
                                self.ax.fill_between(window.photons - x_values, background, peak_y,
                                                     color='none', hatch=final_hatch,
                                                     linewidth=getattr(window, 'peak_line_thickness', 1),
                                                     edgecolor=color, alpha=alpha, label=formatted_label)
                            else:
                                self.ax.fill_between(x_values, background, peak_y,
                                                     color='none', hatch=final_hatch,
                                                     linewidth=getattr(window, 'peak_line_thickness', 1),
                                                     edgecolor=color, alpha=alpha, label=formatted_label)

                    elif fill_type == "None":
                        # Only draw the line (no fill) - similar to other peaks
                        peak_line_style = getattr(window, 'peak_line_style', 'Same Color')
                        if peak_line_style != "No Line":
                            # Determine line color
                            if peak_line_style == "Black":
                                line_color = 'black'
                            elif peak_line_style == "Grey":
                                line_color = 'grey'
                            elif peak_line_style == "Yellow":
                                line_color = 'yellow'
                            else:  # Same Color
                                line_color = color

                            if bg_regions is not None and len(bg_regions) > 0:
                                if window.energy_scale == 'KE':
                                    self.ax.plot(window.photons - x_values, peak_y_masked,
                                                 color=line_color,
                                                 alpha=getattr(window, 'peak_line_alpha', 0.7),
                                                 linewidth=getattr(window, 'peak_line_thickness', 1),
                                                 linestyle=getattr(window, 'peak_line_pattern', '-'),
                                                 label=formatted_label)
                                else:
                                    self.ax.plot(x_values, peak_y_masked,
                                                 color=line_color,
                                                 alpha=getattr(window, 'peak_line_alpha', 0.7),
                                                 linewidth=getattr(window, 'peak_line_thickness', 1),
                                                 linestyle=getattr(window, 'peak_line_pattern', '-'),
                                                 label=formatted_label)
                            else:
                                if window.energy_scale == 'KE':
                                    self.ax.plot(window.photons - x_values, peak_y,
                                                 color=line_color,
                                                 alpha=getattr(window, 'peak_line_alpha', 0.7),
                                                 linewidth=getattr(window, 'peak_line_thickness', 1),
                                                 linestyle=getattr(window, 'peak_line_pattern', '-'),
                                                 label=formatted_label)
                                else:
                                    self.ax.plot(x_values, peak_y,
                                                 color=line_color,
                                                 alpha=getattr(window, 'peak_line_alpha', 0.7),
                                                 linewidth=getattr(window, 'peak_line_thickness', 1),
                                                 linestyle=getattr(window, 'peak_line_pattern', '-'),
                                                 label=formatted_label)

                    # Add peak line on top if peak_line_style is not "No Line" (same as other peaks)
                    peak_line_style = getattr(window, 'peak_line_style', 'Same Color')
                    if peak_line_style != "No Line" and fill_type != "None":
                        # Determine line color
                        if peak_line_style == "Black":
                            line_color = "black"
                        elif peak_line_style == "Grey":
                            line_color = "grey"
                        elif peak_line_style == "Yellow":
                            line_color = "yellow"
                        else:  # Same Color
                            line_color = color

                        # Use the same masking as fill
                        if bg_regions is not None and len(bg_regions) > 0:
                            if window.energy_scale == 'KE':
                                self.ax.plot(window.photons - x_values, peak_y_masked,
                                             color=line_color,
                                             alpha=getattr(window, 'peak_line_alpha', 0.7),
                                             linewidth=getattr(window, 'peak_line_thickness', 1),
                                             linestyle=getattr(window, 'peak_line_pattern', '-'))
                            else:
                                self.ax.plot(x_values, peak_y_masked,
                                             color=line_color,
                                             alpha=getattr(window, 'peak_line_alpha', 0.7),
                                             linewidth=getattr(window, 'peak_line_thickness', 1),
                                             linestyle=getattr(window, 'peak_line_pattern', '-'))
                        else:
                            if window.energy_scale == 'KE':
                                self.ax.plot(window.photons - x_values, peak_y,
                                             color=line_color,
                                             alpha=getattr(window, 'peak_line_alpha', 0.7),
                                             linewidth=getattr(window, 'peak_line_thickness', 1),
                                             linestyle=getattr(window, 'peak_line_pattern', '-'))
                            else:
                                self.ax.plot(x_values, peak_y,
                                             color=line_color,
                                             alpha=getattr(window, 'peak_line_alpha', 0.7),
                                             linewidth=getattr(window, 'peak_line_thickness', 1),
                                             linestyle=getattr(window, 'peak_line_pattern', '-'))

                    return background
                else:
                    # If no data found, skip this peak
                    return background
            else:
                return background

        else:
            raise ValueError(f"Unknown fitting model: {fitting_model}")

        peak_y = peak_model.eval(params, x=x_values) + background

        # Rest of the function remains the same
        if color is None:
            color = self.peak_colors[len(self.ax.lines) % len(self.peak_colors)]
        if alpha is None:
            alpha = self.peak_alpha

        if window.peak_line_style == "Black":
            line_color = "black"
        elif window.peak_line_style == "Grey":
            line_color = "grey"
        elif window.peak_line_style == "Yellow":
            line_color = "yellow"
        else:  # same_color
            line_color = color

        line_alpha = min(alpha + 0.1, 1)

        if self.peak_fill_enabled:
            label = peak_label

            # Get background regions from recorded ranges
            bg_regions = self.get_background_regions(window)

            # Apply background region masking using NaN - DO THIS FIRST
            if bg_regions is not None and len(bg_regions) > 0:
                # Create mask for all background regions
                region_mask = np.zeros(len(x_values), dtype=bool)
                for bg_start, bg_end in bg_regions:
                    region_mask |= (x_values >= bg_start) & (x_values <= bg_end)

                # IMPORTANT: Always work with copies to avoid affecting original data
                peak_y_masked = peak_y.copy()  # This is already a copy
                peak_y_masked = peak_y_masked.astype(float)
                peak_y_masked[~region_mask] = np.nan

                background_masked = background.copy()  # This is already a copy
                background_masked = background_masked.astype(float)
                background_masked[~region_mask] = np.nan
            else:
                # No masking - but still use copies
                peak_y_masked = peak_y.copy()
                background_masked = background.copy()

            # Identify doublets
            num_peaks = window.peak_params_grid.GetNumberRows() // 2
            doublets = []
            for i in range(0, num_peaks - 1):
                current_label = window.peak_params_grid.GetCellValue(i * 2, 1)
                next_label = window.peak_params_grid.GetCellValue((i + 1) * 2, 1)
                if self.is_part_of_doublet(current_label, next_label):
                    doublets.extend([i, i + 1])

            # Find current peak index
            peak_index = 0  # Default value
            for i in range(num_peaks):
                if window.peak_params_grid.GetCellValue(i * 2, 1) == peak_label:
                    peak_index = i
                    break

            # If part of doublet, get fill type from first peak of the pair
            if peak_index in doublets:
                if doublets.index(peak_index) % 2 == 1:  # Second peak of doublet
                    peak_index = peak_index - 1  # Use first peak's settings

            fill_idx = peak_index % len(window.peak_fill_types)
            hatch_idx = peak_index % len(window.peak_hatch_patterns)
            if window.peak_fill_types[fill_idx] == "Solid Fill":
                fill_params = {
                    'color': color,
                    'alpha': alpha,
                    'edgecolor': 'none'
                }
            elif window.peak_fill_types[fill_idx] == "Hatch":
                fill_params = {
                    'color': 'none',
                    'hatch': window.peak_hatch_patterns[hatch_idx] * window.hatch_density,
                    'linewidth': window.peak_line_thickness,
                    'edgecolor': color,
                    'alpha': alpha
                }
            elif window.peak_fill_types[fill_idx] == "None":
                # Skip the fill_between call and only draw the line
                if window.peak_line_style != "No Line":
                    # Determine line color
                    if window.peak_line_style == "Black":
                        line_color = "black"
                    elif window.peak_line_style == "Grey":
                        line_color = "grey"
                    elif window.peak_line_style == "Yellow":
                        line_color = "yellow"
                    else:  # same_color
                        line_color = color

                    if window.energy_scale == 'KE':
                        self.ax.plot(window.photons - x_values, peak_y_masked, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern,
                                     label=peak_label)
                    else:
                        self.ax.plot(x_values, peak_y_masked, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern,
                                     label=peak_label)
                return peak_y  # CRITICAL: Return original peak_y, not peak_y_masked


            if bg_regions is not None and len(bg_regions) > 0:
                # Use masked data - only show within background regions
                if window.energy_scale == 'KE':
                    self.ax.fill_between(window.photons - x_values, background_masked, peak_y_masked, label=label, **fill_params)
                else:
                    self.ax.fill_between(x_values, background_masked, peak_y_masked, label=label, **fill_params)
            else:
                # Use full data (no masking) - show everywhere
                if window.energy_scale == 'KE':
                    self.ax.fill_between(window.photons - x_values, background, peak_y, label=label, **fill_params)
                else:
                    self.ax.fill_between(x_values, background, peak_y, label=label, **fill_params)
            # AND THIS-----------------------------------------

            if window.peak_line_style != "No Line":
                if window.peak_line_style == "Black":
                    line_color = "black"
                elif window.peak_line_style == "Grey":
                    line_color = "grey"
                elif window.peak_line_style == "Yellow":
                    line_color = "yellow"
                else:  # same_color
                    line_color = color


                # Use the same masking as above
                if bg_regions is not None and len(bg_regions) > 0:
                    # Use the same masked data
                    if window.energy_scale == 'KE':
                        self.ax.plot(window.photons - x_values, peak_y_masked, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern)
                    else:
                        self.ax.plot(x_values, peak_y_masked, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern)
                else:
                    # No masking - use original data
                    if window.energy_scale == 'KE':
                        self.ax.plot(window.photons - x_values, peak_y, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern)
                    else:
                        self.ax.plot(x_values, peak_y, color=line_color, alpha=window.peak_line_alpha,
                                     linewidth=window.peak_line_thickness, linestyle=window.peak_line_pattern)



        else:
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - x_values, peak_y, color=color, alpha=line_alpha, label=peak_label)
            else:
                self.ax.plot(x_values, peak_y, color=color, alpha=line_alpha, label=peak_label)

        # Restore green line if active
        self.restore_green_vline_if_active(window)

        self.canvas.draw_idle()

        return peak_y

    def plot_data(self, window):
        if 'Core levels' not in window.Data or not window.Data['Core levels']:
            # wx.MessageBox("No data available to plot.", "Error", wx.OK | wx.ICON_ERROR)
            self.parent.show_popup_message2("Error", "No data available to plot.")
            return

        sheet_name = window.sheet_combobox.GetValue()
        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper()
        is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')

        is_d_parameter = False
        limits = window.plot_config.get_plot_limits(window, sheet_name)

        # CHECK IF THIS IS A PROFILE SHEET
        if sheet_name.startswith('zzProfile'):
            print(f"plot_data: Detected profile sheet: {sheet_name}")
            if self.plot_profile(window):
                return  # Profile plotted successfully

        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            self.ax.clear()

            # Reset background to white for non-EDX plots
            if not (sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot') or sheet_name.startswith('EELS~Plot')):
                self.figure.patch.set_facecolor('white')
                self.ax.set_facecolor('white')
                # Reset spines to black
                for spine in self.ax.spines.values():
                    spine.set_edgecolor('black')
                # Reset tick colors to black
                self.ax.tick_params(colors='black', labelcolor='black')

            # Remove heatmap colorbar AND its axes if it exists
            if hasattr(window, 'heatmap_colorbar') and window.heatmap_colorbar is not None:
                try:
                    if hasattr(window.heatmap_colorbar, 'ax'):
                        self.figure.delaxes(window.heatmap_colorbar.ax)
                    window.heatmap_colorbar = None
                    window.heatmap_data = None
                    window.heatmap_sheets = None
                except:
                    pass

            if hasattr(self, 'residuals_subplot'):
                if self.residuals_subplot:
                    self.figure.delaxes(self.residuals_subplot)
                    self.residuals_subplot = None
                    # self.ax.set_position([0.1, 0.1, 0.8, 0.8])
                    # self.ax.set_position([0.1, 0.125, 0.85, 0.85])
                    self.ax.get_xaxis().set_visible(True)

            # Always set position for consistent layout (including survey/wide plots)
            # self.ax.set_position([0.1, 0.125, 0.85, 0.85])
            self.ax.set_position([0.1, 0.1, 0.85, 0.85])

            x_values = window.Data['Core levels'][sheet_name]['B.E.']
            # if hasattr(window, 'be_correction'):
            #     x_values = np.array(x_values) + window.be_correction
            y_values = window.Data['Core levels'][sheet_name]['Raw Data']

            x_values = np.array(x_values)  # Ensure x_values is a numpy array

            # Update window.x_values and window.y_values
            window.x_values = np.array(x_values)
            window.y_values = np.array(y_values)

            window.background = np.array(
                window.Data['Core levels'][sheet_name]['Background']['Bkg Y']) if 'Background' in \
                            window.Data['Core levels'][sheet_name] and 'Bkg Y' in window.Data['Core levels'][
                                sheet_name]['Background'] else window.y_values

            # Initialize Bkg X if not already present
            if 'Bkg X' not in window.Data['Core levels'][sheet_name]['Background'] or not \
                    window.Data['Core levels'][sheet_name]['Background']['Bkg X']:
                window.Data['Core levels'][sheet_name]['Background']['Bkg X'] = window.x_values.tolist()

            # Set x-axis limits to reverse the direction and match the min and max of the data
            if window.energy_scale == 'KE':
                X_MIN = window.photons-limits['Xmax']
                X_MAX = window.photons-limits['Xmin']
                self.ax.set_xlim(min(X_MIN, X_MAX),max(X_MIN, X_MAX))  # Reverse X-axis
            else:
                self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis

            self.ax.set_ylim(limits['Ymin'], limits['Ymax'])

            self.ax.set_ylabel("Intensity (CPS)")
            x_label = "Kinetic Energy (eV)" if window.energy_scale == 'KE' else "Binding Energy (eV)"
            self.ax.set_xlabel(x_label)
            # self.ax.set_xlabel("Binding Energy (eV)")

            self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
            self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

            if "survey" in sheet_name.lower() or "wide" in sheet_name.lower():
                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons- x_values, y_values, c=self.line_color, linewidth=self.line_width,
                                 alpha=self.line_alpha, linestyle=self.raw_data_linestyle)  # , label='Raw Data')
                else:
                    self.ax.plot(x_values, y_values, c=self.line_color, linewidth=self.line_width,
                             alpha=self.line_alpha, linestyle=self.raw_data_linestyle) #, label='Raw Data')
            elif self.plot_style == "scatter":
                if window.energy_scale == 'KE':
                    self.ax.scatter(window.photons-x_values, y_values, c=self.scatter_color, s=self.scatter_size,
                                marker=self.scatter_marker, label='Raw Data')
                else:
                    self.ax.scatter(x_values, y_values, c=self.scatter_color, s=self.scatter_size,
                                marker=self.scatter_marker, label='Raw Data')
            else:
                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons -x_values, y_values, c=self.line_color, linewidth=self.line_width,
                             alpha=self.line_alpha, linestyle=self.raw_data_linestyle, label='Raw Data')
                else:
                    self.ax.plot(x_values, y_values, c=self.line_color, linewidth=self.line_width,
                                 alpha=self.line_alpha, linestyle=self.raw_data_linestyle, label='Raw Data')

            if 'Labels' in window.Data['Core levels'][sheet_name]:
                table_exists = False

                for label_data in window.Data['Core levels'][sheet_name]['Labels']:
                    # Skip drawing Table entities - they are handled separately
                    if label_data.get('text') == 'Table' and label_data.get('is_table'):
                        table_exists = True
                        continue

                    window.ax.text(
                        label_data['x'],
                        label_data['y'],
                        label_data['text'],
                        rotation=label_data.get('rotation', 90),
                        va='bottom',
                        ha='center',
                        fontsize=window.label_font_size
                    )

                # Redraw survey table if Table entity exists and it's a survey plot
                if table_exists and hasattr(self, 'is_survey_plot') and self.is_survey_plot():
                    if hasattr(self, 'survey_table_state') and self.survey_table_state == 1:
                        self.draw_survey_table()


            # Hide the cross if it exists
            if hasattr(window, 'cross') and window.cross:
                window.cross.remove()
            if hasattr(window, 'peak_letter') and window.peak_letter:
                window.peak_letter.remove()
                window.peak_letter = None
                del self.peak_letter
            if hasattr(window, 'peak_info') and window.peak_info:
                window.peak_info.remove()
                window.peak_info = None
                del self.peak_info


            # Switch to "None" ticked box and hide background lines
            window.vline1 = None
            window.vline2 = None
            window.vline3 = None
            window.vline4 = None
            window.show_hide_vlines()

            # Remove any existing fit or residual lines
            for line in self.ax.lines:
                if line.get_label() in ['Envelope', 'Residuals', 'Background']:
                    line.remove()

            for collection in self.ax.collections:
                if collection.get_label().startswith(window.sheet_combobox.GetValue()):
                    collection.remove()
            if 'Fitting' in window.Data['Core levels'][sheet_name] and 'Peaks' in \
                    window.Data['Core levels'][sheet_name]['Fitting']:
                peaks = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                if peaks and any(peak.get('Fitting Model') == 'D-parameter' for peak in peaks.values()):
                    is_d_parameter = True

            if "survey" in sheet_name.lower() or "wide" in sheet_name.lower() or is_d_parameter:
                pass
            else:
                self.ax.legend(loc='upper left', frameon=True, fancybox=True, framealpha=0.1, edgecolor='gray')

            # # Check if a peak is selected and add cross
            # if window.selected_peak_index is not None:
            #     window.plot_manager.add_cross_to_peak(window, window.selected_peak_index)

            # Remove any existing sheet name text
            for txt in self.ax.texts:
                if getattr(txt, 'sheet_name_text', False):
                    txt.remove()

            # Set appropriate axis labels and direction based on data type
            if is_raman:
                self.ax.set_xlim(min(x_values), max(x_values))  # Normal direction for Raman
                self.ax.set_ylabel("Intensity (a.u.)")
                self.ax.set_xlabel("Wavenumber (cm⁻¹)")

                # Clear existing sheet name text for Raman data
                for txt in self.ax.texts:
                    if getattr(txt, 'sheet_name_text', False):
                        txt.remove()
            elif is_xas:
                self.ax.set_xlim(min(x_values), max(x_values))  # Normal direction for XAS
                # self.ax.set_xlim(limits['Xmax'], limits['Xmin']) # Reverse X-axis for XAS
                self.ax.set_ylabel("Intensity (a.u.)")
                self.ax.set_xlabel("Photon Energy (eV)")

                # Clear existing sheet name text for XAS data
                for txt in self.ax.texts:
                    if getattr(txt, 'sheet_name_text', False):
                        txt.remove()

                # Format and add sheet name text for XAS
                formatted_sheet_name = self.format_xas_sheet_name(sheet_name)
                sheet_name_text = self.ax.text(
                    0.98, 0.98,  # Position (top-right corner)
                    formatted_sheet_name,
                    transform=self.ax.transAxes,
                    fontsize=15,
                    fontweight='bold',
                    verticalalignment='top',
                    horizontalalignment='right',
                    bbox=dict(facecolor='none', edgecolor='none', alpha=0.7),
                )
                sheet_name_text.sheet_name_text = True  # Mark this text object
            else:
                # Existing XPS plotting code with reverse x-axis
                self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis for XPS
                self.ax.set_ylabel("Intensity (CPS)")
                self.ax.set_xlabel("Binding Energy (eV)")

                # Format and add sheet name text for XPS
                formatted_sheet_name = self.format_sheet_name(sheet_name)
                sheet_name_text = self.ax.text(
                    0.98, 0.98,  # Position (top-right corner)
                    formatted_sheet_name,
                    transform=self.ax.transAxes,
                    fontsize=15,
                    fontweight='bold',
                    verticalalignment='top',
                    horizontalalignment='right',
                    bbox=dict(facecolor='none', edgecolor='none', alpha=0.7),
                )
                sheet_name_text.sheet_name_text = True  # Mark this text object

            self.apply_text_settings(window)

            # Plot background only if it meets the criteria
            if 'Background' in window.Data['Core levels'][sheet_name] and 'Bkg Y' in \
                    window.Data['Core levels'][sheet_name]['Background']:
                background_data = window.Data['Core levels'][sheet_name]['Background']
                raw_data = np.array(window.Data['Core levels'][sheet_name]['Raw Data'])
                background = np.array(background_data['Bkg Y'])

                # Use the same conditions as in clear_and_replot
                if (background_data.get('Bkg Type', '') != "" and
                        background_data.get('Bkg Low', '') != "" and
                        background_data.get('Bkg High', '') != ""):
                    if window.energy_scale == 'KE':
                        self.ax.plot(window.photons - x_values, window.background,
                                     color=self.background_color,
                                     linestyle=self.background_linestyle,
                                     alpha=self.background_alpha,
                                     linewidth=self.background_thickness,
                                     label='Background')
                    else:
                        self.ax.plot(x_values, window.background,
                                     color=self.background_color,
                                     linestyle=self.background_linestyle,
                                     alpha=self.background_alpha,
                                     linewidth=self.background_thickness,
                                     label='Background')

            if 'Background' not in window.Data['Core levels'][sheet_name]:
                window.Data['Core levels'][sheet_name]['Background'] = {
                    'Bkg Type': "",
                    'Bkg Low': "",
                    'Bkg High': "",
                    'Bkg Offset Low': "",
                    'Bkg Offset High': "",
                    'Bkg Y': window.y_values.tolist()
                }

            # Restore green line if active
            self.restore_green_vline_if_active(window)

            self.canvas.draw()  # Update the plot

        except Exception as e:
            wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)

    def plot_profile(self, window):
        """Plot profile data (zzProfile sheets) as line plot"""
        import matplotlib.pyplot as plt
        import pandas as pd
        import numpy as np

        sheet_name = window.sheet_combobox.GetValue()

        print(f"plot_profile called for: {sheet_name}")

        # Check if this is a profile sheet
        if not sheet_name.startswith('zzProfile'):
            print(f"Not a profile sheet: {sheet_name}")
            return False

        # Get profile metadata and data from window.Data
        if sheet_name not in window.Data['Core levels']:
            print(f"Sheet not in Data: {sheet_name}")
            return False

        profile_info = window.Data['Core levels'][sheet_name]
        print(f"Profile info keys: {profile_info.keys()}")

        if 'Profile Data' not in profile_info:
            print(f"No 'Profile Data' key found. Keys available: {profile_info.keys()}")
            return False

        # Extract metadata
        y_axis_label = profile_info.get('Y_Axis_Label', 'Atomic Concentration (%)')
        x_axis_label = profile_info.get('X_Axis_Label', 'Number')
        profile_data = profile_info['Profile Data']

        print(f"Y-axis: {y_axis_label}, X-axis: {x_axis_label}")
        print(f"Profile data keys: {profile_data.keys()}")

        # Convert new format back to DataFrame format for compatibility
        df_data = {}

        # Check if data is in new format (with x_values/y_values) or old format
        first_key = list(profile_data.keys())[0] if profile_data else None
        if first_key and isinstance(profile_data[first_key], dict) and 'x_values' in profile_data[first_key]:
            # New format: convert to DataFrame-compatible format
            print("Converting new format to DataFrame")

            # Find the maximum row number needed
            max_row = 0
            for col_name, col_info in profile_data.items():
                x_values = col_info.get('x_values', [])
                if x_values:
                    max_row = max(max_row, max(x_values) + 1)

            # Create Number column using actual y_values from Number data
            if 'Number' in profile_data and isinstance(profile_data['Number'], dict):
                number_info = profile_data['Number']
                number_x_values = number_info.get('x_values', [])
                number_y_values = number_info.get('y_values', [])

                # Create array for Number column
                number_array = np.full(max_row, np.nan)
                for x_pos, y_val in zip(number_x_values, number_y_values):
                    if 0 <= x_pos < max_row:
                        number_array[x_pos] = y_val
                df_data['Number'] = number_array
            else:
                # Fallback to sequential numbers
                df_data['Number'] = list(range(max_row))

            # Fill other columns
            for col_name, col_info in profile_data.items():
                if col_name == 'Number':
                    continue

                x_values = col_info.get('x_values', [])
                y_values = col_info.get('y_values', [])

                # Create array filled with NaN
                col_array = np.full(max_row, np.nan)

                # Fill in the actual values at their correct positions
                for x_pos, y_val in zip(x_values, y_values):
                    if 0 <= x_pos < max_row:
                        col_array[x_pos] = y_val

                df_data[col_name] = col_array
        else:
            # Old format: use as-is
            df_data = profile_data

        # Create DataFrame from processed data
        df = pd.DataFrame(df_data)

        # Sort by Number column to ensure correct line plotting
        if 'Number' in df.columns:
            df = df.sort_values(by='Number').reset_index(drop=True)

        # Clear current plot
        self.ax.clear()

        # Remove residuals subplot if it exists
        if hasattr(self, 'residuals_subplot'):
            if self.residuals_subplot:
                self.figure.delaxes(self.residuals_subplot)
                self.residuals_subplot = None
                # self.ax.set_position([0.1, 0.125, 0.85, 0.85])
                self.ax.set_position([0.1, 0.1, 0.85, 0.85])
                self.ax.get_xaxis().set_visible(True)

        # Get column names (excluding 'Number')
        peak_columns = [col for col in df.columns if col != 'Number']

        print(f"Plotting {len(peak_columns)} peaks: {peak_columns}")

        # Define colors for different peaks
        colors = plt.cm.tab10(range(len(peak_columns)))

        # Get profile settings
        linewidth = getattr(window, 'profile_linewidth', {}).get(sheet_name, 2.0)
        lines_visible = getattr(window, 'profile_lines_visible', {}).get(sheet_name, True)
        interp_type = getattr(window, 'profile_interpolation', {}).get(sheet_name, 'linear')

        # Plot each peak as a line
        for idx, column in enumerate(peak_columns):
            # Extract peak name (remove " At(%)" or other suffix)
            peak_name = column.split(' ')[0] if ' ' in column else column

            x_data = df['Number'].values
            y_data = df[column].values

            # Remove NaN values for plotting
            mask = ~np.isnan(y_data)
            x_data = x_data[mask]
            y_data = y_data[mask]

            # Skip if no valid data
            if len(x_data) == 0:
                continue

            # Determine line style based on settings
            if interp_type == 'markers' or not lines_visible:
                linestyle = ''
                plot_x, plot_y = x_data, y_data
            elif interp_type == 'cubic':
                # Cubic spline interpolation for smooth curves
                from scipy.interpolate import make_interp_spline
                import numpy as np

                # Create smooth curve with more points
                x_smooth = np.linspace(x_data.min(), x_data.max(), 300)
                try:
                    spl = make_interp_spline(x_data, y_data, k=3)  # k=3 for cubic
                    y_smooth = spl(x_smooth)
                    plot_x, plot_y = x_smooth, y_smooth
                    linestyle = '-'
                except:
                    # Fall back to linear if spline fails
                    plot_x, plot_y = x_data, y_data
                    linestyle = '-'
            else:  # linear
                linestyle = '-'
                plot_x, plot_y = x_data, y_data

            self.ax.plot(
                plot_x,
                plot_y,
                marker='o' if interp_type != 'cubic' else '',  # No markers for smooth curves
                linestyle=linestyle,
                linewidth=linewidth if linestyle else 0,
                markersize=6,
                label=peak_name,
                color=colors[idx]
            )

            # Add markers separately for cubic to show data points
            if interp_type == 'cubic':
                self.ax.plot(x_data, y_data, 'o', markersize=4, color=colors[idx])

            print(f"Plotted: {peak_name}")

        # Configure plot - NO TITLE, NO GRID as requested
        self.ax.set_xlabel(x_axis_label, fontsize=12)
        self.ax.set_ylabel(y_axis_label, fontsize=12)

        # Create legend and make it explicitly visible
        legend = self.ax.legend(loc='best', fontsize=10, frameon=True, fancybox=True, framealpha=0.9, edgecolor='gray')
        legend.set_visible(True)  # Explicitly set visible

        # Format axes
        from matplotlib.ticker import ScalarFormatter
        self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.ax.xaxis.set_major_formatter(ScalarFormatter(useMathText=True))

        # Only use scientific notation for Y values above 1000
        y_min, y_max = self.ax.get_ylim()
        if y_max > 1000:
            self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
        else:
            self.ax.ticklabel_format(style='plain', axis='y')

        # Ensure legend stays visible (override any other settings)
        if self.ax.get_legend():
            self.ax.get_legend().set_visible(True)

        print("Profile plot complete")

        # Redraw canvas
        self.canvas.draw_idle()

        return True


    def _plot_xps_map_heatmap(self, window, map_sheet_name):
        """
        Plot any *~Map sheet (XPS~Map, C1s~Map, etc.) as a 2D heatmap on the main window.
        Called from clear_and_replot when the combobox selects a ~Map sheet.
        Uses the same fixed-colorbar approach as FileManager.plot_xps_map_on_main to
        prevent axes shrinkage on repeated calls.
        """
        import numpy as np

        map_data = window.Data['Core levels'].get(map_sheet_name, {})
        be_values = np.array(map_data.get('B.E.', []))
        num_sweeps = map_data.get('_num_sweeps', 0)
        if num_sweeps == 0:
            num_sweeps = sum(1 for k in map_data if k.startswith('Y') and k[1:].isdigit())

        if num_sweeps == 0 or len(be_values) == 0:
            self.ax.clear()
            self.ax.text(0.5, 0.5, f'{map_sheet_name}\n\nNo map data available',
                         ha='center', va='center', transform=self.ax.transAxes,
                         fontsize=12, color='gray')
            self.ax.set_xticks([])
            self.ax.set_yticks([])
            self.canvas.draw_idle()
            return

        # Determine normalisation mode: defer to FileManager combo if open, else default Norm. Auto
        norm_mode = "Norm. Auto"
        if hasattr(window, 'file_manager') and window.file_manager is not None:
            try:
                norm_mode = window.file_manager.norm_type.GetValue()
            except Exception:
                norm_mode = "Norm. Auto"

        # Build 2D array (num_sweeps × num_be_points), normalising each row to 0-1000 if requested
        data_2d = np.zeros((num_sweeps, len(be_values)))
        for i in range(num_sweeps):
            col_name = f'Y{i + 1}'
            if col_name in map_data:
                row = np.array(map_data[col_name], dtype=float)
                if norm_mode == "Norm. Auto":
                    r_min, r_max = row.min(), row.max()
                    if r_max - r_min > 0:
                        row = (row - r_min) / (r_max - r_min) * 1000.0
                    else:
                        row = np.full_like(row, 500.0)
                data_2d[i, :] = row

        # --- clear axes and any existing colorbar ---
        self.ax.clear()
        if hasattr(window, 'heatmap_colorbar') and window.heatmap_colorbar is not None:
            try:
                if hasattr(window.heatmap_colorbar, 'ax'):
                    window.figure.delaxes(window.heatmap_colorbar.ax)
                window.heatmap_colorbar = None
            except Exception:
                pass
        if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
            try:
                window.figure.delaxes(self.residuals_subplot)
            except Exception:
                pass
            self.residuals_subplot = None
            self.ax.get_xaxis().set_visible(True)

        # Pin axes to fixed position before drawing
        self.ax.set_position([0.1, 0.1, 0.73, 0.85])

        # --- draw heatmap ---
        cmap_name = getattr(window, 'heatmap_colormap', 'viridis')
        be_descending = be_values[0] > be_values[-1] if len(be_values) > 1 else False

        vmin_h = 0 if norm_mode == "Norm. Auto" else data_2d.min()
        vmax_h = 1000.0 if norm_mode == "Norm. Auto" else data_2d.max()
        if be_descending:
            heatmap_img = self.ax.imshow(
                data_2d, aspect='auto', origin='lower',
                extent=[float(be_values[0]), float(be_values[-1]), 0, num_sweeps],
                cmap=cmap_name, vmin=vmin_h, vmax=vmax_h)
        else:
            data_flipped = np.fliplr(data_2d)
            heatmap_img = self.ax.imshow(
                data_flipped, aspect='auto', origin='lower',
                extent=[float(be_values.max()), float(be_values.min()), 0, num_sweeps],
                cmap=cmap_name, vmin=vmin_h, vmax=vmax_h)

        self.ax.set_ylim(0, num_sweeps)
        self.ax.set_xlabel('Binding Energy (eV)',
                           fontsize=getattr(window, 'axis_title_size', 10))
        self.ax.set_ylabel('Sweep Number',
                           fontsize=getattr(window, 'axis_title_size', 10))
        self.ax.tick_params(axis='both',
                            labelsize=getattr(window, 'axis_number_size', 9))

        base_name = (map_data.get('_core_level', '')
                     or map_data.get('ExperimentalInfo', {}).get('Core Level', '')
                     or map_sheet_name)
        self.ax.set_title(f'XPS Map \u2013 {base_name}',
                          fontsize=getattr(window, 'axis_title_size', 11))

        # Fixed-position colorbar (avoids axes-shrink bug from ax= form)
        try:
            cbar_ax = window.figure.add_axes([0.84, 0.1, 0.03, 0.85])
            cbar = window.figure.colorbar(heatmap_img, cax=cbar_ax)
            cbar_label = 'Norm. Intensity' if norm_mode == "Norm. Auto" else 'Intensity (CPS)'
            cbar.set_label(cbar_label, rotation=270, labelpad=20,
                           fontsize=getattr(window, 'axis_title_size', 9))
            cbar.ax.tick_params(labelsize=getattr(window, 'axis_number_size', 9))
            window.heatmap_colorbar = cbar
            window.heatmap_cbar_ax = cbar_ax
        except Exception:
            window.heatmap_colorbar = None

        # Re-assert position after colorbar creation
        self.ax.set_position([0.1, 0.1, 0.73, 0.85])

        # Mark that this map is currently displayed so save_plot_to_excel can detect it
        window.xps_map_on_main = map_sheet_name

        self.canvas.draw_idle()

    def clear_and_replot(self, window):
        """
        Clears the current plot and redraws all elements for the selected sheet.

        This function performs the following key operations:
        1. Retrieves the current sheet name and plot limits.
        2. Clears the existing plot and updates axis labels and formatting.
        3. Plots raw data according to the selected energy scale (BE or KE).
        4. Identifies and plots doublet peaks with appropriate coloring.
        5. Plots fitted peaks using stored parameters, handling different fitting models.
        6. Plots the background if available.
        7. Updates overall fit and residuals for fitted data.
        8. Handles special cases for survey/wide scans.
        9. Updates the plot legend and restores or creates sheet name text.
        10. Adjusts plot limits and spine widths for better visibility.
        11. Redraws the canvas to reflect all changes.

        The function adapts its behavior based on the sheet type (e.g., survey vs. core level),
        energy scale, and fitting status of peaks. It ensures that all visual elements
        are correctly positioned and formatted according to the current application state.
        """

        # Remove heatmap colorbar AND its axes if it exists
        if hasattr(window, 'heatmap_colorbar') and window.heatmap_colorbar is not None:
            try:
                # Remove the colorbar axes from the figure
                if hasattr(window.heatmap_colorbar, 'ax'):
                    window.figure.delaxes(window.heatmap_colorbar.ax)
                window.heatmap_colorbar = None
            except:
                pass

        sheet_name = window.sheet_combobox.GetValue()
        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper()
        is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')

        if not sheet_name or 'Core levels' not in window.Data or sheet_name not in window.Data['Core levels']:
            return
        limits = window.plot_config.get_plot_limits(window, sheet_name)
        cst_unfit = ""

        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        if sheet_name == 'EDX~Map':
            print(f"Detected EDX Map sheet: {sheet_name}")
            self.ax.clear()
            self.ax.text(0.5, 0.5, 'EDX Map\n\nOpen EDX Analysis window\nto view the map',
                         ha='center', va='center', transform=self.ax.transAxes,
                         fontsize=12, color='gray')
            self.ax.set_xticks([])
            self.ax.set_yticks([])
            self.canvas.draw_idle()
            return

        if sheet_name == 'EELS~Map':
            print(f"Detected EELS Map sheet: {sheet_name}")
            # Try to open EELS window with the map
            dm3_path = window.Data['Core levels'].get('EELS~Map', {}).get('_DM3_Path')
            print(f"DM3 path from Data: {dm3_path}")

            # If no path stored, try to find the DM3 file based on Excel filename
            if not dm3_path or not os.path.exists(str(dm3_path) if dm3_path else ''):
                excel_path = window.Data.get('FilePath', '')
                if excel_path:
                    # Try various naming patterns
                    for pattern in ['.dm3', '.dm4', '_EELS.dm3', '_EELS.dm4']:
                        test_path = excel_path.replace('_EELS.xlsx', pattern).replace('.xlsx', pattern)
                        if os.path.exists(test_path):
                            dm3_path = test_path
                            print(f"Found DM3 file: {dm3_path}")
                            break

            if dm3_path and os.path.exists(str(dm3_path)):
                try:
                    if hasattr(window, 'eels_window') and window.eels_window:
                        window.eels_window.load_file(dm3_path)
                        window.eels_window.Show()
                        window.eels_window.Raise()
                    else:
                        from libraries.ToolsMenu.EELS_Analysis import open_eels_window
                        eels_window = open_eels_window(window)
                        if eels_window:
                            window.eels_window = eels_window
                            eels_window.load_file(dm3_path)
                except Exception as e:
                    print(f"Error opening EELS window: {e}")
                    import traceback
                    traceback.print_exc()

            # Show placeholder in main plot
            self.ax.clear()
            self.ax.text(0.5, 0.5, 'EELS Map\n\nOpen EELS Analysis window\nto view the map',
                         ha='center', va='center', transform=self.ax.transAxes,
                         fontsize=12, color='gray')
            self.ax.set_xticks([])
            self.ax.set_yticks([])
            self.canvas.draw_idle()
            return

        # Handle any *~Map sheet (XPS~Map, C1s~Map, etc.) — plot as 2D heatmap on main window
        if '~Map' in sheet_name and sheet_name not in ('EDX~Map', 'EELS~Map'):
            self._plot_xps_map_heatmap(window, sheet_name)
            return

        # CHECK IF THIS IS A PROFILE SHEET - if so, use profile plotting
        if sheet_name.startswith('zzProfile'):
            print(f"Detected profile sheet: {sheet_name}")
            result = self.plot_profile(window)
            print(f"Profile plotting result: {result}")
            if result:
                return  # Profile plotted successfully, exit early
            else:
                print("Profile plotting failed, falling back to normal plot")

        core_level_data = window.Data['Core levels'][sheet_name]

        # Handle x-axis direction based on data type
        if is_raman:
            self.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal direction for Raman
        elif window.energy_scale == 'KE':
            self.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])
        else:
            self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis for XPS

        self.ax.set_ylim(limits['Ymin'], limits['Ymax'])

        # Store existing sheet name text
        sheet_name_text = None
        for txt in self.ax.texts:
            if getattr(txt, 'sheet_name_text', False):
                sheet_name_text = txt
                break

        # Clear the plot
        self.ax.clear()

        if hasattr(self, 'residuals_state') and self.residuals_state == 2:
            if self.residuals_subplot:
                self.residuals_subplot.clear()
                # self.ax.set_position([0.1, 0.1, 0.8, 0.8])
                # self.ax.set_position([0.1, 0.125, 0.85, 0.85])
                self.ax.set_position([0.1, 0.1, 0.85, 0.85])
                # gs = self.figure.add_gridspec(20, 1, hspace=0.0)
                gs = self.figure.add_gridspec(20, 1, hspace=0.0, top=0.95, bottom=0.1, left=0.1, right=0.95)
                self.ax.set_position(gs[0:18, 0].get_position(self.figure))  # Main plot takes 6/8
                self.residuals_subplot.set_position(gs[18:, 0].get_position(self.figure))  # Residuals takes 2/8
                self.residuals_subplot.sharex(self.ax)

                # Explicitly ensure visibility
                self.residuals_subplot.set_visible(True)

        else:
            # self.ax.set_position([0.1, 0.1, 0.8, 0.8])
            # self.ax.set_position([0.1, 0.125, 0.85, 0.85])
            self.ax.set_position([0.1, 0.1, 0.85, 0.85])
            self.ax.get_xaxis().set_visible(True)

        # Set appropriate axis labels based on data type
        if is_raman:
            self.ax.set_xlabel("Wavenumber (cm⁻¹)")
            self.ax.set_ylabel("Intensity (a.u.)")
        elif is_xas:
            self.ax.set_xlabel("Photon Energy (eV)")
            self.ax.set_ylabel("Intensity (a.u.)")

        elif window.energy_scale == 'KE':
            self.ax.set_xlabel("Kinetic Energy (eV)")
            self.ax.set_ylabel("Intensity (CPS)")
        else:
            self.ax.set_xlabel(window.x_axis_label)
            self.ax.set_ylabel("Intensity (CPS)")

        self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        # Plot the raw data
        x_values = np.array(core_level_data['B.E.'])
        y_values = np.array(core_level_data['Raw Data'])

        # Get plot limits from PlotConfig
        if sheet_name not in window.plot_config.plot_limits:
            window.plot_config.update_plot_limits(window, sheet_name)
        limits = window.plot_config.plot_limits[sheet_name]

        # Set plot limits based on data type
        if is_raman:
            self.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal direction for Raman
        elif is_xas:
            self.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal direction for Raman
        elif window.energy_scale == 'KE':
            X_MIN = window.photons - limits['Xmax']
            X_MAX = window.photons - limits['Xmin']
            self.ax.set_xlim(min(X_MIN, X_MAX), max(X_MIN, X_MAX))
        else:
            self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis for XPS

        self.ax.set_ylim(limits['Ymin'], limits['Ymax'])

        # Create a color cycle
        colors = plt.cm.tab10(np.linspace(0, 1, 10))
        color_cycle = cycle(colors)

        # Replot all peaks and update indices
        num_peaks = window.peak_params_grid.GetNumberRows() // 2  # Assuming each peak uses two rows

        # First, identify all doublets
        doublets = []
        for i in range(0, num_peaks - 1):
            current_label = window.peak_params_grid.GetCellValue(i * 2, 1)
            next_label = window.peak_params_grid.GetCellValue((i + 1) * 2, 1)
            if self.is_part_of_doublet(current_label, next_label):
                doublets.extend([i, i + 1])

        for i in range(num_peaks):
            row = i * 2

            # Get all necessary values, with error checking
            try:
                fitting_model = window.peak_params_grid.GetCellValue(row, 13)
                position = float(window.peak_params_grid.GetCellValue(row, 2))
                height = float(window.peak_params_grid.GetCellValue(row, 3))
                fwhm = float(window.peak_params_grid.GetCellValue(row, 4))
                fraction = float(window.peak_params_grid.GetCellValue(row, 5))
                if fitting_model in ["Voigt (Area, L/G, \u03c3)","Voigt (Area, \u03c3, \u03b3)",
                                     "ExpGauss.(Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
                    sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                    gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                    if fitting_model == "Voigt (Area, L/G, \u03c3, S)":
                        skew = float(window.peak_params_grid.GetCellValue(row, 9))
                else:
                    sigma = 0
                    gamma = 0
                    skew = 0.01
                label = window.peak_params_grid.GetCellValue(row, 1)

            except ValueError:
                print(f"Warning: Invalid data for peak {i + 1}. Skipping this peak.")
                continue
            if fitting_model == "D-parameter":
                cst_unfit = "D-parameter"
            elif fitting_model == "Fermi":
                cst_unfit = "Fermi"
            elif fitting_model == "VBM":
                cst_unfit = "VBM"
            elif fitting_model == "Cut-Off":
                cst_unfit = "Cut Off"
            elif fitting_model == "SurveyID":
                cst_unfit = "SurveyID"
            if fitting_model == "Unfitted":
                cst_unfit = "Unfitted"
                XrangeMin = float(window.peak_params_grid.GetCellValue(row, 15))
                XrangeMax = float(window.peak_params_grid.GetCellValue(row, 16))
                color = self.peak_colors[i % len(self.peak_colors)]

                if window.energy_scale == 'KE':
                    KErangeMin = window.photons - XrangeMax  # Convert BE range to KE range
                    KErangeMax = window.photons - XrangeMin
                    x_filtered = np.where((window.photons - x_values >= KErangeMin) &
                                          (window.photons - x_values <= KErangeMax),
                                          window.photons - x_values, np.nan)
                    y_filtered = np.where((window.photons - x_values >= KErangeMin) &
                                          (window.photons - x_values <= KErangeMax),
                                          y_values, np.nan)
                    background_filtered = np.where((window.photons - x_values >= KErangeMin) &
                                                   (window.photons - x_values <= KErangeMax),
                                                   window.background, np.nan)
                else:
                    x_filtered = np.where((x_values >= XrangeMin) & (x_values <= XrangeMax),
                                          x_values, np.nan)
                    y_filtered = np.where((x_values >= XrangeMin) & (x_values <= XrangeMax),
                                          y_values, np.nan)
                    background_filtered = np.where((x_values >= XrangeMin) & (x_values <= XrangeMax),
                                                   window.background, np.nan)

                self.ax.fill_between(x_filtered, background_filtered, y_filtered,
                                     where=~np.isnan(x_filtered),
                                     facecolor=color, alpha=0.5, label=label)

            else:
                if i in doublets:
                    if doublets.index(i) % 2 == 0:  # First peak of the doublet
                        color = self.peak_colors[i % len(self.peak_colors)]
                        alpha = self.peak_alpha
                    else:  # Second peak of the doublet
                        # Use the same color as the first peak of the doublet, but with lower alpha
                        color = self.peak_colors[(i - 1) % len(self.peak_colors)]
                        alpha = self.peak_alpha * 0.99  # Reduce alpha for the second peak
                else:
                    color = self.peak_colors[i % len(self.peak_colors)]
                    alpha = self.peak_alpha

                # For fitted peaks, use the existing plot_peak method
                peak_params = {
                    'row': row,
                    'position': position,
                    'height': height,
                    'fwhm': fwhm,
                    'lg_ratio': fraction,
                    'sigma': sigma,
                    'gamma': gamma,
                    'label': label,
                    # 'skew' : skew,
                    'fitting_model': fitting_model
                }
                if window.energy_scale == 'KE':
                    self.plot_peak(window.x_values, window.background, peak_params, sheet_name, window,
                                                color=color, alpha=alpha)
                else:
                    self.plot_peak(window.x_values, window.background, peak_params, sheet_name, window,
                                                color=color, alpha=alpha)

        # Only plot background if it's different from raw data or if Bkg Type is not empty.
        # Skip for Fermi / D-parameter / VBM / Cut Off sheets — those tools draw their own
        # fit overlays and a generic background line is misleading.
        if (core_level_data['Background'].get('Bkg Type') not in ("", "None") and
                core_level_data['Background'].get('Bkg Low') != "" and
                core_level_data['Background'].get('Bkg High') != "" and
                cst_unfit not in ("Fermi", "D-parameter", "VBM", "Cut Off")):

            # Get background regions for masking
            bg_regions = self.get_background_regions(window)

            # Apply background region masking to background line
            bg_x = x_values.copy()
            bg_y = np.array(core_level_data['Background']['Bkg Y'])

            if bg_regions is not None and len(bg_regions) > 0:
                # Create mask for all background regions
                region_mask = np.zeros(len(bg_x), dtype=bool)
                for bg_start, bg_end in bg_regions:
                    region_mask |= (bg_x >= bg_start) & (bg_x <= bg_end)

                # Instead of removing points, set non-region points to NaN
                bg_y_masked = bg_y.copy()
                # Ensure the array can hold NaN values by converting to float
                bg_y_masked = bg_y_masked.astype(float)
                bg_y_masked[~region_mask] = np.nan

                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons - bg_x, bg_y_masked,
                                 color=self.background_color, linewidth=self.background_thickness,
                                 linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background')
                else:
                    self.ax.plot(bg_x, bg_y_masked, color=self.background_color,
                                 linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background',
                                 linewidth=self.background_thickness)
            else:
                # Plot full background if no regions defined
                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons - bg_x, bg_y,
                                 color=self.background_color, linewidth=self.background_thickness,
                                 linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background')
                else:
                    self.ax.plot(bg_x, bg_y, color=self.background_color,
                                 linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background',
                                 linewidth=self.background_thickness)



        # Update overall fit and residuals
        if cst_unfit in ["Unfitted","D-parameter","Fermi","VBM", "Cut Off","SurveyID"] or any(x in sheet_name.lower()
                                                                                            for x in [
            "survey", \
                "wide"]):
            pass
        else:
            window.update_overall_fit_and_residuals()

        # When plotting raw data
        if "survey" in sheet_name.lower() or "wide" in sheet_name.lower():
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - x_values, y_values, c=self.line_color, linewidth=self.line_width,
                         alpha=self.line_alpha, linestyle=self.raw_data_linestyle)
            else:
                self.ax.plot(x_values, y_values, c=self.line_color, linewidth=self.line_width,
                         alpha=self.line_alpha, linestyle=self.raw_data_linestyle)
        elif self.plot_style == "scatter":
            if window.energy_scale == 'KE':
                self.ax.scatter(window.photons - x_values, y_values, c=self.scatter_color, s=self.scatter_size,
                            marker=self.scatter_marker, label='Raw Data')
            else:
                self.ax.scatter(x_values, y_values, c=self.scatter_color, s=self.scatter_size,
                                marker=self.scatter_marker, label='Raw Data')
        else:
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - x_values, y_values, c=self.line_color, linewidth=self.line_width,
                             alpha=self.line_alpha, linestyle=self.raw_data_linestyle, label='Raw Data')
            else:
                self.ax.plot(x_values, y_values, c=self.line_color, linewidth=self.line_width,
                             alpha=self.line_alpha, linestyle=self.raw_data_linestyle, label='Raw Data')

        # Assuming 'ax' is your axes object
        for spine in self.ax.spines.values():
            spine.set_linewidth(1)  # Adjust this value to increase or decrease thickness


        # Update the legend
        if self.legend_visible:
            self.ax.legend(loc='upper left', frameon=True, fancybox=True, framealpha=0.1, edgecolor='gray')
            self.update_legend(window)
        else:
            self.ax.legend().set_visible(False)

        # Only add sheet name text for non-Raman data
        if not is_raman:
            if sheet_name_text is None:
                # Format sheet name based on data type
                if is_xas:
                    formatted_sheet_name = self.format_xas_sheet_name(sheet_name)
                else:
                    formatted_sheet_name = self.format_sheet_name(sheet_name)

                sheet_name_text = self.ax.text(
                    0.98, 0.98,  # Position (top-right corner)
                    formatted_sheet_name,
                    transform=self.ax.transAxes,
                    fontsize=15,
                    fontweight='bold',
                    verticalalignment='top',
                    horizontalalignment='right',
                    bbox=dict(facecolor='none', edgecolor='none', alpha=0.7),
                )
                sheet_name_text.sheet_name_text = True  # Mark this text object
            else:
                self.ax.add_artist(sheet_name_text)

        # Apply text settings and continue with existing functionality
        self.apply_text_settings(window)

        # Draw labels after all peaks are plotted
        if 'Labels' in window.Data['Core levels'][sheet_name]:
            table_exists = False

            for label_data in window.Data['Core levels'][sheet_name]['Labels']:
                # Skip drawing Table entities - they are handled separately
                if label_data.get('text') == 'Table' and label_data.get('is_table'):
                    table_exists = True
                    continue

                window.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=label_data.get('rotation', 90),
                    va='bottom',
                    ha='center',
                    fontsize=window.label_font_size
                )

            # Redraw survey table if Table entity exists and it's a survey plot
            if table_exists and hasattr(self, 'is_survey_plot') and self.is_survey_plot():
                if hasattr(self, 'survey_table_state') and self.survey_table_state == 1:
                    self.draw_survey_table()

        # Apply y-axis state
        if hasattr(self, 'y_axis_state'):
            if self.y_axis_state == 1:  # Hidden
                self.ax.yaxis.set_visible(False)
            elif self.y_axis_state == 2:  # Label only
                self.ax.yaxis.set_visible(True)
                self.ax.set_ylabel("Intensity (a.u.)")
                for label in self.ax.get_yticklabels():
                    label.set_visible(False)

        # Restore vlines if fitting window is open and background tab is selected
        if (hasattr(window, 'fitting_window') and window.fitting_window is not None and
                hasattr(window, 'background_tab_selected') and window.background_tab_selected):
            window.show_hide_vlines()

        # Call survey table drawing if enabled and it's a survey plot
        if (hasattr(self, 'survey_table_state') and
                self.survey_table_state == 1 and
                self.is_survey_plot()):
            self.draw_survey_table()

        # Force position before drawing to prevent matplotlib auto-layout
        if hasattr(self, 'residuals_state') and self.residuals_state == 2 and self.residuals_subplot:
            # When residuals subplot exists, use GridSpec layout
            gs = self.figure.add_gridspec(20, 1, hspace=0.0, top=0.95, bottom=0.1, left=0.1, right=0.95)
            self.ax.set_position(gs[0:18, 0].get_position(self.figure))
            self.residuals_subplot.set_position(gs[18:, 0].get_position(self.figure))
        else:
            # No residuals - use fixed position
            self.ax.set_position([0.1, 0.1, 0.85, 0.85])


        # Draw the canvas
        self.canvas.draw_idle()


        # Make sure checkboxes retain their state
        wx.CallAfter(window.update_checkboxes_from_data)


    def is_part_of_doublet(self, current_label, next_label):
        """
        Determines if two adjacent peaks form a doublet based on their labels.
        Checks for matching core levels and appropriate spin-orbit components.
        """
        current_parts = current_label.split()
        next_parts = next_label.split()

        if len(current_parts) < 1 or len(next_parts) < 1:
            return False

        # Extract core level without spin-orbit component
        def extract_core_level(label):
            match = re.match(r'([A-Za-z]+\d+[spdf])', label)
            return match.group(1) if match else label

        current_core_level = extract_core_level(current_parts[0])
        next_core_level = extract_core_level(next_parts[0])

        if current_core_level != next_core_level:
            return False

        orbital = re.search(r'\d([spdf])', current_core_level)
        if not orbital:
            return False

        orbital = orbital.group(1)

        # Check for spin-orbit components in either the first or second part of the label
        def has_component(parts, component):
            return any(component in part for part in parts)

        if orbital == 'p':
            return ((has_component(current_parts, '3/2') and has_component(next_parts, '1/2')))  # or \
            # (has_component(current_parts, '1/2') and has_component(next_parts, '3/2')))
        elif orbital == 'd':
            return ((has_component(current_parts, '5/2') and has_component(next_parts, '3/2'))) # or \
                # (has_component(current_parts, '3/2') and has_component(next_parts, '5/2'))
        elif orbital == 'f':
            return ((has_component(current_parts, '7/2') and has_component(next_parts, '5/2'))) # or \
                # (has_component(current_parts, '5/2') and has_component(next_parts, '7/2'))

        return False

    def update_peak_plot(self, window, x, y, remove_old_peaks=True):
        """
        Updates the plot for a selected peak when its position or height is changed.
        Recalculates the peak shape based on the current fitting model and parameters.
        """
        # Skip for EDX sheets - they don't use XPS peak fitting
        if hasattr(window, 'sheet_combobox'):
            current_sheet = window.sheet_combobox.GetValue()
            if current_sheet.startswith('EDX~'):
                return

        if window.x_values is None or window.background is None:
            print("Error: x_values or background is None. Cannot update peak plot.")
            return

        if x is None or y is None:
            return
        if len(window.x_values) != len(window.background):
            print(f"Warning: x_values and background have different lengths. "
                  f"x: {len(window.x_values)}, background: {len(window.background)}")

        if window.selected_peak_index is not None and 0 <= window.selected_peak_index < window.peak_params_grid.GetNumberRows() // 2:
            row = window.selected_peak_index * 2
            peak_label = window.peak_params_grid.GetCellValue(row, 1)  # Get the current label

            # Get peak parameters from the grid
            fwhm = float(window.peak_params_grid.GetCellValue(row, 4))
            lg_ratio = float(window.peak_params_grid.GetCellValue(row, 5))
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = lg_ratio/100 * sigma
            bkg_y = window.background[np.argmin(np.abs(window.x_values - x))]

            def try_float(value, default=0.0):
                try:
                    return float(value)
                except ValueError:
                    return default

            sigma = try_float(window.peak_params_grid.GetCellValue(row, 7),0)
            gamma = try_float(window.peak_params_grid.GetCellValue(row, 8),0)



            # Ensure height is not negative
            y = max(y, 0)

            # Create the selected peak using updated position and height
            if window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
                peak_model = lmfit.models.VoigtModel()
                amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0)
                params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma)
            elif window.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
                peak_model = lmfit.models.SkewedVoigtModel()
                skew = float(window.peak_params_grid.GetCellValue(row, 9))
                amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0, skew=skew)
                params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma, skew=skew)
            elif window.selected_fitting_method in ["ExpGauss.(Area, \u03c3, \u03b3)"]:
                peak_model = lmfit.models.VoigtModel()
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                # amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0)
                params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, gamma=gamma)
            elif window.selected_fitting_method == "Pseudo-Voigt (Area)":
                peak_model = lmfit.models.PseudoVoigtModel()
                amplitude = y / peak_model.eval(center=0, amplitude=1, sigma=sigma, fraction=lg_ratio, x=0)
                params = peak_model.make_params(center=x, amplitude=amplitude, sigma=sigma, fraction=lg_ratio)
            elif window.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
                peak_model = lmfit.Model(PeakFunctions.LA)
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                # area = float(window.peak_params_grid.GetCellValue(row, 6))
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                params = peak_model.make_params(center=x,amplitude=amplitude,fwhm=fwhm,sigma=sigma,gamma=gamma)
            elif window.selected_fitting_method in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                peak_model = lmfit.Model(PeakFunctions.LAxG)
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                area = float(window.peak_params_grid.GetCellValue(row, 6))
                params = peak_model.make_params(center=x,amplitude=area,fwhm=fwhm,sigma=sigma,gamma=gamma)
            elif window.selected_fitting_method == "GL (Height)":
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz)
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)
            elif window.selected_fitting_method == "SGL (Height)":
                peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz)
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)
            elif window.selected_fitting_method == "GL (Area)":
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz_Area)
                area = y * fwhm * np.sqrt(np.pi / (4 * np.log(2)))  # Calculate area from height and FWHM
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, area=area)
            elif window.selected_fitting_method == "SGL (Area)":
                peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz_Area)
                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                gamma = fwhm / 2
                area = y * ((1 - lg_ratio / 100) * sigma * np.sqrt(2 * np.pi) + (lg_ratio / 100) * np.pi * gamma)
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, area=area)
            elif window.selected_fitting_method in ["D-parameter", "Fermi", "VBM", "Cut-Off"]:
                # return area, 0, 0  # Return original area and zero for normalized/relative areas
                return 0, 0, 0  # Return original area and zero for normalized/relative areas
            else:  # Default to GL (Height) as a safe bet
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz)
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)

            peak_y = peak_model.eval(params, x=window.x_values) + window.background

            # Update overall fit and residuals
            if peak_model in ["D-parameter", "SurveyID", "Fermi", "VBM", "Cut-Off"]:
                print("")
            else:
                window.update_overall_fit_and_residuals()

            peak_label = window.peak_params_grid.GetCellValue(row, 1)  # Get peak label from grid

            # # Update the selected peak plot line
            # for line in self.ax.get_lines():
            #     if line.get_label() == peak_label:
            #         line.set_ydata(peak_y)
            #         break
            #     else:
            #         print("I AM NOT SURE WHAT THIS DO")
            #         self.ax.plot(window.x_values, peak_y, label=peak_label)

            # Remove previous squares
            for line in self.ax.get_lines():
                if 'Selected Peak Center' in line.get_label():
                    line.remove()

            # Plot the new red square at the top center of the selected peak
            self.ax.plot(x, y + bkg_y, 'bx', label=f'Selected Peak Center {window.selected_peak_index}', markersize=15,
                         markerfacecolor='none')

            # Update the grid with new values
            window.peak_params_grid.SetCellValue(row, 2, f"{x:.2f}")  # Position
            window.peak_params_grid.SetCellValue(row, 3, f"{y:.2f}")  # Height
            window.peak_params_grid.SetCellValue(row, 4, f"{fwhm:.2f}")  # FWHM
            window.peak_params_grid.SetCellValue(row, 5, f"{lg_ratio:.2f}")  # L/G ratio

            self.canvas.draw_idle()

    def update_overall_fit_and_residuals(self, window):
        """
        Recalculates and updates the overall fit and residuals for all peaks.
        Handles different fitting models for each peak and scales residuals for better visibility.
        """
        # Calculate the overall fit as the sum of all peaks
        # Ensure all arrays have same length at the start
        x_values = window.x_values
        y_values = window.y_values[:len(x_values)]
        overall_fit = window.background.astype(float).copy()[:len(x_values)]

        num_peaks = window.peak_params_grid.GetNumberRows() // 2  # Assuming each peak uses two rows

        # Check if peaks exist in the peak fitting grids
        if num_peaks == 0:
            if hasattr(self, 'rsd_text') and self.rsd_text:
                self.rsd_text.remove()
                self.rsd_text = None
            return

        fitting_model = ""  # Default empty string

        for i in range(num_peaks):
            row = i * 2  # Each peak uses two rows in the grid

            # Get cell values
            position_str = window.peak_params_grid.GetCellValue(row, 2)  # Position
            height_str = window.peak_params_grid.GetCellValue(row, 3)  # Height
            fwhm_str = window.peak_params_grid.GetCellValue(row, 4)  # FWHM
            lg_ratio_str = window.peak_params_grid.GetCellValue(row, 5)  # L/G
            # sigma = window.peak_params_grid.GetCellValue(row, 7)
            # gamma = window.peak_params_grid.GetCellValue(row, 8)
            fitting_model = window.peak_params_grid.GetCellValue(row, 13)  # Fitting Model

            # Check if any of the cells are empty
            if not all([position_str, height_str, fwhm_str, lg_ratio_str, fitting_model]):
                print(f"Warning: Incomplete data for peak {i + 1}. Skipping this peak.")
                continue

            try:
                peak_x = float(position_str)
                peak_y = float(height_str)
                fwhm = float(fwhm_str)
                lg_ratio = float(lg_ratio_str)
            except ValueError:
                print(f"Warning: Invalid data for peak {i + 1}. Skipping this peak.")
                continue

            # sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            # gamma = lg_ratio/100 * sigma
            bkg_y = window.background[np.argmin(np.abs(window.x_values - peak_x))]

            if fitting_model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
                peak_model = lmfit.models.VoigtModel()
                sigma = float(window.peak_params_grid.GetCellValue(row, 7)) / 2.355
                gamma = float(window.peak_params_grid.GetCellValue(row, 8)) / 2
                amplitude = peak_y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0)
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma, gamma=gamma)
            elif fitting_model in ["Voigt (Area, L/G, \u03c3, S)"]:
                peak_model = lmfit.models.SkewedVoigtModel()
                skew = float(window.peak_params_grid.GetCellValue(row, 9))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7)) / 2.355
                gamma = float(window.peak_params_grid.GetCellValue(row, 8)) / 2
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma, gamma=gamma, skew=skew)
            elif fitting_model == "DS (A, \u03c3, \u03b3)":
                peak_model = lmfit.models.DoniachModel()
                height = float(window.peak_params_grid.GetCellValue(row, 3))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                skew = float(window.peak_params_grid.GetCellValue(row, 9))
                amplitude = PeakFunctions.doniach_sunjic_height_to_amplitude(height, sigma, gamma, skew)
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma, gamma=gamma,
                                                asymmetry=skew)
            elif fitting_model == "DS*G (A, \u03c3, \u03b3, S)":
                peak_model = lmfit.Model(PeakFunctions.DS_G)
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                skew = float(window.peak_params_grid.GetCellValue(row, 9))
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, gamma=gamma,
                                                skew=skew, sigma=sigma)
            elif fitting_model == "ExpGauss.(Area, \u03c3, \u03b3)":
                peak_model = lmfit.models.ExponentialGaussianModel()
                area = float(window.peak_params_grid.GetCellValue(row, 6))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                amplitude = area  # Use area directly as amplitude for area-based model
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma, gamma=gamma)
            elif fitting_model == "Pseudo-Voigt (Area)":
                sigma = fwhm / 2
                peak_model = lmfit.models.PseudoVoigtModel()
                amplitude = peak_y / peak_model.eval(center=0, amplitude=1, sigma=sigma, fraction=lg_ratio / 100, x=0)
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma,
                                                fraction=lg_ratio / 100)
            elif fitting_model in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
                peak_model = lmfit.Model(PeakFunctions.LA)
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                # area = float(window.peak_params_grid.GetCellValue(row, 6))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                params = peak_model.make_params(center=peak_x,amplitude=amplitude,fwhm=fwhm,sigma=sigma,gamma=gamma)
            elif fitting_model in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                peak_model = lmfit.Model(PeakFunctions.LAxG)
                area = float(window.peak_params_grid.GetCellValue(row, 6))
                sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                fwhm_g = float(window.peak_params_grid.GetCellValue(row, 9))
                params = peak_model.make_params(center=peak_x,amplitude=area,fwhm=fwhm,sigma=sigma,gamma=gamma,
                                                fwhm_g=fwhm_g)
            elif fitting_model == "GL (Height)":
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz)
                params = peak_model.make_params(center=peak_x, fwhm=fwhm, fraction=lg_ratio, amplitude=peak_y)
            elif fitting_model == "SGL (Height)":
                peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz)
                params = peak_model.make_params(center=peak_x, fwhm=fwhm, fraction=lg_ratio, amplitude=peak_y)
            elif fitting_model == "GL (Area)":
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz_Area)
                area = float(window.peak_params_grid.GetCellValue(row, 6))  # Assuming area is in column 6
                params = peak_model.make_params(center=peak_x, fwhm=fwhm, fraction=lg_ratio, area=area)
            elif fitting_model == "SGL (Area)":
                peak_model = lmfit.Model(PeakFunctions.S_gauss_lorentz_Area)
                area = float(window.peak_params_grid.GetCellValue(row, 6))  # Assuming area is in column 6
                params = peak_model.make_params(center=peak_x, fwhm=fwhm, fraction=lg_ratio, area=area)
            elif fitting_model == "D-parameter":
                # Skip D-parameter in overall fit calculation
                continue
            elif fitting_model == "Fermi":
                continue
            elif fitting_model == "VBM":
                continue
            elif fitting_model == "Cut-Off":
                # Skip Cut-Off in overall fit calculation
                continue
            elif fitting_model == "SurveyID":
                # Skip D-parameter in overall fit calculation
                continue
            elif fitting_model == "SingleEntity":
                # Handle SingleEntity envelope - get shift and scale from grid
                position_shift = float(window.peak_params_grid.GetCellValue(row, 7))  # Sigma = shift
                area_scale = float(window.peak_params_grid.GetCellValue(row, 8))  # Gamma = scale

                # Find the peak data by letter/index
                sheet_name = window.sheet_combobox.GetValue()
                peak_letter = window.peak_params_grid.GetCellValue(row, 1)
                peaks_dict = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']

                peak_data = None
                for peak_name, data in peaks_dict.items():
                    if data.get('Fitting Model') == 'SingleEntity':
                        if peak_letter in peak_name:
                            peak_data = data
                            break

                if peak_data and 'x_data' in peak_data and 'y_data' in peak_data:
                    # Get stored envelope data
                    x_env = np.array(peak_data['x_data'])
                    y_env = np.array(peak_data['y_data'])

                    # Create interpolator
                    from scipy.interpolate import interp1d
                    interpolator = interp1d(x_env, y_env, kind='cubic',
                                            bounds_error=False, fill_value=0.0)

                    # Shift x and interpolate
                    x_shifted = window.x_values - position_shift
                    y_interpolated = interpolator(x_shifted)

                    # Scale by area
                    peak_fit = y_interpolated * area_scale
                    overall_fit += peak_fit
                    continue
                else:
                    print(f"Warning: SingleEntity data not found for peak {i + 1}")
                    continue
            else:
                print(f"Warning: Unknown fitting model '{fitting_model}' for peak {i + 1}. Skipping this peak.")
                continue
            peak_fit = peak_model.eval(params, x=window.x_values)
            overall_fit += peak_fit

        # Calculate residuals
        # Before the subtraction, ensure arrays have same length
        min_length = min(len(y_values), len(overall_fit))
        y_values = y_values[:min_length]
        overall_fit = overall_fit[:min_length]

        residuals = y_values - overall_fit


        # Determine the scaling factor
        max_raw_data = max(window.y_values) - min(window.y_values)
        desired_max_residual = 0.05 * max_raw_data
        actual_max_residual = max(abs(residuals))
        scaling_factor = desired_max_residual / actual_max_residual if actual_max_residual != 0 else 1

        # Scale residuals
        scaled_residuals = residuals * scaling_factor

        # # Create a masked array where 0 values are masked
        # masked_residuals = ma.masked_where(np.isclose(scaled_residuals, 0, atol=5e-1), scaled_residuals)
        # masked_residuals2 = ma.masked_where(np.isclose(scaled_residuals, 0, atol=5e-1), residuals)

        # # Calculate threshold as percentage of maximum intensity
        # threshold = 0.0005 * np.max(np.abs(scaled_residuals))  # 5% of max residual
        #
        # # Create a masked array using threshold instead of zero
        # masked_residuals = ma.masked_where(np.abs(scaled_residuals) < threshold, scaled_residuals)
        # masked_residuals2 = ma.masked_where(np.abs(scaled_residuals) < threshold, residuals)

        # No masking - use all residual values
        masked_residuals = scaled_residuals
        masked_residuals2 = residuals

        # Remove old overall fit and residuals, keep background lines
        for line in self.ax.lines:
            if line.get_label() in ['Overall Fit', 'Residuals']:
                line.remove()


        # Get background regions for envelope masking too
        bg_regions = self.get_background_regions(window)

        # Plot the overall fit with region masking
        try:
            good_indices = ~np.isnan(overall_fit)
            x_plot = x_values[good_indices]
            y_plot = overall_fit[good_indices]

            # Apply background region masking to envelope
            if bg_regions is not None and len(bg_regions) > 0:
                region_mask = np.zeros(len(x_plot), dtype=bool)
                for bg_start, bg_end in bg_regions:
                    region_mask |= (x_plot >= bg_start) & (x_plot <= bg_end)

                # Set non-region points to NaN instead of removing them
                y_plot_masked = y_plot.copy()
                y_plot_masked = y_plot_masked.astype(float)
                y_plot_masked[~region_mask] = np.nan

                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons - x_plot, y_plot_masked, color=self.envelope_color,
                                 linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                                 linewidth=self.envelope_thickness,
                                 label='D-parameter' if fitting_model == "D-parameter" else
                                 'Fermi' if fitting_model == "Fermi" else
                                 'VBM' if fitting_model == "VBM" else
                                 'Cut-Off' if fitting_model == "Cut-Off" else
                                 'Overall Fit')
                else:
                    self.ax.plot(x_plot, y_plot_masked, color=self.envelope_color,
                                 linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                                 linewidth=self.envelope_thickness,
                                 label='D-parameter' if fitting_model == "D-parameter" else
                                 'Fermi' if fitting_model == "Fermi" else
                                 'VBM' if fitting_model == "VBM" else
                                 'Cut-Off' if fitting_model == "Cut-Off" else
                                 'Overall Fit')
            else:
                # Plot full envelope if no masking
                if window.energy_scale == 'KE':
                    self.ax.plot(window.photons - x_plot, y_plot, color=self.envelope_color,
                                 linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                                 linewidth=self.envelope_thickness,
                                 label='D-parameter' if fitting_model == "D-parameter" else
                                 'Fermi' if fitting_model == "Fermi" else
                                 'VBM' if fitting_model == "VBM" else
                                 'Cut-Off' if fitting_model == "Cut-Off" else
                                 'Overall Fit')
                else:
                    self.ax.plot(x_plot, y_plot, color=self.envelope_color,
                                 linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                                 linewidth=self.envelope_thickness,
                                 label='D-parameter' if fitting_model == "D-parameter" else
                                 'Fermi' if fitting_model == "Fermi" else
                                 'VBM' if fitting_model == "VBM" else
                                 'Cut-Off' if fitting_model == "Cut-Off" else
                                 'Overall Fit')
        except:
            return


        # Handle residuals based on state
        if hasattr(self, 'residuals_state'):
            if self.residuals_state == 1:  # On main plot
                residual_height = 1.07 * max(window.y_values)
                residual_base = self.ax.axhline(y=residual_height, color='grey', linestyle='-.', alpha=0.1)

                # Apply background region masking to residuals
                bg_regions = self.get_background_regions(window)

                if bg_regions is not None and len(bg_regions) > 0:
                    # Create mask for residuals
                    region_mask = np.zeros(len(window.x_values), dtype=bool)
                    for bg_start, bg_end in bg_regions:
                        region_mask |= (window.x_values >= bg_start) & (window.x_values <= bg_end)

                    # Only mask if there are some data points in the regions
                    if np.any(region_mask):
                        masked_residuals_region = masked_residuals.copy()
                        masked_residuals_region[~region_mask] = np.nan
                    else:
                        # No data points in background regions - skip residuals
                        masked_residuals_region = np.full_like(masked_residuals, np.nan)
                else:
                    # No masking - use original masked residuals
                    masked_residuals_region = masked_residuals

                if window.energy_scale == 'KE':
                    residual_line = self.ax.plot(window.photons - window.x_values, masked_residuals_region + residual_height,
                                                 color=self.residual_color, linestyle=self.residual_linestyle,
                                                 alpha=self.residual_alpha, label='Residuals',
                                                 linewidth=self.residual_thickness)
                else:
                    residual_line = self.ax.plot(window.x_values, masked_residuals_region + residual_height,
                                                 color=self.residual_color, linestyle=self.residual_linestyle,
                                                 alpha=self.residual_alpha, label='Residuals',
                                                 linewidth=self.residual_thickness)

                residual_line[0].set_visible(True)
                residual_base.set_visible(True)
                self.ax.get_xaxis().set_visible(True)
            elif self.residuals_state == 2:  # Separate subplot
                # Apply same masking to subplot residuals
                bg_regions = self.get_background_regions(window)

                if bg_regions is not None and len(bg_regions) > 0:
                    region_mask = np.zeros(len(x_values), dtype=bool)
                    for bg_start, bg_end in bg_regions:
                        region_mask |= (x_values >= bg_start) & (x_values <= bg_end)

                    # Only proceed if there are valid data points
                    if np.any(region_mask):
                        masked_residuals_subplot = masked_residuals.copy()
                        masked_residuals_subplot[~region_mask] = np.nan
                        self.setup_residual_subplot(window, x_values, masked_residuals_subplot, self.residual_thickness,
                                                    scaling_factor=1.0)
                    # If no valid points, skip creating the subplot
                else:
                    masked_residuals_subplot = masked_residuals
                    self.setup_residual_subplot(window, x_values, masked_residuals_subplot, self.residual_thickness,
                                                scaling_factor=1.0)


            else:
                self.ax.get_xaxis().set_visible(True)

        # Handle RSD text - shows all three metrics for comparison
        rsd_result = PeakFunctions.calculate_rsd(window.y_values, overall_fit)
        if rsd_result is not None:
            old_rsd, norm_chi, rsd_pct = rsd_result
            rsd_label = f'RSD: {old_rsd:.2f}, χ: {norm_chi:.2f}' #  RSD%: {rsd_pct:.2f}'

            if self.residuals_state == 1:  # For main plot
                self.ax.get_xaxis().set_visible(True)
                y_max = self.ax.get_ylim()[1]
                residual_height = 1.07 * max(window.y_values)
                if residual_height <= y_max:
                    if window.energy_scale == 'KE':
                        x_min = self.ax.get_xlim()[1] - 0.4
                    else:
                        x_min = self.ax.get_xlim()[1] + 0.4
                    if self.rsd_text:
                        try:
                            self.rsd_text.remove()
                        except:
                            print("Error: RSD text removal failed.")
                            pass
                    self.rsd_text = self.ax.text(x_min, residual_height,
                                                 rsd_label,
                                                 horizontalalignment='right',
                                                 verticalalignment='center',
                                                 fontsize=9,
                                                 color=self.residual_color,
                                                 alpha=self.residual_alpha + 0.2,
                                                 bbox=dict(facecolor='white', edgecolor='none'))
            elif self.residuals_state == 2:  # For subplot
                if self.residuals_subplot:
                    if window.energy_scale == 'KE':
                        x_min = self.residuals_subplot.get_xlim()[1] - 0.4
                    else:
                        x_min = self.residuals_subplot.get_xlim()[1] + 0.2
                    y_pos = np.mean(self.residuals_subplot.get_ylim())
                    if self.rsd_text:
                        try:
                            self.rsd_text.remove()
                        except:
                            print("Error: RSD text removal failed.")
                            pass
                    self.rsd_text = self.residuals_subplot.text(x_min, y_pos,
                                                                rsd_label,
                                                                horizontalalignment='right',
                                                                verticalalignment='center',
                                                                fontsize=9,
                                                                bbox=dict(facecolor='none', edgecolor='none'))

        # Only update main plot ylabel if residuals are not in subplot
        if self.residuals_state != 2:
            self.ax.set_ylabel(f'Intensity (CPS), residual x {scaling_factor:.2f}')
        else:
            self.ax.set_ylabel('Intensity (CPS)')

        self.canvas.draw_idle()
        return residuals


    def setup_residual_subplot(self, window, x_values, masked_residuals, residual_thickness=1, scaling_factor=1):
        # Create gridspec at start
        # gs = self.figure.add_gridspec(20, 1, hspace=0.0)
        gs = self.figure.add_gridspec(20, 1, hspace=0.0, top=0.95, bottom=0.1, left=0.1, right=0.95)

        if not self.residuals_subplot:
            self.ax.set_position(gs[0:18, 0].get_position(self.figure))
            self.residuals_subplot = self.figure.add_subplot(gs[18:, 0])

        self.residuals_subplot.clear()

        # Check if we have any valid (non-NaN) data points
        scaled_residuals = masked_residuals * scaling_factor
        if np.all(np.isnan(scaled_residuals)):
            # All values are NaN - set a default range and return early
            self.residuals_subplot.set_ylim(-1, 1)
            self.residuals_subplot.set_ylabel('Res.')
            return

        # Determine x values based on energy scale
        x_plot = window.photons - x_values if window.energy_scale == 'KE' else x_values

        # Plot residuals
        self.residuals_subplot.plot(x_plot, scaled_residuals,
                                    color=self.residual_color,
                                    linestyle=self.residual_linestyle,
                                    alpha=self.residual_alpha,
                                    linewidth=residual_thickness)

        # Configure main plot
        self.ax.get_xaxis().set_visible(False)

        sheet_name = window.sheet_combobox.GetValue()
        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper()
        is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')

        # Configure subplot
        self.residuals_subplot.set_ylabel('Res.')
        if is_raman:
            self.residuals_subplot.set_xlabel('Wavenumber (cm$^{-1}$)')
        elif is_xas:
            self.residuals_subplot.set_xlabel('Photon Energy (eV)')
        elif window.energy_scale == 'KE':
            self.residuals_subplot.set_xlabel('Kinetic Energy (eV)')
        else:
            self.residuals_subplot.set_xlabel('Binding Energy (eV)')
        # x_label = "Kinetic Energy (eV)" if window.energy_scale == 'KE' else "Binding Energy (eV)"
        # self.residuals_subplot.set_xlabel(x_label)
        # self.residuals_subplot.set_xlabel('Binding Energy (eV)')
        self.residuals_subplot.tick_params(axis='x', bottom=True, labelbottom=True,
                                           labelsize=window.axis_number_size, pad=8)
        self.residuals_subplot.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.residuals_subplot.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        # Set font sizes
        self.residuals_subplot.xaxis.label.set_size(window.axis_title_size)
        self.residuals_subplot.yaxis.label.set_size(window.axis_title_size)

        # Set y limits with margin - only use valid (non-NaN) values
        valid_residuals = masked_residuals[~np.isnan(masked_residuals)]
        if len(valid_residuals) > 0:
            y_min, y_max = np.min(valid_residuals), np.max(valid_residuals)
            margin = 0.1 * (y_max - y_min) if y_max != y_min else 0.1
            self.residuals_subplot.set_ylim(y_min - margin, y_max + margin)
        else:
            # Fallback if somehow we still have no valid residuals
            self.residuals_subplot.set_ylim(-1, 1)

        # Get current main plot limits
        main_xlim = self.ax.get_xlim()

        self.residuals_subplot.set_xlim(main_xlim[0], main_xlim[1])

        # Set subplot to share x axis
        self.residuals_subplot.sharex(self.ax)

        # Final styling
        self.residuals_subplot.tick_params(axis='both', labelsize=window.axis_number_size)
        self.residuals_subplot.grid(True, alpha=0.8)
        self.residuals_subplot.set_position(gs[18:, 0].get_position(self.figure))
        self.residuals_subplot.set_visible(True)
        self.residuals_subplot.yaxis.set_visible(self.y_axis_visible)

    def update_peak_fwhm(self, window, x):
        if window.initial_fwhm is not None and window.initial_x is not None:
            row = window.selected_peak_index * 2
            peak_label = window.peak_params_grid.GetCellValue(row, 1)

            delta_x = x - window.initial_x
            new_fwhm = max(window.initial_fwhm + delta_x * 1, 0.3)  # Ensure minimum FWHM of 0.3 eV

            window.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")

            # Update FWHM in window.Data
            sheet_name = window.sheet_combobox.GetValue()
            if sheet_name in window.Data['Core levels'] and 'Fitting' in window.Data['Core levels'][
                sheet_name] and 'Peaks' in window.Data['Core levels'][sheet_name]['Fitting']:
                peaks = window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                if peak_label in peaks:
                    peaks[peak_label]['FWHM'] = new_fwhm

            # Clear and replot
            window.clear_and_replot()

            # Add the cross back
            # window.plot_manager.add_cross_to_peak(window, window.selected_peak_index)

            # Add this line to update the displayed FWHM
            self.peak_manipulation.highlight_selected_peak()

            # Redraw the canvas
            window.canvas.draw_idle()

    def try_float(self, value, default=0.0):
        try:
            return float(value)
        except (ValueError, TypeError):
            return default

    def add_cross_to_peak(self, window, index, skip_fwhm_calc=False):
        try:
            row = index * 2
            peak_x = float(window.peak_params_grid.GetCellValue(row, 2))
            peak_y = float(window.peak_params_grid.GetCellValue(row, 3))

            # Only get these values if we're not skipping FWHM calculation or they're needed
            area = float(window.peak_params_grid.GetCellValue(row, 6))
            model = window.peak_params_grid.GetCellValue(row, 13)

            # Find background value at peak position using vectorized operations
            closest_index = np.argmin(np.abs(window.x_values - peak_x))
            bkg_y = window.background[closest_index]
            peak_y_with_bg = peak_y + bkg_y

            # Handle FWHM determination - use cached value during dragging
            fwhm = window.actual_fwhms.get(index)
            if not skip_fwhm_calc and fwhm is None:
                # Only do expensive calculation when needed and not cached
                grid_fwhm = float(window.peak_params_grid.GetCellValue(row, 4))
                lg_ratio = float(window.peak_params_grid.GetCellValue(row, 5))
                sigma = self.try_float(window.peak_params_grid.GetCellValue(row, 7), 0.0)
                gamma = self.try_float(window.peak_params_grid.GetCellValue(row, 8), 0.0)
                skew = self.try_float(window.peak_params_grid.GetCellValue(row, 9), 0.0)

                from libraries.Peak_Functions import PeakFunctions
                fwhm = PeakFunctions.calculate_actual_fwhm(
                    window.x_values, peak_x, peak_y, grid_fwhm, lg_ratio, area, sigma, gamma, skew, model
                )
                window.actual_fwhms[index] = fwhm

            # Clean up previous elements
            for attr in ['cross', 'peak_info_t', 'peak_letter_t']:
                if hasattr(self, attr) and getattr(self, attr):
                    getattr(self, attr).remove()

            # Always add the cross marker
            self.cross, = self.ax.plot(peak_x, peak_y_with_bg, 'bx', markersize=15,
                                       markerfacecolor='none', picker=5, linewidth=3)
            self.peak_letter = chr(65 + index)

            # Add letter label
            max_y = window.ax.get_ylim()[1]
            y_offset = max_y * 0.02
            self.peak_letter_t = self.ax.text(peak_x, peak_y_with_bg + y_offset,
                                              self.peak_letter, ha='center',
                                              va='bottom', fontsize=12)

            # Only add detailed info if not dragging
            if not skip_fwhm_calc and fwhm is not None:
                self.peak_info = (f'Model: {model}\n'
                                  f'Position: {peak_x} eV\n'
                                  f'FWHM meas.: {fwhm:.3f} eV\n'
                                  f'Area: {area} CPS')
                self.peak_info_t = self.ax.text(peak_x - (fwhm / 2 if fwhm else 0.5),
                                                peak_y_with_bg + y_offset, self.peak_info,
                                                ha='left', va='top', fontsize=8, color='grey')
            else:
                self.peak_info_t = None

            # Update event connections
            self.canvas.mpl_disconnect('motion_notify_event')
            self.canvas.mpl_disconnect('button_release_event')
            self.motion_notify_id = self.canvas.mpl_connect('motion_notify_event', self.peak_manipulation.on_cross_drag)
            self.button_release_id = self.canvas.mpl_connect('button_release_event', self.peak_manipulation.on_cross_release)

            # Only redraw if not dragging to avoid frequent redraws
            if not skip_fwhm_calc:
                self.canvas.draw_idle()

        except Exception as e:
            print(f"Error adding cross to peak: {e}")

    def toggle_residuals(self, window):
        if not hasattr(self, 'residuals_state'):
            self.residuals_state = 0

        self.residuals_state = (self.residuals_state + 1) % 3

        # Remove existing residuals and baselines from main plot
        for line in self.ax.lines:
            if line.get_label() == 'Residuals':
                line.remove()
        if hasattr(self, 'residual_base'):
            self.residual_base.remove()

        # Remove residuals subplot if it exists
        if hasattr(self, 'residuals_subplot'):
            if self.residuals_subplot:
                self.figure.delaxes(self.residuals_subplot)
                self.residuals_subplot = None

        if self.rsd_text:
            self.rsd_text.set_visible(self.residuals_state > 0)

        # Force a full replot which will handle residuals in new state
        window.clear_and_replot()

        self.canvas.draw_idle()

    def toggle_fitting_results(self):
        if self.fitting_results_text:
            self.fitting_results_visible = not self.fitting_results_visible
            self.fitting_results_text.set_visible(self.fitting_results_visible)
            self.canvas.draw_idle()

    def set_fitting_results_text(self, text):
        # Method to set or update the fitting results text
        if self.fitting_results_text:
            self.fitting_results_text.remove()
        self.fitting_results_text = self.ax.text(
            0.02, 0.04, text,
            transform=self.ax.transAxes,
            fontsize=9,
            verticalalignment='bottom',
            horizontalalignment='left',
            bbox=dict(facecolor='white', edgecolor='grey', alpha=0.7),
        )
        self.fitting_results_text.set_visible(self.fitting_results_visible)


    def toggle_legend(self):
        self.legend_visible = (self.legend_visible + 1) % 3

        # Check if any peak has Fermi fitting model
        handles, labels = self.ax.get_legend_handles_labels()
        has_fermi_peak = False

        # Quick check by looking for Fermi-specific labels
        for label in labels:
            if 'Center:' in label or '16%-84%:' in label:
                has_fermi_peak = True
                break

        # If Fermi peak exists, always show full natural legend regardless of mode
        if has_fermi_peak:
            legend = self.ax.get_legend()
            if legend:
                legend.set_visible(True)  # Always visible when Fermi present
            return  # Exit early, ignore legend modes

        # ORIGINAL TOGGLE LOGIC (only when no Fermi)
        legend = self.ax.get_legend()
        if legend:
            if self.legend_visible == 0:
                legend.set_visible(False)
            elif self.legend_visible == 1:
                legend.set_visible(True)
            else:
                handles, labels = self.ax.get_legend_handles_labels()
                filtered_handles = []
                filtered_labels = []
                for h, l in zip(handles, labels):
                    if l not in ["Raw Data", "Background", "Overall Fit"]:
                        clean_label = re.sub(r'\$.*?\$', '', l)
                        split_label = clean_label.split()
                        if len(split_label) > 1 and split_label[1].strip():
                            filtered_handles.append(h)
                            filtered_labels.append(l)
                self.ax.legend(filtered_handles, filtered_labels, loc='upper left', frameon=True, fancybox=True,
                               framealpha=0.1, edgecolor='gray')
        self.canvas.draw_idle()


    def update_legend(self, window):
        sheet_name = window.sheet_combobox.GetValue()

        # Skip legend update for zzProfile sheets - they handle their own legend
        if sheet_name.startswith('zzProfile'):
            return

        handles, labels = self.ax.get_legend_handles_labels()

        # Check if any peak has Fermi, VBM, or Cut-Off fitting model
        has_fermi_peak = False
        has_vbm_peak = False
        has_cutoff_peak = False
        num_peaks = window.peak_params_grid.GetNumberRows() // 2
        for i in range(num_peaks):
            fitting_model = window.peak_params_grid.GetCellValue(i * 2, 13)  # Column 13 is fitting model
            if fitting_model == "Fermi":
                has_fermi_peak = True
            elif fitting_model == "VBM":
                has_vbm_peak = True
            elif fitting_model == "Cut-Off":
                has_cutoff_peak = True

        # If Fermi, VBM, or Cut-Off peak exists, use natural legend from plot_peak (bypass all modes)
        if has_fermi_peak or has_vbm_peak or has_cutoff_peak:
            if handles and labels:
                self.ax.legend(handles, labels, loc='upper left', frameon=True, fancybox=True,
                               framealpha=0.1, edgecolor='gray')
            return  # Exit early, ignore legend modes

        # ORIGINAL LEGEND MODE LOGIC (only when no Fermi or VBM)
        # Extract current core level name for compact legend
        current_core_level = self.extract_core_level_name(sheet_name)

        peak_labels = []
        filtered_peak_labels = []
        compact_peak_labels = []  # Separate list for compact labels

        for i in range(num_peaks):
            label = window.peak_params_grid.GetCellValue(i * 2, 1)
            formatted_label = re.sub(r'(\d+/\d+)', r'$_{\1}$', label)

            # Create compact version for peaks-only mode
            compact_label = self.make_compact_legend_label(formatted_label, current_core_level)

            clean_label = re.sub(r'\$.*?\$', '', formatted_label)
            split_label = clean_label.split()
            if len(split_label) > 1 and split_label[1].strip():
                peak_labels.append(label)
                filtered_peak_labels.append(formatted_label)  # Full name for full legend
                compact_peak_labels.append(compact_label)  # Compact name for peaks-only

        if self.legend_visible == 2:
            ordered_handles = []
            for l in peak_labels:
                for index, label in enumerate(labels):
                    if label == l:
                        ordered_handles.append(handles[index])
                        break
            if ordered_handles and compact_peak_labels:
                self.ax.legend(ordered_handles, compact_peak_labels, loc='upper left', frameon=True, fancybox=True,
                               framealpha=0.1, edgecolor='gray')
            else:
                self.ax.legend().set_visible(False)
        else:
            has_overall_fit = "Overall Fit" in labels
            has_raw_data = "Raw Data" in labels

            legend_order = []
            legend_order2 = []

            if has_raw_data:
                legend_order.append("Raw Data")
                legend_order2.append("Raw Data")
            legend_order.append("Background")
            legend_order2.append("Background")
            if has_overall_fit:
                legend_order.append("Overall Fit")
                legend_order2.append("Overall Fit")

            legend_order += peak_labels
            legend_order2 += filtered_peak_labels  # Use full names for full legend

            if legend_order and self.legend_visible:
                ordered_handles = []
                for l in legend_order:
                    for index, label in enumerate(labels):
                        if label == l:
                            ordered_handles.append(handles[index])
                            break
                self.ax.legend(ordered_handles, legend_order2, loc='upper left', frameon=True, fancybox=True,
                               framealpha=0.1, edgecolor='gray')
            else:
                self.ax.legend().remove()
                self.ax.legend().set_visible(False)

        self.canvas.draw_idle()

    def extract_core_level_name(self, sheet_name):
        """Extract core level name from sheet name (e.g., 'Sr3d1' -> 'Sr3d', 'C1s1' -> 'C1s')"""
        import re
        # Match pattern like C1s, N1s, Sr3d, etc.
        match = re.match(r'([A-Z][a-z]?\d+[spdf])', sheet_name)
        if match:
            return match.group(1)
        return None

    def make_compact_legend_label(self, original_label, current_core_level):
        """
        Make legend labels compact by removing core level prefix if it matches current core level.
        Handles doublets for p, d, f orbitals.
        Examples:
        - If on C1s: "C1s C-C" becomes "C-C"
        - If on Sr3d: "Sr3d5/2 SrO" becomes "SrO"
        - If on Sr3d: "Sr3d$_{5/2}$ SrO" becomes "SrO"
        """
        import re

        # Skip processing if no current core level
        if not current_core_level:
            return original_label

        # For s orbitals (like C1s), check for exact match
        if current_core_level.endswith('s'):
            words = original_label.split()
            if len(words) >= 2 and words[0] == current_core_level:
                return ' '.join(words[1:])

        # For p, d, f orbitals that can have doublets
        elif current_core_level[-1] in ['p', 'd', 'f']:
            # Pattern to match doublet notation: Sr3d5/2, Sr3d3/2, Sr3d$_{5/2}$, Sr3d$_{3/2}$, etc.
            doublet_pattern = rf'^{re.escape(current_core_level)}(?:\d+/\d+|\$_{{[\d/]+}}\$)\s+(.+)$'
            match = re.match(doublet_pattern, original_label)
            if match:
                return match.group(1)  # Return everything after the doublet part

            # Also check for simple core level match (fallback)
            words = original_label.split()
            if len(words) >= 2 and words[0] == current_core_level:
                return ' '.join(words[1:])

        return original_label


    @staticmethod
    def format_sheet_name(sheet_name):
        import re
        # First check if it's a survey or wide scan
        survey_match = re.match(r'(Survey|SURVEY|survey|Wide|WIDE|wide)(\d+)', sheet_name)
        if survey_match:
            scan_type = survey_match.group(1)
            # Capitalize first letter, lowercase the rest
            return scan_type.capitalize()

        # Original code for element detection
        match = re.match(r'([A-Z][a-z]*)(\d+[spdfg])', sheet_name)
        if match:
            element, shell = match.groups()
            return f"{element} {shell}"
        else:
            return sheet_name

    def update_plots_be_correction(self, window, delta_correction):
        sheet_name = window.sheet_combobox.GetValue()
        limits = window.plot_config.get_plot_limits(window, sheet_name)

        # Update x-axis limits
        limits['Xmin'] += delta_correction
        limits['Xmax'] += delta_correction

        # Update plot
        window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
        window.canvas.draw_idle()

    def update_plot_style(self, style, scatter_size, line_width, line_alpha, scatter_color, line_color, scatter_marker,
                          background_color, background_alpha, background_linestyle,
                          envelope_color, envelope_alpha, envelope_linestyle,
                          residual_color, residual_alpha, residual_linestyle,
                          raw_data_linestyle, peak_colors, peak_alpha,
                          background_thickness, envelope_thickness, residual_thickness
    ):
        self.plot_style = style
        self.scatter_size = scatter_size
        self.line_width = line_width
        self.line_alpha = line_alpha
        self.scatter_color = scatter_color
        self.line_color = line_color
        self.scatter_marker = scatter_marker
        self.background_color = background_color
        self.background_alpha = background_alpha
        self.background_linestyle = background_linestyle
        self.envelope_color = envelope_color
        self.envelope_alpha = envelope_alpha
        self.envelope_linestyle = envelope_linestyle
        self.residual_color = residual_color
        self.residual_alpha = residual_alpha
        self.residual_linestyle = residual_linestyle
        self.raw_data_linestyle = raw_data_linestyle
        self.peak_colors = peak_colors
        self.peak_alpha = peak_alpha
        self.background_thickness = background_thickness
        self.envelope_thickness = envelope_thickness
        self.residual_thickness = residual_thickness

    def toggle_peak_fill(self):
        self.peak_fill_enabled = not self.peak_fill_enabled
        return self.peak_fill_enabled  # Return the new state

    def plot_background(self, window, use_smoothing=False):
        """
        Calculate and plot the background for the selected sheet.

        This method computes the background using various methods (Multi-Regions Smart, Shirley, Linear, Smart)
        based on the user's selection. It updates the background data in the window.Data structure
        and plots the result on the main graph.

        Args:
            window: The main application window containing all necessary data and UI elements.

        Raises:
            ValueError: If an unknown background method is selected.
        """
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            # Extract x and y values for the current sheet
            x_values = np.array(window.Data['Core levels'][sheet_name]['B.E.'], dtype=float)
            y_values = np.array(window.Data['Core levels'][sheet_name]['Raw Data'], dtype=float)

            # Apply smoothing if requested
            if use_smoothing:
                from scipy.ndimage import gaussian_filter1d
                y_values_for_calculation = gaussian_filter1d(y_values, sigma=5)
            else:
                y_values_for_calculation = y_values

            # Remove any existing background lines from the plot
            lines_to_remove = [line for line in self.ax.lines if line.get_label().startswith("Background")]
            for line in lines_to_remove:
                line.remove()

            # Initialize or retrieve the background data
            if 'Bkg Y' not in window.Data['Core levels'][sheet_name]['Background'] or not \
                    window.Data['Core levels'][sheet_name]['Background']['Bkg Y']:
                window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values.tolist()

            try:
                method = window.background_method
            except (AttributeError, ValueError):
                method = "Multi-Regions Smart"

            try:
                offset_h = float(window.offset_h)
            except (AttributeError, ValueError):
                offset_h = 0

            try:
                offset_l = float(window.offset_l)
            except (AttributeError, ValueError):
                offset_l = 0

            # Calculate background based on the selected method
            if method == "Multi-Regions Smart":
                background_filtered, label = self._calculate_adaptive_smart_background(window, x_values,
                                                                                       y_values_for_calculation,
                                                                                       offset_h, offset_l)
            else:
                background_filtered, label = self._calculate_other_background(window, x_values,
                                                                              y_values_for_calculation, method,
                                                                              offset_h, offset_l)

            # Update the background data in the window.Data structure
            self._update_background_data(window, sheet_name, x_values, background_filtered, method, offset_h, offset_l)
            window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = background_filtered.tolist()
            window.background = background_filtered

            # Plot the calculated background
            self.ax.plot(x_values, window.background, color=self.background_color,
                         linestyle=self.background_linestyle,
                         alpha=self.background_alpha,
                         linewidth=self.background_thickness,
                         label=label)

            # Replot everything if peaks exist
            if window.peak_params_grid.GetNumberRows() > 0:
                window.clear_and_replot()
                self.update_legend(window)

            self.ax.legend(loc='upper left', frameon=True, fancybox=True, framealpha=0.1, edgecolor='gray')
            self.canvas.draw()

        except Exception as e:
            print("Error in plot_background:", str(e))
            # import traceback
            # traceback.print_exc()
            # wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)

    def _calculate_adaptive_smart_background(self, window, x_values, y_values, offset_h, offset_l):
        """Helper method to calculate Multi-Regions Smart background."""
        sheet_name = window.sheet_combobox.GetValue()  # Get the current sheet name

        # Use vline positions instead of min/max of x_values
        if window.vline1 is not None and window.vline2 is not None:
            vline1_x = window.vline1.get_xdata()[0]
            vline2_x = window.vline2.get_xdata()[0]
            bg_min_energy = min(vline1_x, vline2_x)
            bg_max_energy = max(vline1_x, vline2_x)
        else:
            # Fallback to full range only if no vlines
            bg_min_energy, bg_max_energy = min(x_values), max(x_values)

        window.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = bg_min_energy
        window.Data['Core levels'][sheet_name]['Background']['Bkg High'] = bg_max_energy

        # Determine the adaptive range
        if window.vline1 is not None and window.vline2 is not None:
            vline1_x, vline2_x = window.vline1.get_xdata()[0], window.vline2.get_xdata()[0]
            adaptive_range = (min(vline1_x, vline2_x), max(vline1_x, vline2_x))
        else:
            adaptive_range = (bg_min_energy, bg_max_energy)

        current_background = np.array(window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])


        background_filtered = BackgroundCalculations.calculate_adaptive_smart_background(
            x_values, y_values, adaptive_range, current_background, offset_h, offset_l, num_points=averaging_points
        )
        return background_filtered, 'Background'

    def _calculate_other_background(self, window, x_values, y_values, method, offset_h, offset_l):
        """Helper method to calculate background for non-Multi-Regions Smart methods."""
        sheet_name = window.sheet_combobox.GetValue()

        # Get the proper energy range from vlines or from stored values
        if window.vline1 is not None and window.vline2 is not None:
            # Convert display positions back to BE for calculations (x_values are always in BE)
            vline1_be = window.convert_energy_from_display(window.vline1.get_xdata()[0])
            vline2_be = window.convert_energy_from_display(window.vline2.get_xdata()[0])
            bg_min_energy = min(vline1_be, vline2_be)
            bg_max_energy = max(vline1_be, vline2_be)

            # Safety check: ensure the converted range overlaps with actual data
            data_min = min(x_values)
            data_max = max(x_values)

            # Clamp the range to actual data bounds to prevent empty arrays
            bg_min_energy = max(bg_min_energy, data_min)
            bg_max_energy = min(bg_max_energy, data_max)

            # If range is invalid after clamping, use stored values as fallback
            if bg_min_energy >= bg_max_energy:
                bg_min_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg Low')
                bg_max_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg High')
        else:
            bg_min_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg Low')
            bg_max_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg High')

        if bg_min_energy is None or bg_max_energy is None or bg_min_energy >= bg_max_energy:
            # Default to full range if invalid
            bg_min_energy = min(x_values)
            bg_max_energy = max(x_values)

        try:
            bg_min_energy = float(bg_min_energy) if bg_min_energy != '' else min(x_values)
            bg_max_energy = float(bg_max_energy) if bg_max_energy != '' else max(x_values)
        except (ValueError, TypeError):
            # If conversion fails, use default values
            bg_min_energy = min(x_values)
            bg_max_energy = max(x_values)



        # Create mask and ensure we have at least some data points
        mask = (x_values >= bg_min_energy) & (x_values <= bg_max_energy)

        # Safety check: ensure we have at least 2 data points for background calculation
        if np.sum(mask) < 2:
            print(f"Warning: Background range too narrow, using full data range")
            print(f"Range: {bg_min_energy:.2f} - {bg_max_energy:.2f}")
            print(f"Data range: {min(x_values):.2f} - {max(x_values):.2f}")
            mask = np.ones(len(x_values), dtype=bool)  # Use full range

        x_values_filtered = x_values[mask]
        y_values_filtered = y_values[mask]

        # Additional safety check
        if len(x_values_filtered) < 2:
            print("Error: Still no data points after fallback, using minimal range")
            # Use first and last points as minimal range
            mask = np.zeros(len(x_values), dtype=bool)
            mask[0] = True
            mask[-1] = True
            x_values_filtered = x_values[mask]
            y_values_filtered = y_values[mask]

        # Get averaging points from window
        averaging_points = getattr(window, 'averaging_points', 5)

        if method == "Shirley":
            background_filtered = BackgroundCalculations.calculate_shirley_background(x_values_filtered,
                                                                                      y_values_filtered, offset_h,
                                                                                      offset_l, num_points=averaging_points)
            label = 'Background (Shirley)'
        elif method == "Linear":
            background_filtered = BackgroundCalculations.calculate_linear_background(x_values_filtered,
                                                                                     y_values_filtered, offset_h,
                                                                                     offset_l, num_points=averaging_points)
            label = 'Background (Linear)'
        elif method in ["Smart", "Multi-Regions Smart", "Multiple Regions Smart"]:
            background_filtered = BackgroundCalculations.calculate_smart_background(x_values_filtered,
                                                                                    y_values_filtered, offset_h,
                                                                                    offset_l, num_points=averaging_points)
            label = 'Background (Smart)'
        elif method == "Offset":
            background_filtered = BackgroundCalculations.calculate_offset_background(x_values_filtered,
                                                                                     y_values_filtered, offset_h,
                                                                                     offset_l)
            label = 'Background (Offset)'
        elif method == "U4-Tougaard":
            background_filtered = BackgroundCalculations.calculate_tougaard_background(x_values_filtered,
                                                                                       y_values_filtered,
                                                                                       sheet_name,
                                                                                       window)
            label = 'Background (U4-Tougaard)'
        elif method == "U2-Tougaard":
            background_filtered = BackgroundCalculations.calculate_u2_tougaard_background(x_values_filtered,
                                                                                          y_values_filtered,
                                                                                          sheet_name,
                                                                                          window)
            label = 'Background (U2-Tougaard)'
        elif method == "2x U4-Tougaard":
            background_filtered = BackgroundCalculations.calculate_double_tougaard_background(x_values_filtered,
                                                                                              y_values_filtered,
                                                                                              sheet_name,
                                                                                              window)
            label = 'Background (Tougaard)'
        elif method == "3x U4-Tougaard":
            background_filtered = BackgroundCalculations.calculate_triple_tougaard_background(x_values_filtered,
                                                                                              y_values_filtered,
                                                                                              sheet_name,
                                                                                              window)
            label = 'Background (Tougaard)'
        elif method == "Arctan-XAS":
            background_filtered = BackgroundCalculations.calculate_arctan_background(x_values_filtered,
                                                                                     y_values_filtered, offset_h,
                                                                                     offset_l, num_points=averaging_points)
            label = 'Background (Arctan)'
        elif method == "ALS-Raman":
            # Get ALS parameters if available
            lambda_val = 1e5
            p_val = 0.001

            if hasattr(window, 'raman_window'):
                if hasattr(window.raman_window, 'als_lambda'):
                    lambda_val = float(window.raman_window.als_lambda.GetValue())
                if hasattr(window.raman_window, 'als_p'):
                    p_val = float(window.raman_window.als_p.GetValue())

            # Use lmfit-based ALS
            background_filtered = BackgroundCalculations.calculate_als_background_spectral(
                x_values_filtered, y_values_filtered, lambda_val=lambda_val, p=p_val
            )
            label = 'Background (ALS-Raman)'
        elif method == "Active Shirley":
            # Check if we have stored k and base const from previous fitting
            stored_k = window.Data['Core levels'][sheet_name]['Background'].get('Active_Shirley_k', None)
            stored_const_base = window.Data['Core levels'][sheet_name]['Background'].get('Active_Shirley_const_base', None)

            if stored_k is not None and stored_const_base is not None:
                # Use stored base const and apply current offset_l
                new_const = stored_const_base + offset_l
                dx = np.abs(np.mean(np.diff(x_values_filtered))) if len(x_values_filtered) > 1 else 1
                y_above_const = np.maximum(y_values_filtered - new_const, 0)

                # Cumulative integral from low BE toward high BE
                cumulative_integral = np.cumsum(y_above_const[::-1])[::-1] * dx
                cumulative_integral = np.roll(cumulative_integral, -1)
                cumulative_integral[-1] = 0

                background_filtered = new_const + stored_k * cumulative_integral
            else:
                # No stored values, use initial flat background
                background_filtered = BackgroundCalculations.calculate_active_shirley_background(
                    x_values_filtered, y_values_filtered, offset_h, offset_l, num_points=averaging_points)
            label = 'Background (Active Shirley)'
        elif method == "Active Tougaard":
            # Check if we have stored B from previous fitting
            stored_B = window.Data['Core levels'][sheet_name]['Background'].get('Active_Tougaard_B', None)

            if stored_B is not None:
                C = float(window.Data['Core levels'][sheet_name]['Background'].get('Tougaard_C', 1643))

                # Get base baseline at low BE (without offset) and apply current offset_l
                base_baseline = BackgroundCalculations.calculate_endpoint_average(
                    x_values_filtered, y_values_filtered, x_values_filtered[-1], averaging_points)
                baseline = base_baseline + offset_l

                y_shifted = np.maximum(y_values_filtered - baseline, 0)
                dx = np.abs(np.mean(np.diff(x_values_filtered)))
                n = len(x_values_filtered)
                bg = np.zeros(n, dtype=float)

                # Vectorized Tougaard integral
                for i in range(n - 1):
                    T = np.abs(x_values_filtered[i + 1:] - x_values_filtered[i])
                    K = stored_B * T / ((C + T ** 2) ** 2)
                    bg[i] = np.sum(K * y_shifted[i + 1:]) * dx

                background_filtered = bg + baseline
            else:
                # No stored values, use initial flat background
                background_filtered = BackgroundCalculations.calculate_active_tougaard_background(
                    x_values_filtered, y_values_filtered, offset_h, offset_l, num_points=averaging_points)
            label = 'Background (Active Tougaard)'
        else:
            background_filtered = BackgroundCalculations.calculate_smart_background(x_values_filtered,
                                                                                    y_values_filtered, offset_h,
                                                                                    offset_l)
            label = 'Background (Smart)'

        new_background = np.array(window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
        new_background[mask] = background_filtered
        return new_background, label

    def _update_background_data(self, window, sheet_name, x_values, background, method, offset_h, offset_l):
        """Helper method to update the background data in window.Data."""

        # Use vline positions instead of min/max of x_values
        if window.vline1 is not None and window.vline2 is not None:
            vline1_x = window.vline1.get_xdata()[0]
            vline2_x = window.vline2.get_xdata()[0]
            bg_low = round(min(vline1_x, vline2_x),2)
            bg_high = round(max(vline1_x, vline2_x),2)
        else:
            # Fallback to full range only if no vlines
            bg_low = round(min(x_values),2)
            bg_high = round(max(x_values),2)

        window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = background.tolist()
        window.background = background
        window.Data['Core levels'][sheet_name]['Background'].update({
            'Bkg Y': background.tolist(),
            'Bkg Type': method,
            'Bkg Low': bg_low,
            'Bkg High': bg_high,
            'Bkg Offset Low': offset_l,
            'Bkg Offset High': offset_h,
            'Bkg X': x_values.tolist()
        })

    def clear_background(self, window):
        """
        Clear the background and reset related data for the current sheet.

        This method resets the background to the raw data, clears all peak information,
        and resets various plot-related parameters. It's used when the user wants to
        start fresh with background subtraction and peak fitting.

        Args:
            window: The main application window containing all necessary data and UI elements.

        Raises:
            Exception: If any error occurs during the clearing process.
        """
        sheet_name = window.sheet_combobox.GetValue()

        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            # Clear the current plot
            self.ax.clear()

            # Retrieve raw data for the current sheet
            x_values = window.Data['Core levels'][sheet_name]['B.E.']
            y_values = window.Data['Core levels'][sheet_name]['Raw Data']

            # Plot the raw data
            self.ax.scatter(x_values, y_values, facecolors='black', marker='o', s=15, edgecolors='black',
                            label='Raw Data')

            # Update main window's x and y values
            window.x_values = np.array(x_values)
            window.y_values = np.array(y_values)

            # Reset background to raw data
            window.Data['Core levels'][sheet_name]['Background']['Bkg X'] = x_values
            window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values
            window.background = np.array(y_values)

            # Reset background parameters
            window.Data['Core levels'][sheet_name]['Background'].update({
                'Bkg Type': '',
                'Bkg Low': '',
                'Bkg High': '',
                'Bkg Offset Low': '',
                'Bkg Offset High': ''
            })

            # Set plot limits and formatting
            self.ax.set_xlim([max(window.x_values), min(window.x_values)])
            self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
            self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
            self.ax.legend(loc='upper left', frameon=True, fancybox=True, framealpha=0.1, edgecolor='gray')

            # Hide the peak selection cross if it exists
            if hasattr(window, 'cross') and window.cross:
                window.cross.set_visible(False)

            # Reset vertical lines used for background selection
            window.vline1 = window.vline2 = window.vline3 = window.vline4 = None
            window.show_hide_vlines()

            # Clear all peak data from the grid
            num_rows = window.peak_params_grid.GetNumberRows()
            if num_rows > 0:
                window.peak_params_grid.DeleteRows(0, num_rows)

            # Reset peak-related variables
            window.peak_count = 0
            window.selected_peak_index = None

            # Clear fitting data from window.Data
            if 'Fitting' in window.Data['Core levels'][sheet_name]:
                window.Data['Core levels'][sheet_name]['Fitting'] = {}

            window.offset_l = 0
            window.offset_h = 0
            window.fitting_window.offset_l_text.SetValue('0')
            window.fitting_window.offset_h_text.SetValue('0')

            # Redraw the canvas and update layout
            self.canvas.draw_idle()
            window.panel.Layout()

        except Exception as e:
            return
            # wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)


    def clear_background_only(self, window):
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name in window.Data['Core levels']:
            # First, remove any existing background lines from the plot
            for line in window.ax.lines:
                if line.get_label() == 'Background':
                    line.remove()

            # Reset background to raw data
            raw_data = window.Data['Core levels'][sheet_name]['Raw Data']
            window.Data['Core levels'][sheet_name]['Background'] = {
                'Bkg Type': '',
                'Bkg Low': None,
                'Bkg High': None,
                'Bkg Offset Low': 0,
                'Bkg Offset High': 0,
                'Bkg Y': raw_data.copy() if isinstance(raw_data, list) else raw_data
            }

            # Update window.background attribute
            window.background = np.array(raw_data)

            # Reset vlines and background parameters
            window.vline1 = None
            window.vline2 = None
            window.bg_min_energy = None
            window.bg_max_energy = None
            window.show_hide_vlines()

            window.offset_l = 0
            window.offset_h = 0
            window.fitting_window.offset_l_text.SetValue('0')
            window.fitting_window.offset_h_text.SetValue('0')

            # Use plot_data instead of clear_and_replot to start fresh
            self.plot_data(window)

            # Force canvas update
            window.canvas.draw_idle()

    def get_background_regions(self, window):
        """Get background regions from recorded ranges in window.Data"""
        bg_regions = []
        try:
            sheet_name = window.sheet_combobox.GetValue()
            if (sheet_name in window.Data['Core levels'] and
                    'Background' in window.Data['Core levels'][sheet_name] and
                    'Recorded_Ranges' in window.Data['Core levels'][sheet_name]['Background']):

                recorded_ranges = window.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges']

                for range_data in recorded_ranges:
                    if len(range_data) >= 4:  # (offset_h, offset_l, min_range, max_range)
                        offset_h, offset_l, min_range, max_range = range_data[:4]
                        try:
                            min_range = float(min_range)
                            max_range = float(max_range)
                            bg_regions.append((min(min_range, max_range), max(min_range, max_range)))
                        except (ValueError, TypeError):
                            continue
        except:
            pass

        # Return None if no valid regions found (means show everything)
        return bg_regions if bg_regions else None

    # Add this method to PlotManager class in Plot_Operations.py

    def toggle_survey_table(self):
        """Toggle survey table display on/off"""
        if hasattr(self, 'survey_table_text') and self.survey_table_text:
            for text_obj in self.survey_table_text[:]:  # Create a copy of the list
                try:
                    if text_obj in self.ax.texts or hasattr(text_obj, 'remove'):
                        text_obj.remove()
                except (ValueError, AttributeError):
                    # Text object may have already been removed or doesn't exist
                    pass
            self.survey_table_text.clear()

        # If turning on and it's a survey plot, draw the table
        if (hasattr(self, 'survey_table_state') and self.survey_table_state == 1):
            self.draw_survey_table()

        self.canvas.draw_idle()

    def is_survey_plot(self):
        """Check if current sheet is a survey/wide scan"""
        if not hasattr(self.window, 'sheet_combobox'):
            return False

        sheet_name = self.window.sheet_combobox.GetValue().lower()
        return any(x in sheet_name for x in ['survey', 'wide'])

    def draw_survey_table_OLD(self):
        """Draw table with core levels, BE, and atomic concentrations for survey plots only"""
        if not self.is_survey_plot():
            return

        if not hasattr(self, 'survey_table_state') or self.survey_table_state == 0:
            return

        # Hide the plot legend when drawing survey table
        legend = self.ax.get_legend()
        if legend:
            legend.set_visible(False)

        # Clear existing table
        if hasattr(self, 'survey_table_text') and self.survey_table_text:
            for text_obj in self.survey_table_text[:]:  # Create a copy of the list
                try:
                    if text_obj in self.ax.texts or hasattr(text_obj, 'remove'):
                        text_obj.remove()
                except (ValueError, AttributeError):
                    # Text object may have already been removed or doesn't exist
                    pass
            self.survey_table_text.clear()

        self.survey_table_text = []

        # Get data from peak fitting grid
        grid = self.window.peak_params_grid
        num_peaks = grid.GetNumberRows() // 2

        if num_peaks == 0:
            return

        # Collect peak data with color indices
        table_data = []
        for i in range(num_peaks):
            row = i * 2
            try:
                core_level = grid.GetCellValue(row, 1)  # Peak name
                be = float(grid.GetCellValue(row, 2))  # BE position
                atomic_conc = float(grid.GetCellValue(row, 10))  # Atomic concentration

                # Clean up core level name (remove " p1", " p2" etc.)
                import re
                core_level_clean = re.sub(r'\s+p\d+$', '', core_level)

                # Skip if atomic concentration is 0 (not significant)
                if atomic_conc > 0.1:  # Only show peaks with >0.1% concentration
                    table_data.append((core_level_clean, be, atomic_conc, i))  # Include peak index
            except (ValueError, IndexError):
                continue

        if not table_data:
            return

        # Sort by atomic concentration (highest first)
        table_data.sort(key=lambda x: x[2], reverse=True)

        # Get font size from preferences
        font_size = getattr(self.window, 'label_font_size', 12)

        # Table dimensions
        row_height = 0.04
        header_height = 0.045
        col_widths = [0.09, 0.09, 0.09]  # Core Level, BE(eV), At.(%)
        table_width = sum(col_widths)
        table_height = header_height + len(table_data) * row_height

        # Position table in top-left corner to hide legend underneath
        x_start = 0.02
        y_start = 0.98

        # Draw table background - fully opaque
        from matplotlib.patches import Rectangle
        table_bg = Rectangle((x_start, y_start - table_height), table_width, table_height,
                             transform=self.ax.transAxes, facecolor='white',
                             edgecolor='black', linewidth=2, alpha=1.0)
        self.ax.add_patch(table_bg)
        self.survey_table_text.append(table_bg)

        # Draw table header with merged cells look
        headers = ['', 'BE (eV)', 'At. (%)']
        header_y = y_start - header_height / 2

        x_pos = x_start
        for col, (header, width) in enumerate(zip(headers, col_widths)):
            # Draw header cell background - fully opaque
            header_bg = Rectangle((x_pos, y_start - header_height), width, header_height,
                                  transform=self.ax.transAxes, facecolor='lightgray',
                                  edgecolor='gray', linewidth=0.5, alpha=1.0)
            self.ax.add_patch(header_bg)
            self.survey_table_text.append(header_bg)

            # Add header text
            text_obj = self.ax.text(x_pos + width / 2, header_y, header,
                                    transform=self.ax.transAxes,
                                    fontsize=font_size - 1,
                                    fontweight='bold',
                                    ha='center', va='center',
                                    color='black')
            self.survey_table_text.append(text_obj)

            x_pos += width

        # Draw data rows
        for row_idx, (core_level, be, atomic_conc, peak_idx) in enumerate(table_data):
            y_pos = y_start - header_height - (row_idx + 1) * row_height
            row_y = y_pos + row_height / 2

            x_pos = x_start
            row_data = [core_level, f"{be:.2f}", f"{atomic_conc:.2f}"]

            for col, (data, width) in enumerate(zip(row_data, col_widths)):
                # Set cell background color - peak name column gets peak color, others alternating
                if col == 0:  # First column (peak name)
                    if hasattr(self.window, 'peak_colors') and len(self.window.peak_colors) > 0:
                        cell_color = self.window.peak_colors[peak_idx % len(self.window.peak_colors)]
                    else:
                        cell_color = 'white'
                else:  # Other columns use alternating colors
                    cell_color = 'white' if row_idx % 2 == 0 else '#f8f8f8'

                # Draw cell background - fully opaque
                cell_bg = Rectangle((x_pos, y_pos), width, row_height,
                                    transform=self.ax.transAxes, facecolor=cell_color,
                                    edgecolor='gray', linewidth=0.5, alpha=1.0)
                self.ax.add_patch(cell_bg)
                self.survey_table_text.append(cell_bg)

                # Add cell text
                text_obj = self.ax.text(x_pos + width / 2, row_y, data,
                                        transform=self.ax.transAxes,
                                        fontsize=font_size - 1,
                                        ha='center', va='center',
                                        color='black')
                self.survey_table_text.append(text_obj)

                x_pos += width

    def draw_survey_table(self):
        """Draw table with core levels, BE, and atomic concentrations for survey plots only"""
        if not self.is_survey_plot():
            return

        if not hasattr(self, 'survey_table_state') or self.survey_table_state == 0:
            return

        # Hide the plot legend when drawing survey table
        legend = self.ax.get_legend()
        if legend:
            legend.set_visible(False)

        # Clear existing table
        if hasattr(self, 'survey_table_text') and self.survey_table_text:
            for text_obj in self.survey_table_text[:]:  # Create a copy of the list
                try:
                    if text_obj in self.ax.texts or hasattr(text_obj, 'remove'):
                        text_obj.remove()
                except (ValueError, AttributeError):
                    # Text object may have already been removed or doesn't exist
                    pass
            self.survey_table_text.clear()

        self.survey_table_text = []

        # Get data from peak fitting grid
        grid = self.window.peak_params_grid
        num_peaks = grid.GetNumberRows() // 2

        if num_peaks == 0:
            return

        # Collect peak data with color indices
        table_data = []
        for i in range(num_peaks):
            row = i * 2
            try:
                core_level = grid.GetCellValue(row, 1)  # Peak name
                be = float(grid.GetCellValue(row, 2))  # BE position
                atomic_conc = float(grid.GetCellValue(row, 10))  # Atomic concentration

                # Clean up core level name (remove " p1", " p2" etc.)
                import re
                core_level_clean = re.sub(r'\s+p\d+$', '', core_level)

                # Skip if atomic concentration is 0 (not significant)
                if atomic_conc > 0.1:  # Only show peaks with >0.1% concentration
                    table_data.append((core_level_clean, be, atomic_conc, i))  # Include peak index
            except (ValueError, IndexError):
                continue

        if not table_data:
            return

        # Sort by atomic concentration (highest first)
        table_data.sort(key=lambda x: x[2], reverse=True)

        # Get font size from preferences
        font_size = getattr(self.window, 'label_font_size', 12)

        # Get alpha transparency from config
        table_alpha = getattr(self.parent, 'self.peak_alpha', 0.9)
        name_alpha = 0.6

        # Table dimensions
        row_height = 0.04
        header_height = 0.045
        col_widths = [0.09, 0.09, 0.09]  # Core Level, BE(eV), At.(%)
        table_width = sum(col_widths)
        table_height = header_height + len(table_data) * row_height

        # Get table position from existing table entry in labels or use defaults
        sheet_name = self.window.sheet_combobox.GetValue()
        x_start = 0.02
        y_start = 0.98

        # Check if table already exists in labels
        if 'Labels' in self.window.Data['Core levels'][sheet_name]:
            for label in self.window.Data['Core levels'][sheet_name]['Labels']:
                if label.get('text') == 'Table':
                    x_start = label['x']
                    y_start = label['y']
                    break

        # Draw table background - using config alpha
        from matplotlib.patches import Rectangle, Circle
        table_bg = Rectangle((x_start, y_start - table_height), table_width, table_height,
                             transform=self.ax.transAxes, facecolor='white',
                             edgecolor='black', linewidth=2, alpha=table_alpha)
        self.ax.add_patch(table_bg)
        self.survey_table_text.append(table_bg)

        # Draw table header with merged cells look
        headers = ['', 'BE (eV)', 'At. (%)']
        header_y = y_start - header_height / 2

        x_pos = x_start
        for col, (header, width) in enumerate(zip(headers, col_widths)):
            # Draw header cell background - using config alpha
            header_bg = Rectangle((x_pos, y_start - header_height), width, header_height,
                                  transform=self.ax.transAxes, facecolor=(79/255, 190/255, 159/255),
                                  edgecolor='gray', linewidth=0.5, alpha=table_alpha)
            self.ax.add_patch(header_bg)
            self.survey_table_text.append(header_bg)

            # Add header text
            text_obj = self.ax.text(x_pos + width / 2, header_y, header,
                                    transform=self.ax.transAxes,
                                    fontsize=font_size - 1,
                                    fontweight='bold',
                                    ha='center', va='center',
                                    color='black')
            self.survey_table_text.append(text_obj)

            x_pos += width

        # Draw data rows
        for row_idx, (core_level, be, atomic_conc, peak_idx) in enumerate(table_data):
            y_pos = y_start - header_height - (row_idx + 1) * row_height
            row_y = y_pos + row_height / 2

            x_pos = x_start
            row_data = [core_level, f"{be:.2f}", f"{atomic_conc:.2f}"]

            for col, (data, width) in enumerate(zip(row_data, col_widths)):
                # Set cell background color - peak name column gets peak color, others alternating
                if col == 0:  # First column (peak name)
                    if hasattr(self.window, 'peak_colors') and len(self.window.peak_colors) > 0:
                        cell_color = self.window.peak_colors[peak_idx % len(self.window.peak_colors)]
                    else:
                        cell_color = 'white'
                else:  # Other columns use alternating colors
                    cell_color = 'white' if row_idx % 2 == 0 else '#f8f8f8'

                # Draw cell background - using config alpha
                cell_bg = Rectangle((x_pos, y_pos), width, row_height,
                                    transform=self.ax.transAxes, facecolor=cell_color,
                                    edgecolor='gray', linewidth=0.5, alpha=name_alpha)
                self.ax.add_patch(cell_bg)
                self.survey_table_text.append(cell_bg)

                # Add cell text
                text_obj = self.ax.text(x_pos + width / 2, row_y, data,
                                        transform=self.ax.transAxes,
                                        fontsize=font_size - 1,
                                        ha='center', va='center',
                                        color='black')
                self.survey_table_text.append(text_obj)

                x_pos += width

            # Add/update table entry in Labels data structure with axes coordinates
            if 'Labels' not in self.window.Data['Core levels'][sheet_name]:
                self.window.Data['Core levels'][sheet_name]['Labels'] = []

            labels = self.window.Data['Core levels'][sheet_name]['Labels']

            # Check if Table already exists in labels
            table_found = False
            for label in labels:
                if label.get('text') == 'Table' and label.get('is_table'):
                    # Update existing table position in axes coordinates
                    label['x'] = x_start
                    label['y'] = y_start
                    table_found = True
                    break

            # If table doesn't exist in labels, add it with axes coordinates
            if not table_found:
                table_label = {
                    'text': 'Table',
                    'x': x_start,
                    'y': y_start,
                    'rotation': 0,
                    'fontsize': 10,
                    'fontfamily': 'Arial',
                    'is_table': True,  # Special flag to identify table entities
                    'coordinate_system': 'axes'  # Flag to indicate axes coordinates (0-1)
                }
                labels.append(table_label)

    def format_xas_sheet_name(self, sheet_name):
        """Format XAS sheet names to extract and format the core level name from the XAS~ prefix."""
        if sheet_name.startswith('XAS~'):
            # Remove XAS~ prefix
            name_part = sheet_name[4:]  # Remove 'XAS~'

            # Split by ~ and take the first part (before any trailing number)
            parts = name_part.split('~')
            core_level_name = parts[0]

            # Check if there's already a hyphen/dash in the name
            if '-' not in core_level_name:
                # Use regex to separate element and transition (e.g., "CoL8MN" -> "Co L8MN")
                import re
                match = re.match(r'([A-Z][a-z]?)([L|K|M|N]\d+[A-Z]*)', core_level_name)
                if match:
                    element = match.group(1)
                    transition = match.group(2)

                    # Make the number subscript in the transition (e.g., "L8MN" -> "L₈MN")
                    transition_formatted = re.sub(r'(\d+)', lambda m: ''.join(['₀₁₂₃₄₅₆₇₈₉'[int(d)] for d in m.group(1)]), transition)

                    # Format as "Element Transition edge"
                    formatted_name = f"{element} {transition_formatted} edge"
                    return formatted_name
            else:
                # If there's already a dash, just add subscript to numbers and " edge"
                import re
                # Make numbers subscript
                formatted_name = re.sub(r'(\d+)', lambda m: ''.join(['₀₁₂₃₄₅₆₇₈₉'[int(d)] for d in m.group(1)]), core_level_name)
                formatted_name += " edge"
                return formatted_name

            # Fallback if regex doesn't match
            return core_level_name + " edge"
        else:
            # Fallback for non-XAS sheets or use existing format_sheet_name if available
            if hasattr(self, 'format_sheet_name'):
                return self.format_sheet_name(sheet_name)
            else:
                return sheet_name

    def plot_edx_data(self, window, sheet_name):
        """Plot EDX data when EDX~Plot or EDX~Map sheet is selected"""
        import os
        import numpy as np
        from libraries.ToolsMenu.EDX_SEM_Analysis import open_edx_sem_window

        try:
            # Get the data for this sheet
            if sheet_name not in window.Data['Core levels']:
                return False

            sheet_data = window.Data['Core levels'][sheet_name]

            # Handle EDX~Plot and EDX~Plot1, EDX~Plot2, etc.
            if sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot'):
                # Plot EDX spectrum - support both old and new formats
                if 'B.E.' in sheet_data and 'Raw Data' in sheet_data:
                    energy = np.array(sheet_data['B.E.'])
                    intensity = np.array(sheet_data['Raw Data'])
                elif 'Energy_keV' in sheet_data and 'Intensity' in sheet_data:
                    # OLD FORMAT - still support it
                    energy = np.array(sheet_data['Energy_keV'])
                    intensity = np.array(sheet_data['Intensity'])
                else:
                    print(f"ERROR: {sheet_name} missing required data keys")
                    return False

                # Clear the main plot
                self.ax.clear()

                # Reset background to white for non-EDX plots
                if not (sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')):
                    self.figure.patch.set_facecolor('white')
                    self.ax.set_facecolor('white')
                    # Reset spines to black
                    for spine in self.ax.spines.values():
                        spine.set_edgecolor('black')
                    # Reset tick colors to black
                    self.ax.tick_params(colors='black', labelcolor='black')

                # Remove heatmap colorbar AND its axes if it exists (same as plot_data)
                if hasattr(window, 'heatmap_colorbar') and window.heatmap_colorbar is not None:
                    try:
                        if hasattr(window.heatmap_colorbar, 'ax'):
                            self.figure.delaxes(window.heatmap_colorbar.ax)
                        window.heatmap_colorbar = None
                        window.heatmap_data = None
                        window.heatmap_sheets = None
                    except:
                        pass

                # REMOVE RSD subplot if it exists (same as plot_data)
                if hasattr(self, 'residuals_subplot'):
                    if self.residuals_subplot:
                        self.figure.delaxes(self.residuals_subplot)
                        self.residuals_subplot = None
                        # Restore X-axis visibility
                        self.ax.get_xaxis().set_visible(True)

                # REMOVE RSD text if it exists
                if hasattr(self, 'rsd_text') and self.rsd_text:
                    try:
                        self.rsd_text.remove()
                        self.rsd_text = None
                    except:
                        pass

                # Reset plot position to full size
                self.ax.set_position([0.1, 0.1, 0.85, 0.85])

                # Get stored style preference from window config or sheet data
                style = getattr(window, 'edx_plot_style', 'black')

                print(f'Using EDX plot style: {style}')
                if '_EDX_style' in sheet_data:
                    style = sheet_data['_EDX_style']

                # Apply EDX plot style
                label_color = self.apply_edx_plot_style(self.ax, self.figure, energy, intensity, style)

                # Set title based on selection info if available
                selection_info = sheet_data.get('_EDX_selection')
                title_color = label_color
                if selection_info:
                    sel_type = selection_info.get('type', 'unknown')
                    if sel_type == 'point':
                        x, y = selection_info.get('x', 0), selection_info.get('y', 0)
                        size = selection_info.get('size', 1)
                        if size == 1:
                            self.ax.set_title(f'EDX Spectrum - Point ({x}, {y})', color=title_color)
                        else:
                            self.ax.set_title(f'EDX Spectrum - Point ({x}, {y}) [{size}×{size} px]', color=title_color)
                    elif sel_type == 'line':
                        self.ax.set_title('EDX Spectrum - Line Profile', color=title_color)
                    elif sel_type == 'rectangle':
                        angle = selection_info.get('angle', 0)
                        self.ax.set_title(f'EDX Spectrum - Rectangle (θ={angle:.1f}°)', color=title_color)
                    elif sel_type == 'area':
                        self.ax.set_title('EDX Spectrum - Area', color=title_color)
                    else:
                        self.ax.set_title(f'EDX Spectrum ({sel_type})', color=title_color)
                else:
                    self.ax.set_title('EDX Sum Spectrum', color=title_color)

                self.ax.set_title(' ', color=title_color)

                # Set Y-axis to scientific format
                self.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

                # Set data range for zoom and store in window for later use
                if sheet_name in window.Data['Core levels']:
                    window.Data['Core levels'][sheet_name]['_EDX_min'] = np.min(energy)
                    window.Data['Core levels'][sheet_name]['_EDX_max'] = np.max(energy)

                # Store as window attributes for plotting
                window.x_values = energy
                window.y_values = intensity

                # Set initial display range - check all EDX~Plot sheets for global max
                display_x_max = 20  # default
                for sname in window.Data['Core levels']:
                    if sname == 'EDX~Plot' or sname.startswith('EDX~Plot'):
                        if '_EDX_display_max' in window.Data['Core levels'][sname]:
                            display_x_max = window.Data['Core levels'][sname]['_EDX_display_max']
                            break
                # Store in current sheet too
                sheet_data['_EDX_display_max'] = display_x_max
                self.ax.set_xlim(0, display_x_max)
                self.ax.set_ylim(np.min(intensity) * 0.95, np.max(intensity) * 1.1)

                # Add peak labels using EDX window if available
                hdf5_path = None
                if hasattr(window, 'current_file_path') and window.current_file_path:
                    # Try different HDF5 path variations
                    base_path = window.current_file_path.replace('_EDX.xlsx', '')
                    hdf5_variants = [
                        f"{base_path}_EDX.hdf5",
                        f"{base_path}_EDX.h5",
                        f"{base_path}.hdf5",
                        f"{base_path}.h5"
                    ]

                    for variant in hdf5_variants:
                        if os.path.exists(variant):
                            hdf5_path = variant
                            break

                # Also check if stored in sheet data
                if not hdf5_path:
                    # HDF5 path is derived from FilePath
                    from libraries.ToolsMenu.EDX_SEM_Analysis import get_hdf5_path_from_filepath
                    hdf5_path = get_hdf5_path_from_filepath(window.Data.get('FilePath'))


                # Add peak labels - check if they already exist
                if len(self.ax.texts) == 0:
                    # No labels yet - add them
                    try:
                        from libraries.ToolsMenu.EDX_SEM_Analysis import EDXSEMWindow
                        temp_edx = EDXSEMWindow(window)
                        temp_edx.add_peak_labels(self.ax, energy, intensity)

                        # Apply label color based on style
                        for text in self.ax.texts:
                            text.set_color(label_color)
                            text.set_fontweight('bold')

                        temp_edx.Destroy()
                        print("Added EDX peak labels")
                    except Exception as e:
                        print(f"Could not add EDX labels: {e}")
                        import traceback
                        traceback.print_exc()

                self.canvas.draw()

                # Draw saved selection on EDX map window if open
                if hasattr(window, 'edx_window') and window.edx_window:
                    try:
                        if not window.edx_window.IsBeingDeleted():
                            # Clear any previous saved selection
                            window.edx_window._clear_saved_selection()
                            # Draw the selection for this plot if available
                            if selection_info:
                                window.edx_window.draw_saved_selection(selection_info)
                                window.edx_window.map_canvas.draw_idle()
                    except Exception as e:
                        print(f"Could not draw saved selection: {e}")

                # Restore grid data if available
                grid_data = sheet_data.get('_EDX_grid_data')
                if grid_data and hasattr(window, 'peak_params_grid'):
                    from libraries.Sheet_Operations import _restore_grid_data
                    _restore_grid_data(window, grid_data)

                return True

            elif sheet_name == 'EDX~Map':
                # Open EDX HeatMap window if not already open
                from libraries.ToolsMenu.EDX_SEM_Analysis import open_edx_sem_window

                # Check if EDX window exists and is not deleted
                if hasattr(window, 'edx_window') and window.edx_window:
                    try:
                        if not window.edx_window.IsBeingDeleted():
                            # Window exists, just raise it to front
                            window.edx_window.Raise()
                            window.edx_window.SetFocus()
                            return True
                    except (RuntimeError, wx.wxAssertionError):
                        # Window was deleted, clear reference
                        window.edx_window = None

                # Window doesn't exist or was deleted, create new one
                window.edx_window = open_edx_sem_window(window)

                # Try to load the HDF5 data if available
                from libraries.ToolsMenu.EDX_SEM_Analysis import get_hdf5_path_from_filepath
                hdf5_path = get_hdf5_path_from_filepath(window.Data.get('FilePath'))

                if hdf5_path and os.path.exists(hdf5_path):
                    try:
                        window.edx_window.load_file(hdf5_path, 'EDX Map')
                        print(f"Loaded EDX map from: {hdf5_path}")
                    except Exception as e:
                        print(f"Could not load EDX map: {e}")
                        import traceback
                        traceback.print_exc()
                else:
                    print(f"HDF5 file not found: {hdf5_path}")

                return True

            return False

        except Exception as e:
            print(f"Error plotting EDX data: {e}")
            import traceback
            traceback.print_exc()
            return False

    def apply_edx_plot_style(self, ax, figure, energy, intensity, style='black', label_color='black'):
        """Apply EDX plot style to axes

        Args:
            ax: matplotlib axes
            figure: matplotlib figure
            energy: energy array
            intensity: intensity array
            style: 'black', 'blue_yellow', 'white_red', or 'white_blue'
            label_color: color for peak labels (used after add_peak_labels is called)
        """
        if style == 'black':
            # Default style: Black line, black labels, white background
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.plot(energy, intensity, 'k-', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        # elif style == 'blue_yellow':
        #     # Blue background with yellow filled line and yellow labels
        #     figure.patch.set_facecolor('#1e3a5f')
        #     ax.set_facecolor('#1e3a5f')
        #     ax.fill_between(energy, intensity, color='yellow', alpha=0.8)
        #     ax.plot(energy, intensity, 'yellow', linewidth=1.2)
        #     ax.set_xlabel('Energy (keV)', color='yellow')
        #     ax.set_ylabel('Counts', color='yellow')
        #     ax.tick_params(colors='yellow', labelcolor='yellow')
        #     for spine in ax.spines.values():
        #         spine.set_edgecolor('yellow')
        #     return 'yellow'

        elif style == 'blue_yellow':
            # Blue axes background with yellow filled line, white figure background, black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('#1e3a5f')
            ax.fill_between(energy, intensity, color='yellow', alpha=0.8)
            ax.plot(energy, intensity, 'yellow', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'yellow'

        elif style == 'white_green':
            # White background with red filled line and black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.fill_between(energy, intensity, color='#16A085', alpha=0.6)
            ax.plot(energy, intensity, '#16A085', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        elif style == 'white_red':
            # White background with red filled line and black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.fill_between(energy, intensity, color='red', alpha=0.6)
            ax.plot(energy, intensity, 'red', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        elif style == 'white_blue':
            # White background with blue filled line and black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.fill_between(energy, intensity, color='blue', alpha=0.6)
            ax.plot(energy, intensity, 'blue', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        elif style == 'white_purple':
            # White background with blue filled line and black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.fill_between(energy, intensity, color='purple', alpha=0.6)
            ax.plot(energy, intensity, 'purple', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        elif style == 'white_pink':
            # White background with blue filled line and black labels
            figure.patch.set_facecolor('white')
            ax.set_facecolor('white')
            ax.fill_between(energy, intensity, color='pink', alpha=0.8)
            ax.plot(energy, intensity, 'purple', linewidth=1.0)
            ax.set_xlabel('Energy (keV)', color='black')
            ax.set_ylabel('Counts', color='black')
            ax.tick_params(colors='black', labelcolor='black')
            for spine in ax.spines.values():
                spine.set_edgecolor('black')
            return 'black'

        # Default fallback
        return 'black'

    def plot_eels_data(self, window, sheet_name):
        """Plot EELS data when EELS~Plot or EELS~Map sheet is selected"""
        import os
        import numpy as np

        try:
            # Get the data for this sheet
            if sheet_name not in window.Data['Core levels']:
                return False

            sheet_data = window.Data['Core levels'][sheet_name]

            # Handle EELS~Plot and EELS~Plot1, EELS~Plot2, etc.
            if sheet_name == 'EELS~Plot' or sheet_name.startswith('EELS~Plot'):
                # Plot EELS spectrum - support both old and new formats
                if 'B.E.' in sheet_data and 'Raw Data' in sheet_data:
                    energy = np.array(sheet_data['B.E.'])
                    intensity = np.array(sheet_data['Raw Data'])
                else:
                    print(f"ERROR: {sheet_name} missing required data keys")
                    return False

                # Clear the main plot and all text annotations
                self.ax.clear()

                # Explicitly remove any lingering text labels
                for txt in self.ax.texts[:]:
                    try:
                        txt.remove()
                    except:
                        pass

                # Reset background to white for non-EDX plots
                if not (sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')):
                    self.figure.patch.set_facecolor('white')
                    self.ax.set_facecolor('white')
                    # Reset spines to black
                    for spine in self.ax.spines.values():
                        spine.set_edgecolor('black')
                    # Reset tick colors to black
                    self.ax.tick_params(colors='black', labelcolor='black')

                # Remove heatmap colorbar AND its axes if it exists
                if hasattr(window, 'heatmap_colorbar') and window.heatmap_colorbar is not None:
                    try:
                        if hasattr(window.heatmap_colorbar, 'ax'):
                            self.figure.delaxes(window.heatmap_colorbar.ax)
                        window.heatmap_colorbar = None
                        window.heatmap_data = None
                        window.heatmap_sheets = None
                    except:
                        pass

                # REMOVE RSD subplot if it exists
                if hasattr(self, 'residuals_subplot'):
                    if self.residuals_subplot:
                        self.figure.delaxes(self.residuals_subplot)
                        self.residuals_subplot = None
                        self.ax.get_xaxis().set_visible(True)

                # REMOVE RSD text if it exists
                if hasattr(self, 'rsd_text') and self.rsd_text:
                    try:
                        self.rsd_text.remove()
                        self.rsd_text = None
                    except:
                        pass

                # Reset plot position to full size
                self.ax.set_position([0.1, 0.1, 0.85, 0.85])

                # Plot EELS spectrum
                self.ax.plot(energy, intensity, 'k-', linewidth=0.8)
                self.ax.set_xlabel('Energy Loss (eV)')
                self.ax.set_ylabel('Counts')

                # Set title based on selection info if available
                selection_info = sheet_data.get('_EELS_selection')
                if selection_info:
                    sel_type = selection_info.get('type', 'unknown')
                    if sel_type == 'point':
                        x, y = selection_info.get('x', 0), selection_info.get('y', 0)
                        size = selection_info.get('size', 1)
                        if size == 1:
                            self.ax.set_title(f'EELS Spectrum - Point ({x}, {y})')
                        else:
                            self.ax.set_title(f'EELS Spectrum - Point ({x}, {y}) [{size}×{size} px]')
                    elif sel_type == 'line':
                        self.ax.set_title('EELS Spectrum - Line Profile')
                    elif sel_type == 'area':
                        self.ax.set_title('EELS Spectrum - Area')
                    else:
                        self.ax.set_title(f'EELS Spectrum ({sel_type})')
                else:
                    self.ax.set_title('EELS Sum Spectrum')

                # Set Y-axis to scientific format
                self.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

                # Store data range
                if sheet_name in window.Data['Core levels']:
                    window.Data['Core levels'][sheet_name]['_EELS_min'] = np.min(energy)
                    window.Data['Core levels'][sheet_name]['_EELS_max'] = np.max(energy)

                # Store as window attributes for plotting
                window.x_values = energy
                window.y_values = intensity

                # Set display range
                self.ax.set_xlim(np.min(energy), np.max(energy))
                self.ax.set_ylim(np.min(intensity) * 0.95, np.max(intensity) * 1.1)

                self.canvas.draw()

                # Draw saved selection on EELS map window if open
                if hasattr(window, 'eels_window') and window.eels_window:
                    try:
                        if not window.eels_window.IsBeingDeleted():
                            window.eels_window._clear_saved_selection()
                            if selection_info:
                                window.eels_window.draw_saved_selection(selection_info)
                                window.eels_window.map_canvas.draw_idle()
                    except Exception as e:
                        print(f"Could not draw saved selection: {e}")

                return True

            elif sheet_name == 'EELS~Map':
                # Open EELS window for map view
                from libraries.ToolsMenu.EELS_Analysis import open_eels_window

                if not hasattr(window, 'eels_window') or window.eels_window is None:
                    window.eels_window = open_eels_window(window)

                    # Try to reload data if source path is available
                    source_path = sheet_data.get('_Source_Path')
                    if source_path and os.path.exists(source_path):
                        window.eels_window.load_file(source_path)
                else:
                    window.eels_window.Raise()

                return True

            return False

        except Exception as e:
            print(f"Error plotting EELS data: {e}")
            import traceback
            traceback.print_exc()
            return False

    #STart