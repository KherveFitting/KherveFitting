# plot_operations.py

import re
import wx
import os
import lmfit
import numpy as np
import numpy.ma as ma
import matplotlib.pyplot as plt
from matplotlib import colormaps
from matplotlib.ticker import AutoMinorLocator, ScalarFormatter
from itertools import cycle
from PIL import Image, ImageDraw, ImageFont
import matplotlib.colors as mcolors
from scipy.ndimage import gaussian_filter

from libraries.Peak_Functions import PeakFunctions, BackgroundCalculations, OtherCalc

from libraries.Save import save_state


class PlotManager:
    def __init__(self, ax, canvas):
        self.ax = ax
        self.canvas = canvas
        self.figure = ax.figure # Add this line
        self.cross = None
        self.peak_fill_enabled = True
        self.fitting_results_text = None
        self.fitting_results_visible = False

        self.residuals_state = 0  # Add this line
        self.residuals_subplot = None  # Add this line
        self.residuals_visible = True  # Keep existing

        # init for preference window
        self.plot_style = "scatter"
        self.scatter_size = 20
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

    def toggle_y_axis_OLD(self):
        self.y_axis_visible = not self.y_axis_visible
        self.ax.yaxis.set_visible(self.y_axis_visible)

        # Also toggle y-axis for residuals subplot if it exists
        if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
            self.residuals_subplot.yaxis.set_visible(self.y_axis_visible)

        self.canvas.draw_idle()

    def toggle_y_axis(self):
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

            # Identify doublets
            num_peaks = window.peak_params_grid.GetNumberRows() // 2
            doublets = []
            for i in range(0, num_peaks - 1):
                current_label = window.peak_params_grid.GetCellValue(i * 2, 1)
                next_label = window.peak_params_grid.GetCellValue((i + 1) * 2, 1)
                if self.is_part_of_doublet(current_label, next_label):
                    doublets.extend([i, i + 1])

            # Find current peak index
            for i in range(num_peaks):
                if window.peak_params_grid.GetCellValue(i * 2, 1) == peak_label:
                    peak_index = i
                    break

            # If part of doublet, get fill type from first peak of the pair
            if peak_index in doublets:
                if doublets.index(peak_index) % 2 == 1:  # Second peak of doublet
                    peak_index = peak_index - 1  # Use first peak's settings

            if window.peak_fill_types[peak_index] == "Solid Fill":
                fill_params = {
                    'color': color,
                    'alpha': alpha,
                    'edgecolor': 'none'
                }
            else:  # Hatch
                fill_params = {
                    'color': 'none',
                    'hatch': window.peak_hatch_patterns[peak_index] * window.hatch_density,
                    'linewidth': window.peak_line_thickness,
                    'edgecolor': color,
                    'alpha': alpha
                }

            if window.energy_scale == 'KE':
                self.ax.fill_between(window.photons - x_values, background, peak_y,
                                     interpolate=True, label=peak_label, **fill_params)
            else:
                self.ax.fill_between(x_values, background, peak_y,
                                     interpolate=True, label=peak_label, **fill_params)

            if window.peak_line_style != "No Line":
                if window.peak_line_style == "Black":
                    line_color = "black"
                elif window.peak_line_style == "Grey":
                    line_color = "grey"
                elif window.peak_line_style == "Yellow":
                    line_color = "yellow"
                else:  # same_color
                    line_color = color
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

        self.canvas.draw_idle()

        return peak_y

    def plot_data(self, window):
        if 'Core levels' not in window.Data or not window.Data['Core levels']:
            wx.MessageBox("No data available to plot.", "Error", wx.OK | wx.ICON_ERROR)
            return

        sheet_name = window.sheet_combobox.GetValue()
        is_d_parameter = False
        limits = window.plot_config.get_plot_limits(window, sheet_name)
        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            self.ax.clear()

            if hasattr(self, 'residuals_subplot'):
                if self.residuals_subplot:
                    self.figure.delaxes(self.residuals_subplot)
                    self.residuals_subplot = None
                    # self.ax.set_position([0.1, 0.1, 0.8, 0.8])
                    self.ax.set_position([0.1, 0.125, 0.85, 0.85])
                    self.ax.get_xaxis().set_visible(True)

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

                for label_data in window.Data['Core levels'][sheet_name]['Labels']:
                    window.ax.text(
                        label_data['x'],
                        label_data['y'],
                        label_data['text'],
                        rotation=label_data.get('rotation', 90),
                        va='bottom',
                        ha='center',
                        fontsize=window.label_font_size
                    )


            # Hide the cross if it exists
            if hasattr(window, 'cross') and window.cross:
                window.cross.remove()

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
                self.ax.legend(loc='upper left')

            # Check if a peak is selected and add cross
            if window.selected_peak_index is not None:
                window.plot_manager.add_cross_to_peak(window, window.selected_peak_index)

            # Remove any existing sheet name text
            for txt in self.ax.texts:
                if getattr(txt, 'sheet_name_text', False):
                    txt.remove()

            # Format and add sheet name text
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

            # Plot background only if it's different from raw data
            if 'Background' in window.Data['Core levels'][sheet_name] and 'Bkg Y' in \
                    window.Data['Core levels'][sheet_name]['Background']:
                raw_data = np.array(window.Data['Core levels'][sheet_name]['Raw Data'])
                background = np.array(window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])

                if not np.array_equal(raw_data, background):
                    # if "survey" in sheet_name.lower() or "wide" in sheet_name.lower():
                    #     pass
                    # else:
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

            self.canvas.draw()  # Update the plot

        except Exception as e:
            wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)


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

        sheet_name = window.sheet_combobox.GetValue()
        if not sheet_name or 'Core levels' not in window.Data or sheet_name not in window.Data['Core levels']:
            return
        limits = window.plot_config.get_plot_limits(window, sheet_name)
        cst_unfit = ""

        if sheet_name not in window.Data['Core levels']:
            wx.MessageBox(f"No data available for sheet: {sheet_name}", "Error", wx.OK | wx.ICON_ERROR)
            return

        core_level_data = window.Data['Core levels'][sheet_name]

        if window.energy_scale == 'KE':
            self.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])  # Reverse X-axis
        else:
            self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis

        # self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
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
                self.ax.set_position([0.1, 0.125, 0.85, 0.85])
                gs = self.figure.add_gridspec(20, 1, hspace=0.0)
                self.ax.set_position(gs[0:18, 0].get_position(self.figure))  # Main plot takes 6/8
                self.residuals_subplot.set_position(gs[18:, 0].get_position(self.figure))  # Residuals takes 2/8
                self.residuals_subplot.sharex(self.ax)

                # Explicitly ensure visibility
                self.residuals_subplot.set_visible(True)

        else:
            # self.ax.set_position([0.1, 0.1, 0.8, 0.8])
            self.ax.set_position([0.1, 0.125, 0.85, 0.85])
            self.ax.get_xaxis().set_visible(True)


        x_label = "Kinetic Energy (eV)" if window.energy_scale == 'KE' else window.x_axis_label
        self.ax.set_xlabel(x_label)
        # self.ax.set_xlabel("Binding Energy (eV)")
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

        # Set plot limits

        self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis

        if window.energy_scale == 'KE':
            X_MIN = window.photons - limits['Xmax']
            X_MAX = window.photons - limits['Xmin']
            self.ax.set_xlim(min(X_MIN, X_MAX), max(X_MIN, X_MAX))  # Reverse X-axis
            # self.ax.set_xlim(window.photons - limits['Xmax'],window.photons - limits['Xmin'])  # Reverse X-axis

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
            if fitting_model == "SurveyID":
                cst_unfit = "SurveyID"
            if 'Labels' in window.Data['Core levels'][sheet_name]:

                for label_data in window.Data['Core levels'][sheet_name]['Labels']:
                    window.ax.text(
                        label_data['x'],
                        label_data['y'],
                        label_data['text'],
                        rotation=label_data.get('rotation', 90),
                        va = 'bottom',
                        ha = 'center',
                        fontsize = window.label_font_size
                    )
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
                        alpha = self.peak_alpha * 0.8  # Reduce alpha for the second peak
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
                                                color=color, alpha=alpha) #, linewidth=self.background_thickness)
                else:
                    self.plot_peak(window.x_values, window.background, peak_params, sheet_name, window,
                                                color=color, alpha=alpha)


        # # Plot the background if it exists
        # if window.energy_scale == 'KE':
        #     self.ax.plot(window.photons - x_values, core_level_data['Background']['Bkg Y'],
        #                  color=self.background_color, linewidth=self.background_thickness,
        #                  linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background')
        # else:
        #     self.ax.plot(x_values, core_level_data['Background']['Bkg Y'], color=self.background_color,
        #                  linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background',
        #                  linewidth=self.background_thickness)

        # Only plot background if it's different from raw data or if Bkg Type is not empty
        if (core_level_data['Background'].get('Bkg Type') != "" and
                core_level_data['Background'].get('Bkg Low') != "" and
                core_level_data['Background'].get('Bkg High') != ""):
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - x_values, core_level_data['Background']['Bkg Y'],
                             color=self.background_color, linewidth=self.background_thickness,
                             linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background')
            else:
                self.ax.plot(x_values, core_level_data['Background']['Bkg Y'], color=self.background_color,
                             linestyle=self.background_linestyle, alpha=self.background_alpha, label='Background',
                             linewidth=self.background_thickness)


        # Update overall fit and residuals
        if cst_unfit in ["Unfitted","D-parameter","SurveyID"] or any(x in sheet_name.lower() for x in ["survey", "wide"]):
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
            self.ax.legend(loc='upper left')
            self.update_legend(window)
        else:
            self.ax.legend().set_visible(False)


        # Restore sheet name text or create new one if it doesn't exist
        if sheet_name_text is None:
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

        # Add this line before canvas.draw_idle()
        self.apply_text_settings(window)

        # Draw the canvas
        self.canvas.draw_idle()
        # window.update_checkbox_visuals()

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
        if window.x_values is None or window.background is None:
            print("Error: x_values or background is None. Cannot update peak plot.")
            return

        if x is None or y is None:
            print("Error: x or y is None. Cannot update peak plot.")
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
            elif window.selected_fitting_method == "D-parameter":
                return area, 0, 0  # Return original area and zero for normalized/relative areas
            else:  # Default to GL (Height) as a safe bet
                peak_model = lmfit.Model(PeakFunctions.gauss_lorentz)
                params = peak_model.make_params(center=x, fwhm=fwhm, fraction=lg_ratio, amplitude=y)

            peak_y = peak_model.eval(params, x=window.x_values) + window.background

            # Update overall fit and residuals
            if peak_model in ["D-parameter", "SurveyID"]:
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
                # sigma = float(window.peak_params_grid.GetCellValue(row, 7))
                gamma = float(window.peak_params_grid.GetCellValue(row, 8)) / 2
                # gamma = float(window.peak_params_grid.GetCellValue(row, 8))
                # amplitude = peak_y / peak_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0, skew=skew)
                amplitude = float(window.peak_params_grid.GetCellValue(row, 6))
                params = peak_model.make_params(center=peak_x, amplitude=amplitude, sigma=sigma, gamma=gamma, skew=skew)
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
            elif fitting_model == "SurveyID":
                # Skip D-parameter in overall fit calculation
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

        # Start of change for new residuals

        # Create a masked array where 0 values are masked
        masked_residuals = ma.masked_where(np.isclose(scaled_residuals, 0, atol=5e-1), scaled_residuals)
        masked_residuals2 = ma.masked_where(np.isclose(scaled_residuals, 0, atol=5e-1), residuals)

        # Remove old overall fit and residuals, keep background lines
        for line in self.ax.lines:
            if line.get_label() in ['Overall Fit', 'Residuals']:
                line.remove()



        # Plot the overall fit
        try:
            good_indices = ~np.isnan(overall_fit)
            x_plot = x_values[good_indices]
            y_plot = overall_fit[good_indices]
            # x_plot = x_values
            # y_plot = overall_fit
            if window.energy_scale == 'KE':
                self.ax.plot(window.photons - window.x_values, overall_fit, color=self.envelope_color,
                             linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                             linewidth=self.envelope_thickness,
                             label='D-parameter' if fitting_model == "D-parameter" else 'Overall Fit')
            else:
                # self.ax.plot(window.x_values, overall_fit, color=self.envelope_color,
                self.ax.plot(x_plot, y_plot, color=self.envelope_color,
                             linestyle=self.envelope_linestyle, alpha=self.envelope_alpha,
                             linewidth=self.envelope_thickness,
                             label='D-parameter' if fitting_model == "D-parameter" else 'Overall Fit')
        except:
            return

        # Handle residuals based on state
        if hasattr(self, 'residuals_state'):
            if self.residuals_state == 1:  # On main plot
                residual_height = 1.07 * max(window.y_values)
                residual_base = self.ax.axhline(y=residual_height, color='grey', linestyle='-.', alpha=0.1)

                if window.energy_scale == 'KE':
                    residual_line = self.ax.plot(window.photons - window.x_values, masked_residuals + residual_height,
                                                 color=self.residual_color, linestyle=self.residual_linestyle,
                                                 alpha=self.residual_alpha, label='Residuals',
                                                 linewidth=self.residual_thickness)
                else:
                    residual_line = self.ax.plot(window.x_values, masked_residuals + residual_height,
                                                 color=self.residual_color, linestyle=self.residual_linestyle,
                                                 alpha=self.residual_alpha, label='Residuals',
                                                 linewidth=self.residual_thickness)

                residual_line[0].set_visible(True)
                residual_base.set_visible(True)
                self.ax.get_xaxis().set_visible(True)
            elif self.residuals_state == 2:  # Separate subplot
                self.setup_residual_subplot(window, x_values, masked_residuals, self.residual_thickness,
                                            scaling_factor=1.0)

            else:
                self.ax.get_xaxis().set_visible(True)

        # Handle RSD text
        rsd = PeakFunctions.calculate_rsd(window.y_values, overall_fit)
        if rsd is not None:
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
                                                 f'RSD: {rsd:.2f}',
                                                 horizontalalignment='right',
                                                 verticalalignment='center',
                                                 fontsize=9,
                                                 color=self.residual_color,
                                                 alpha=self.residual_alpha + 0.2,
                                                 bbox=dict(facecolor='white', edgecolor='none'))
            elif self.residuals_state == 2:  # For subplot, don't check plot limits
                if self.residuals_subplot:
                    if window.energy_scale == 'KE':
                        x_min = self.residuals_subplot.get_xlim()[1] - 0.4
                    else:
                        x_min = self.residuals_subplot.get_xlim()[1] + 0.4
                    y_pos = np.mean(self.residuals_subplot.get_ylim())
                    if self.rsd_text:
                        try:
                            self.rsd_text.remove()
                        except:
                            print("Error: RSD text removal failed.")
                            pass
                    self.rsd_text = self.residuals_subplot.text(x_min, y_pos,
                                                                f'RSD: {rsd:.2f}',
                                                                horizontalalignment='right',
                                                                verticalalignment='center',
                                                                fontsize=9,
                                                                color=self.residual_color,
                                                                alpha=self.residual_alpha + 0.2,
                                                                bbox=dict(facecolor='white', edgecolor='none'))

        # Only update main plot ylabel if residuals are not in subplot
        if self.residuals_state != 2:
            self.ax.set_ylabel(f'Intensity (CPS), residual x {scaling_factor:.2f}')
        else:
            self.ax.set_ylabel('Intensity (CPS)')

        self.canvas.draw_idle()
        return residuals

    def setup_residual_subplot(self, window, x_values, masked_residuals, residual_thickness=1, scaling_factor=1):
        # Create gridspec at start
        gs = self.figure.add_gridspec(20, 1, hspace=0.0)

        if not self.residuals_subplot:
            self.ax.set_position(gs[0:18, 0].get_position(self.figure))
            self.residuals_subplot = self.figure.add_subplot(gs[18:, 0])

        self.residuals_subplot.clear()

        # Determine x values based on energy scale
        x_plot = window.photons - x_values if window.energy_scale == 'KE' else x_values

        # Plot residuals
        self.residuals_subplot.plot(x_plot, masked_residuals * scaling_factor,
                                    color=self.residual_color,
                                    linestyle=self.residual_linestyle,
                                    alpha=self.residual_alpha,
                                    linewidth=residual_thickness)

        # Configure main plot
        self.ax.get_xaxis().set_visible(False)

        # Configure subplot
        self.residuals_subplot.set_ylabel('Res.')
        x_label = "Kinetic Energy (eV)" if window.energy_scale == 'KE' else "Binding Energy (eV)"
        self.residuals_subplot.set_xlabel(x_label)
        # self.residuals_subplot.set_xlabel('Binding Energy (eV)')
        self.residuals_subplot.tick_params(axis='x', bottom=True, labelbottom=True,
                                           labelsize=window.axis_number_size, pad=8)
        self.residuals_subplot.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.residuals_subplot.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        # Set font sizes
        self.residuals_subplot.xaxis.label.set_size(window.axis_title_size)
        self.residuals_subplot.yaxis.label.set_size(window.axis_title_size)

        # Set y limits with margin
        y_min, y_max = np.min(masked_residuals), np.max(masked_residuals)
        margin = 0.1 * (y_max - y_min)
        self.residuals_subplot.set_ylim(y_min - margin, y_max + margin)

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
            window.plot_manager.add_cross_to_peak(window, window.selected_peak_index)

            # Redraw the canvas
            window.canvas.draw_idle()

    def add_cross_to_peak(self, window, index):
        try:
            row = index * 2  # Each peak uses two rows in the grid
            peak_x = float(window.peak_params_grid.GetCellValue(row, 2))  # Position
            peak_y = float(window.peak_params_grid.GetCellValue(row, 3))  # Height

            # Find the closest background value
            closest_index = np.argmin(np.abs(window.x_values - peak_x))
            bkg_y = window.background[closest_index]

            # Add background to peak height
            peak_y += bkg_y

            # Remove existing cross if it exists
            if self.cross:
                self.cross.remove()

            # Plot new cross
            self.cross, = self.ax.plot(peak_x, peak_y, 'bx', markersize=15, markerfacecolor='none', picker=5,
                                       linewidth=3)

            # Get peak letter (A, B, C etc)
            peak_letter = chr(65 + index)

            # Add letter above cross with offset
            max_y = window.ax.get_ylim()[1]
            y_offset = max_y * 0.02  # 2% of plot height for offset
            self.ax.text(peak_x, peak_y + y_offset, peak_letter,
                         ha='center', va='bottom', fontsize=12)

            # Connect event handlers
            self.canvas.mpl_disconnect('motion_notify_event')  # Disconnect existing handlers
            self.canvas.mpl_disconnect('button_release_event')
            self.motion_notify_id = self.canvas.mpl_connect('motion_notify_event', window.on_cross_drag)
            self.button_release_id = self.canvas.mpl_connect('button_release_event', window.on_cross_release)

            # Redraw canvas
            self.canvas.draw_idle()

        except ValueError as e:
            print(f"Error adding cross to peak: {e}")
            # You might want to show an error message to the user here
        except Exception as e:
            print(f"Unexpected error adding cross to peak: {e}")
            # You might want to show an error message to the user here


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

    def toggle_legend_OLD(self):
        self.legend_visible = not self.legend_visible
        legend = self.ax.get_legend()
        if legend:
            legend.set_visible(self.legend_visible)
        self.canvas.draw_idle()

    def toggle_legend(self):
        self.legend_visible = (self.legend_visible + 1) % 3
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
                self.ax.legend(filtered_handles, filtered_labels, loc='upper left')
        self.canvas.draw_idle()

    def update_legend_OLD(self, window):
        sheet_name = window.sheet_combobox.GetValue()  # Add this line
        handles, labels = self.ax.get_legend_handles_labels()

        if self.legend_visible == 2:
            # Filter out Raw Data, Background, Overall Fit
            handles = [h for h, l in zip(handles, labels)
                       if l not in ["Raw Data", "Background", "Overall Fit"]]
            labels = [l for l in labels
                      if l not in ["Raw Data", "Background", "Overall Fit"]]

        has_overall_fit = "Overall Fit" in labels
        has_raw_data = "Raw Data" in labels

        if hasattr(self, 'residuals_state') and self.residuals_state == 2:
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
        else:
            legend_order = []
            legend_order2 = []
            if has_raw_data:
                legend_order.append("Raw Data")
                legend_order2.append("Raw Data")
            legend_order.append("Background")
            legend_order2.append("Background")
            if has_overall_fit:
                legend_order.extend(["Overall Fit", "Residuals"])
                legend_order2.extend(["Overall Fit", "Residuals"])
            elif "Residuals" in labels:
                legend_order.append("Residuals")
                legend_order2.append("Residuals")

        num_peaks = window.peak_params_grid.GetNumberRows() // 2
        peak_labels = []
        filtered_peak_labels = []
        for i in range(num_peaks):
            label = window.peak_params_grid.GetCellValue(i * 2, 1)
            formatted_label = re.sub(r'(\d+/\d+)', r'$_{\1}$', label)
            clean_label = re.sub(r'\$.*?\$', '', formatted_label)
            split_label = clean_label.split()

            if "survey" in sheet_name.lower() or "wide" in sheet_name.lower():
                peak_labels.append(label)
                filtered_peak_labels.append(formatted_label)
            else:
                if len(split_label) > 1 and split_label[1].strip():
                    peak_labels.append(label)
                    filtered_peak_labels.append(formatted_label)

        legend_order += peak_labels
        legend_order2 += filtered_peak_labels

        if legend_order and self.legend_visible:
            ordered_handles = []
            for l in legend_order:
                for index, label in enumerate(labels):
                    if label == l:
                        ordered_handles.append(handles[index])
                        break


            self.ax.legend(ordered_handles, legend_order2, loc='upper left')
        else:
            self.ax.legend().remove()
            self.ax.legend().set_visible(False)

        self.canvas.draw_idle()

    def update_legend(self, window):
        sheet_name = window.sheet_combobox.GetValue()
        handles, labels = self.ax.get_legend_handles_labels()

        num_peaks = window.peak_params_grid.GetNumberRows() // 2
        peak_labels = []
        filtered_peak_labels = []

        for i in range(num_peaks):
            label = window.peak_params_grid.GetCellValue(i * 2, 1)
            formatted_label = re.sub(r'(\d+/\d+)', r'$_{\1}$', label)
            clean_label = re.sub(r'\$.*?\$', '', formatted_label)
            split_label = clean_label.split()
            if len(split_label) > 1 and split_label[1].strip():
                peak_labels.append(label)
                filtered_peak_labels.append(formatted_label)

        if self.legend_visible == 2:
            ordered_handles = []
            for l in peak_labels:
                for index, label in enumerate(labels):
                    if label == l:
                        ordered_handles.append(handles[index])
                        break
            if ordered_handles and filtered_peak_labels:
                self.ax.legend(ordered_handles, filtered_peak_labels, loc='upper left')
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
            legend_order2 += filtered_peak_labels

            if legend_order and self.legend_visible:
                ordered_handles = []
                for l in legend_order:
                    for index, label in enumerate(labels):
                        if label == l:
                            ordered_handles.append(handles[index])
                            break
                self.ax.legend(ordered_handles, legend_order2, loc='upper left')
            else:
                self.ax.legend().remove()
                self.ax.legend().set_visible(False)

        self.canvas.draw_idle()

    @staticmethod
    def format_sheet_name(sheet_name):
        import re
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

    def plot_background(self, window):
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
                background_filtered, label = self._calculate_adaptive_smart_background(window, x_values, y_values,
                                                                                       offset_h, offset_l)
            else:
                background_filtered, label = self._calculate_other_background(window, x_values, y_values, method,
                                                                              offset_h, offset_l)

            # Update the background data in the window.Data structure
            self._update_background_data(window, sheet_name, x_values, background_filtered, method, offset_h, offset_l)
            window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = background_filtered.tolist()
            window.background = background_filtered

            # Plot the calculated background
            self.ax.plot(x_values, window.background, color='grey', linestyle='--', label=label,
                         linewidth=self.background_thickness)

            # Replot everything if peaks exist
            if window.peak_params_grid.GetNumberRows() > 0:
                window.clear_and_replot()
                self.update_legend(window)

            self.ax.legend(loc='upper left')
            self.canvas.draw()

        except Exception as e:
            print("Error in plot_background:", str(e))
            import traceback
            traceback.print_exc()
            wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)

    def _calculate_adaptive_smart_background(self, window, x_values, y_values, offset_h, offset_l):
        """Helper method to calculate Multi-Regions Smart background."""
        sheet_name = window.sheet_combobox.GetValue()  # Get the current sheet name
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
            x_values, y_values, adaptive_range, current_background, offset_h, offset_l
        )
        return background_filtered, 'Background (Multi-Regions Smart)'

    def _calculate_other_background(self, window, x_values, y_values, method, offset_h, offset_l):
        """Helper method to calculate background for non-Multi-Regions Smart methods."""
        sheet_name = window.sheet_combobox.GetValue()
        bg_min_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg Low')
        bg_max_energy = window.Data['Core levels'][sheet_name]['Background'].get('Bkg High')

        if bg_min_energy is None or bg_max_energy is None or bg_min_energy > bg_max_energy:
            wx.MessageBox("Invalid energy range selected.", "Warning", wx.OK | wx.ICON_INFORMATION)
            return None, None

        # Add this before the mask line to remove the
        try:
            bg_min_energy = float(bg_min_energy) if bg_min_energy != '' else min(x_values)
            bg_max_energy = float(bg_max_energy) if bg_max_energy != '' else max(x_values)
        except (ValueError, TypeError):
            # If conversion fails, use default values
            bg_min_energy = min(x_values)
            bg_max_energy = max(x_values)

        mask = (x_values >= bg_min_energy) & (x_values <= bg_max_energy)
        x_values_filtered = x_values[mask]
        y_values_filtered = y_values[mask]

        if method == "Shirley":
            background_filtered = BackgroundCalculations.calculate_shirley_background(x_values_filtered,
                                                                                      y_values_filtered, offset_h,
                                                                                      offset_l)
            label = 'Background (Shirley)'
        elif method == "Linear":
            background_filtered = BackgroundCalculations.calculate_linear_background(x_values_filtered,
                                                                                     y_values_filtered, offset_h,
                                                                                     offset_l)
            label = 'Background (Linear)'
        elif method in ["Smart", "Multi-Regions Smart", "Multiple Regions Smart"]:
            background_filtered = BackgroundCalculations.calculate_smart_background(x_values_filtered,
                                                                                    y_values_filtered, offset_h,
                                                                                    offset_l)
            label = 'Background (Smart)'

        elif method == "1x U4-Tougaard":
            background_filtered = BackgroundCalculations.calculate_tougaard_background(x_values_filtered,
                                                                                       y_values_filtered,
                                                                                       sheet_name,
                                                                                       window)
            label = 'Background (Tougaard)'
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
        else:
            background_filtered = BackgroundCalculations.calculate_smart_background(x_values_filtered,
                                                                                    y_values_filtered, offset_h,
                                                                                    offset_l)
            label = 'Background (Smart)'
            # raise ValueError(f"Unknown background method: {method}")

        new_background = np.array(window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
        new_background[mask] = background_filtered
        return new_background, label

    def _update_background_data(self, window, sheet_name, x_values, background, method, offset_h, offset_l):
        """Helper method to update the background data in window.Data."""
        window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = background.tolist()
        window.background = background
        window.Data['Core levels'][sheet_name]['Background'].update({
            'Bkg Y': background.tolist(),
            'Bkg Type': method,
            'Bkg Low': min(x_values),
            'Bkg High': max(x_values),
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
            self.ax.legend(loc='upper left')

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


    def clear_background_only_OLD(self, window):
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name in window.Data['Core levels']:
            # Reset background to raw data
            window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = window.Data['Core levels'][sheet_name][
                'Raw Data']
            window.background = np.array(window.Data['Core levels'][sheet_name]['Raw Data'])

            # Reset background parameters
            window.Data['Core levels'][sheet_name]['Background'].update({
                'Bkg Type': '',
                'Bkg Low': '',
                'Bkg High': '',
                'Bkg Offset Low': '',
                'Bkg Offset High': ''
            })

            # Reset vlines
            window.vline1 = None
            window.vline2 = None
            window.show_hide_vlines()

            window.offset_l = 0
            window.offset_h = 0
            window.fitting_window.offset_l_text.SetValue('0')
            window.fitting_window.offset_h_text.SetValue('0')

            # Redraw the plot
            self.clear_and_replot(window)

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
# --------------------- HISTORY --------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------

