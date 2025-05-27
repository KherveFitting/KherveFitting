import wx
import numpy as np
# from libraries.Save import save_state
# from libraries.Sheet_Operations import on_sheet_selected

class PeakManipulation:
    def __init__(self, main_window):
        self.window = main_window

    def highlight_selected_peak(self):
        if self.window.selected_peak_index is not None:
            num_peaks = self.window.peak_params_grid.GetNumberRows() // 2

            if self.window.selected_peak_index >= num_peaks:
                print(f"Warning: selected_peak_index ({self.window.selected_peak_index}) is out of range. Max index: {num_peaks - 1}")
                self.window.selected_peak_index = None
                return

            for i in range(num_peaks):
                row = i * 2
                is_selected = (i == self.window.selected_peak_index)
                self.window.peak_params_grid.SetCellBackgroundColour(row, 0, wx.LIGHT_GREY if is_selected else wx.WHITE)
                self.window.peak_params_grid.SetCellBackgroundColour(row + 1, 0, wx.LIGHT_GREY if is_selected else wx.WHITE)

            row = self.window.selected_peak_index * 2

            if row < self.window.peak_params_grid.GetNumberRows():
                peak_label = self.window.peak_params_grid.GetCellValue(row, 1)
                x_str = self.window.peak_params_grid.GetCellValue(row, 2)
                y_str = self.window.peak_params_grid.GetCellValue(row, 3)
                pos = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 2), 0)
                grid_fwhm = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 4), 1.6)
                lg_ratio = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 5), 20.0)
                area = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 6), 100.0)
                sigma = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 7), 0.0)
                gamma = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 8), 0.0)
                skew = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 9), 0.0)
                model = self.window.peak_params_grid.GetCellValue(row, 13)

                if x_str and y_str:
                    try:
                        x = float(x_str)
                        y = float(y_str)
                        y_with_bg = y + self.window.background[np.argmin(np.abs(self.window.x_values - x))]

                        from libraries.Peak_Functions import PeakFunctions
                        actual_fwhm = PeakFunctions.calculate_actual_fwhm(
                            self.window.x_values, x, y, grid_fwhm, lg_ratio, area, sigma, gamma, skew, model
                        )

                        self.window.actual_fwhms[self.window.selected_peak_index] = actual_fwhm

                        self.remove_cross_from_peak()
                        if hasattr(self.window, 'peak_letter') and self.window.peak_letter:
                            self.window.peak_letter.remove()
                            self.window.peak_letter = None
                            del self.window.peak_letter
                        if hasattr(self.window, 'peak_info') and self.window.peak_info:
                            self.window.peak_info.remove()
                            self.window.peak_info = None
                            del self.window.peak_info
                        if hasattr(self.window, 'peak_letter_t') and self.window.peak_letter_t:
                            self.window.peak_letter_t.remove()
                            self.window.peak_letter_t = None
                            del self.window.peak_letter_t
                        if hasattr(self.window, 'peak_info_t') and self.window.peak_info_t:
                            self.window.peak_info_t.remove()
                            self.window.peak_info_t = None
                            del self.window.peak_info_t

                        self.window.cross, = self.window.ax.plot(x, y_with_bg, 'bx', markersize=15, markerfacecolor='none', picker=5, linewidth=3)

                        peak_letter = chr(65 + self.window.selected_peak_index)
                        peak_info = (f'Model: {model}\n'
                                     f'Position: {pos} eV\n'
                                     f'FWHM meas.: {actual_fwhm:.3f} eV\n'
                                     f'Area: {area:.1f} CPS\n\n'
                                     f'\u00BF Change width ?\n'
                                     f'SHIFT + wheel button')

                        max_y = self.window.ax.get_ylim()[1]
                        y_offset = max_y * 0.02
                        self.window.peak_letter_t = self.window.ax.text(x, y_with_bg + y_offset, peak_letter, ha='center', va='bottom', fontsize=12)
                        self.window.peak_info_t = self.window.ax.text(x - actual_fwhm / 2, y_with_bg + y_offset, peak_info, ha='left', va='top', fontsize=8, color='grey')

                        self.window.peak_params_grid.ClearSelection()
                        self.window.peak_params_grid.SelectRow(row, addToSelected=False)
                        self.window.peak_params_grid.Refresh()
                        self.window.canvas.draw_idle()

                        sheet_name = self.window.sheet_combobox.GetValue()
                        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in self.window.Data['Core levels'][sheet_name]['Fitting']:
                            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                            if self.window.selected_peak_index < len(peaks):
                                old_label = list(peaks.keys())[self.window.selected_peak_index]
                                if old_label != peak_label:
                                    peaks[peak_label] = peaks.pop(old_label)
                            else:
                                print(f"Warning: selected_peak_index ({self.window.selected_peak_index}) is out of range for peaks in Data structure")

                        self.window.canvas.mpl_connect('motion_notify_event', self.on_cross_drag)
                        self.window.canvas.mpl_connect('button_release_event', self.on_cross_release)
                    except ValueError as e:
                        print(f"Warning: Invalid data for selected peak. Cannot highlight. Error: {e}")
                else:
                    print(f"Warning: Empty data for selected peak. Cannot highlight.")
            else:
                print(f"Warning: Row {row} does not exist in peak_params_grid")

            self.window.peak_params_grid.Refresh()
        else:
            print("No peak selected (selected_peak_index is None)")

    def change_selected_peak(self, direction):
        num_peaks = self.window.peak_params_grid.GetNumberRows() // 2

        if self.window.selected_peak_index is None:
            self.window.selected_peak_index = 0 if direction > 0 else num_peaks - 1
        else:
            self.window.selected_peak_index = (self.window.selected_peak_index + direction) % num_peaks

        self.remove_cross_from_peak()

        if self.window.peak_fitting_tab_selected:
            row = self.window.selected_peak_index * 2
            self.window.initial_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
            self.window.initial_x = float(self.window.peak_params_grid.GetCellValue(row, 2))
            self.highlight_selected_peak()

        self.window.canvas.draw_idle()

    def deselect_all_peaks(self):
        self.window.selected_peak_index = None
        self.remove_cross_from_peak()

        self.window.peak_params_grid.ClearSelection()

        for row in range(self.window.results_grid.GetNumberRows()):
            self.window.results_grid.SetCellValue(row, 7, '0')

        self.window.update_peak_plot(None, None, remove_old_peaks=True)

        self.window.peak_params_grid.ForceRefresh()
        self.window.results_grid.ForceRefresh()

    def get_peak_index_from_position(self, x, y):
        x_display, y_display = self.window.ax.transData.transform((x, y))
        num_peaks = self.window.peak_params_grid.GetNumberRows() // 2

        for i in range(num_peaks):
            row = i * 2
            if self.window.peak_params_grid.IsInSelection(row, 0):
                x_peak = float(self.window.peak_params_grid.GetCellValue(row, 2))
                y_peak = float(self.window.peak_params_grid.GetCellValue(row, 3))
                bkg_y = self.window.background[np.argmin(np.abs(self.window.x_values - x_peak))]
                y_peak += bkg_y

                x_peak_display, y_peak_display = self.window.ax.transData.transform((x_peak, y_peak))
                distance = np.sqrt((x_display - x_peak_display) ** 2 + (y_display - y_peak_display) ** 2)

                if distance < 100:
                    return i
        return None

    def update_peak(self, peak_index, new_x, new_height, area=None):
        row = peak_index * 2
        sheet_name = self.window.sheet_combobox.GetValue()
        peak_label = self.window.peak_params_grid.GetCellValue(row, 1)
        fitting_model = self.window.peak_params_grid.GetCellValue(row, 13)

        self.window.peak_params_grid.SetCellValue(row, 2, f"{new_x:.2f}")

        if "LA" in fitting_model and area is not None:
            self.window.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")
            self.window.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")
        else:
            self.window.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")

        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in self.window.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            if peak_label in peaks:
                peaks[peak_label]['Position'] = new_x
                if "LA" in fitting_model and area is not None:
                    peaks[peak_label]['Area'] = area
                    peaks[peak_label]['Height'] = new_height
                else:
                    peaks[peak_label]['Height'] = new_height

        if not "LA" in fitting_model:
            self.window.recalculate_peak_area(peak_index)

        grid_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
        lg_ratio = float(self.window.peak_params_grid.GetCellValue(row, 5))
        area = float(self.window.peak_params_grid.GetCellValue(row, 6))
        sigma = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 7), 2.7)
        gamma = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 8), 2.7)
        skew = self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 9), 0.1)

        from libraries.Peak_Functions import PeakFunctions
        actual_fwhm = PeakFunctions.calculate_actual_fwhm(
            self.window.x_values, new_x, new_height, grid_fwhm, lg_ratio, area, sigma, gamma, skew, fitting_model
        )

        self.window.actual_fwhms[peak_index] = actual_fwhm

    def update_peak_grid(self, index, x, y):
        row = index * 2
        self.window.peak_params_grid.SetCellValue(row, 2, f"{x:.2f}")
        self.window.peak_params_grid.SetCellValue(row, 3, f"{y:.2f}")
        self.window.peak_params_grid.ForceRefresh()

    def update_fwhm_grid(self, index, fwhm):
        row = index * 2
        self.window.peak_params_grid.SetCellValue(row, 4, f"{fwhm:.2f}")

    def on_cross_drag(self, event):
        if event.inaxes and self.window.selected_peak_index is not None:
            row = self.window.selected_peak_index * 2

            if row >= self.window.peak_params_grid.GetNumberRows():
                self.window.selected_peak_index = None
                return

            fitting_model = self.window.peak_params_grid.GetCellValue(row, 13)

            if event.button == 1:
                try:
                    if event.key == 'shift':
                        new_fwhm = self.window.update_peak_fwhm(event.xdata)
                        if new_fwhm is not None:
                            self.window.update_linked_fwhm_recursive(self.window.selected_peak_index, new_fwhm)

                    elif self.window.is_mouse_on_peak(event):
                        closest_index = np.argmin(np.abs(self.window.x_values - event.xdata))
                        bkg_y = self.window.background[closest_index]
                        new_x = event.xdata
                        new_height = max(event.ydata - bkg_y, 0)

                        if "LA" in fitting_model:
                            fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                            sigma = float(self.window.peak_params_grid.GetCellValue(row, 7))
                            gamma = float(self.window.peak_params_grid.GetCellValue(row, 8))
                            skew = float(self.window.peak_params_grid.GetCellValue(row, 9)) if "LA*G" in fitting_model else None
                            new_area = self.window.calculate_peak_area(fitting_model, new_height, fwhm, 0, sigma, gamma, skew)

                            self.update_peak(self.window.selected_peak_index, new_x, new_height, new_area)
                            self.window.update_linked_peaks_recursive(self.window.selected_peak_index, new_x, new_height, new_area)
                        else:
                            self.update_peak(self.window.selected_peak_index, new_x, new_height)
                            self.window.update_linked_peaks_recursive(self.window.selected_peak_index, new_x, new_height)

                    self.window.update_ratios()
                    self.window.clear_and_replot()
                    self.window.plot_manager.add_cross_to_peak(self.window, self.window.selected_peak_index, skip_fwhm_calc=True)
                    self.window.canvas.draw_idle()

                except Exception as e:
                    print(f"Error during cross drag: {e}")

    def on_cross_release(self, event, save_state_func=None):
        if save_state_func:
            save_state_func(self.window)
        if event.inaxes and self.window.selected_peak_index is not None:
            row = self.window.selected_peak_index * 2
            fitting_model = self.window.peak_params_grid.GetCellValue(row, 13)
            peak_label = self.window.peak_params_grid.GetCellValue(row, 1)
            sheet_name = self.window.sheet_combobox.GetValue()

            x = event.xdata
            y = event.ydata
            bkg_y = self.window.background[np.argmin(np.abs(self.window.x_values - x))]

            if event.button == 1:
                if event.key == 'shift':
                    new_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                    self.window.update_linked_fwhm_recursive(self.window.selected_peak_index, new_fwhm)
                else:
                    y = max(y - bkg_y, 0)

                    if "LA" in fitting_model:
                        current_area = float(self.window.peak_params_grid.GetCellValue(row, 6))
                        self.update_peak(self.window.selected_peak_index, x, y, current_area)
                        self.window.update_linked_peaks_recursive(self.window.selected_peak_index, x, y, current_area)
                    else:
                        self.update_peak(self.window.selected_peak_index, x, y)
                        self.window.update_linked_peaks_recursive(self.window.selected_peak_index, x, y)

            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            self.remove_cross_from_peak()
            self.window.clear_and_replot()
            from libraries.Sheet_Operations import on_sheet_selected
            wx.CallAfter(self.highlight_selected_peak)

        self.window.peak_fitting_grid.refresh_peak_params_grid_release()

    def remove_cross_from_peak(self):
        if hasattr(self.window, 'cross'):
            if self.window.cross in self.window.ax.lines:
                self.window.cross.remove()
            del self.window.cross

        if hasattr(self.window, 'peak_letter_t') and self.window.peak_letter_t:
            self.window.peak_letter_t.remove()
            self.window.peak_letter_t = None
            del self.window.peak_letter_t
        if hasattr(self.window, 'peak_info_t') and self.window.peak_info_t:
            self.window.peak_info_t.remove()
            self.window.peak_info_t = None
            del self.window.peak_info_t
        if hasattr(self.window, 'peak_letter') and self.window.peak_letter:
            self.window.peak_letter.remove()
            self.window.peak_letter = None
            del self.window.peak_letter
        if hasattr(self.window, 'peak_info') and self.window.peak_info:
            self.window.peak_info.remove()
            self.window.peak_info = None
            del self.window.peak_info

        self.window.canvas.mpl_disconnect('motion_notify_event')
        self.window.canvas.mpl_disconnect('button_release_event')