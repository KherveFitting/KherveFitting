# libraries/On_Key_defs.py

import wx
import numpy as np


class KeyEventHandlers:
    """Class containing all key event handlers for KherveFitting"""

    def __init__(self, main_frame):
        self.main_frame = main_frame

    def on_key_press(self, event):
        """Handle matplotlib key press events"""
        if event.key == 'shift':
            self.main_frame.shift_key_pressed = True
            return

        if self.main_frame.selected_peak_index is not None:
            num_peaks = self.main_frame.peak_params_grid.GetNumberRows() // 2  # Assuming each peak uses two rows

            if event.key == 'tab':
                if not self.main_frame.peak_fitting_tab_selected:
                    self.main_frame.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
                else:
                    self.main_frame.peak_manipulation.change_selected_peak(1)  # Move to next peak
                return  # Prevent event from propagating
            elif event.key == 'q':
                pass

            self.main_frame.peak_manipulation.highlight_selected_peak()
            self.main_frame.clear_and_replot()
            self.main_frame.canvas.draw_idle()

        # Connect to key release event
        self.main_frame.canvas.mpl_connect('key_release_event', self.on_key_release)

    def on_key_release(self, event):
        """Handle matplotlib key release events"""
        if event.key == 'shift':
            self.main_frame.shift_key_pressed = False

            # Store FWHM for currently selected peak when shift is released
            if self.main_frame.selected_peak_index is not None:
                row = self.main_frame.selected_peak_index * 2
                peak_x = float(self.main_frame.peak_params_grid.GetCellValue(row, 2))
                peak_y = float(self.main_frame.peak_params_grid.GetCellValue(row, 3))
                grid_fwhm = float(self.main_frame.peak_params_grid.GetCellValue(row, 4))
                lg_ratio = float(self.main_frame.peak_params_grid.GetCellValue(row, 5))
                area = float(self.main_frame.peak_params_grid.GetCellValue(row, 6))
                sigma = self.main_frame.try_float(self.main_frame.peak_params_grid.GetCellValue(row, 7), 0.0)
                gamma = self.main_frame.try_float(self.main_frame.peak_params_grid.GetCellValue(row, 8), 0.0)
                skew = self.main_frame.try_float(self.main_frame.peak_params_grid.GetCellValue(row, 9), 0.0)
                model = self.main_frame.peak_params_grid.GetCellValue(row, 13)

                from libraries.Peak_Functions import PeakFunctions
                actual_fwhm = PeakFunctions.calculate_actual_fwhm(
                    self.main_frame.x_values, peak_x, peak_y, grid_fwhm, lg_ratio, area, sigma, gamma, skew, model
                )

                self.main_frame.actual_fwhms[self.main_frame.selected_peak_index] = actual_fwhm

    def on_key_press_global(self, event):
        """Handle global wxPython key press events"""
        keycode = event.GetKeyCode()

        # Handle Alt + Arrow keys for peak manipulation
        if event.AltDown() and self.main_frame.selected_peak_index is not None:
            if event.ShiftDown() and keycode in [wx.WXK_LEFT, wx.WXK_RIGHT]:
                self._handle_alt_shift_arrow_keys(event, keycode)
                return
            elif keycode in [wx.WXK_LEFT, wx.WXK_RIGHT]:
                self._handle_alt_arrow_keys(event, keycode)
                return
            elif keycode in [wx.WXK_UP, wx.WXK_DOWN]:
                self._handle_alt_up_down_keys(event, keycode)
                return

        # Handle Ctrl key combinations
        elif event.ControlDown():
            if self._handle_ctrl_key_combinations(event, keycode):
                return

        # Handle Shift + Arrow keys
        if event.ShiftDown() and keycode in [wx.WXK_LEFT, wx.WXK_RIGHT]:
            self._handle_shift_arrow_keys(event, keycode)
            return

        # Handle other special keys
        if self._handle_special_keys(event, keycode):
            return

        event.Skip()  # Let other key events propagate normally

    def _handle_alt_shift_arrow_keys(self, event, keycode):
        """Handle Alt+Shift+Arrow key combinations for peak width adjustment"""
        from libraries.Save import save_state
        save_state(self.main_frame)

        row = self.main_frame.selected_peak_index * 2
        model = self.main_frame.peak_params_grid.GetCellValue(row, 13)
        delta = 0.05 if keycode == wx.WXK_RIGHT else -0.05
        current_fwhm = float(self.main_frame.peak_params_grid.GetCellValue(row, 4))
        new_fwhm = current_fwhm

        if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
            current_sigma = float(self.main_frame.peak_params_grid.GetCellValue(row, 7))
            lg_ratio = float(self.main_frame.peak_params_grid.GetCellValue(row, 5))
            new_sigma = max(current_sigma + delta, 0.4)
            new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
            self.main_frame.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")
            self.main_frame.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
            if model == "Voigt (Area, L/G, \u03c3, S)":
                skew = float(self.main_frame.peak_params_grid.GetCellValue(row, 9))
        elif model in ["DS (A, \u03c3, \u03b3)"]:
            current_sigma = float(self.main_frame.peak_params_grid.GetCellValue(row, 7))
            current_gamma = float(self.main_frame.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.main_frame.peak_params_grid.GetCellValue(row, 9))

            new_sigma = max(current_sigma + delta, 0.4)
            # gamma is independent in DS model, no need to recalculate

            self.main_frame.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")
        elif model == "DS*G (A, \u03c3, \u03b3, S)":
            current_sigma = float(self.main_frame.peak_params_grid.GetCellValue(row, 7))
            current_gamma = float(self.main_frame.peak_params_grid.GetCellValue(row, 8))
            current_skew = float(self.main_frame.peak_params_grid.GetCellValue(row, 9))

            # Adjust Gaussian width
            new_sigma = max(current_sigma + delta, 0.2)
            self.main_frame.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")

            # Can also adjust gamma slightly
            new_gamma = max(current_gamma + delta * 0.5, 0.1)
            self.main_frame.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
        elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
            current_gamma = float(self.main_frame.peak_params_grid.GetCellValue(row, 8))
            new_gamma = max(current_gamma + delta, 0.2)
            self.main_frame.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
        else:
            new_fwhm = max(current_fwhm + delta, 0.3)
            self.main_frame.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")

        self.main_frame.recalculate_peak_area(self.main_frame.selected_peak_index)
        self.main_frame.update_linked_fwhm_recursive(self.main_frame.selected_peak_index, new_fwhm)
        self.main_frame.clear_and_replot()
        self.main_frame.peak_manipulation.highlight_selected_peak()

    def _handle_alt_arrow_keys(self, event, keycode):
        """Handle Alt+Arrow key combinations for peak position adjustment"""
        from libraries.Save import save_state
        save_state(self.main_frame)

        row = self.main_frame.selected_peak_index * 2
        current_position = float(self.main_frame.peak_params_grid.GetCellValue(row, 2))

        # Move 0.1 eV left or right
        delta = 0.1 if keycode == wx.WXK_LEFT else -0.1
        new_position = current_position + delta

        # Update peak position
        self.main_frame.peak_params_grid.SetCellValue(row, 2, f"{new_position:.2f}")

        # Update linked peaks if any
        self.main_frame.update_linked_peaks_recursive(self.main_frame.selected_peak_index, new_position,
                                                      float(self.main_frame.peak_params_grid.GetCellValue(row, 3)))

        # Refresh display
        self.main_frame.clear_and_replot()
        self.main_frame.plot_manager.add_cross_to_peak(self.main_frame, self.main_frame.selected_peak_index,
                                                       skip_fwhm_calc=False)

    def _handle_alt_up_down_keys(self, event, keycode):
        """Handle Alt+Up/Down key combinations for peak height adjustment"""
        from libraries.Save import save_state
        save_state(self.main_frame)

        row = self.main_frame.selected_peak_index * 2
        current_height = float(self.main_frame.peak_params_grid.GetCellValue(row, 3))
        intensity_factor = 0.05

        delta = intensity_factor * current_height * (1 if keycode == wx.WXK_UP else -1)
        new_height = max(0, current_height + delta)

        self.main_frame.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")

        # Recalculate area after height change
        self.main_frame.recalculate_peak_area(self.main_frame.selected_peak_index)

        self.main_frame.update_linked_peaks_recursive(self.main_frame.selected_peak_index,
                                                      float(self.main_frame.peak_params_grid.GetCellValue(row, 2)),
                                                      new_height)

        self.main_frame.clear_and_replot()
        self.main_frame.plot_manager.add_cross_to_peak(self.main_frame, self.main_frame.selected_peak_index,
                                                       skip_fwhm_calc=False)

    def _handle_ctrl_key_combinations(self, event, keycode):
        """Handle Ctrl key combinations"""
        if keycode == ord('B'):
            self.main_frame.toggle_energy_scale()
            self.main_frame.toggle_energy_scale()
            self.main_frame.clear_and_replot()
            return True
        elif keycode == ord('Z'):
            from libraries.Save import undo
            undo(self.main_frame)
            return True
        elif keycode == ord('Y'):
            from libraries.Save import redo
            redo(self.main_frame)
            return True
        elif keycode == ord('S'):
            print("Saving")
            from Functions import on_save
            on_save(self.main_frame)
            return True
        elif keycode == ord('O'):
            print("Opening")
            from libraries.Open import open_xlsx_file
            open_xlsx_file(self.main_frame)
            return True
        elif keycode == ord('Q'):
            print("Quiting")
            from Functions import on_exit
            on_exit(self.main_frame, event)
            return True
        elif keycode == ord('K'):
            print("Shortcuts")
            self.main_frame.show_popup_message2("Keyboard Shortcuts",
                                                "-Tab: Select next peak\n"
                                                "-Q: Select previous peak\n"
                                                "-Ctrl+Minus (-): Zoom out\n"
                                                "-Ctrl+Equal (=): Zoom in\n"
                                                "-Ctrl+Left bracket [: Select previous core level\n"
                                                "-Ctrl+Right bracket ]: Select next core level\n"
                                                "-Ctrl+Up: Increase plot intensity\n"
                                                "-Ctrl+Down: Decrease plot intensity\n"
                                                "-Ctrl+Left: Move plot to High BE\n"
                                                "-Ctrl+Right: Move plot to Low BE\n"
                                                "-SHIFT+Left: Decrease High BE\n"
                                                "-SHIFT+Right: Increase High BE\n"
                                                "-Ctrl+Z: Undo up to 30 events\n"
                                                "-Ctrl+Y: Redo\n"
                                                "-Ctrl+S: Save. Only works on the grid and not on the figure canvas\n"
                                                "-Ctrl+P: Open peak fitting window\n"
                                                "-Ctrl+A: Open Area window\n"
                                                "-Ctrl+K: Show Keyboard shortcut\n"
                                                "-Alt+Up: Increase peak intensity\n"
                                                "-Alt+Down: Decrease peak intensity\n"
                                                "-Alt+Left: Move peak to High BE\n"
                                                "-Alt+Right: Move peak to Low BE\n"
                                                )
            return True
        elif keycode == ord('A'):
            print("Opening Area Window")
            self.main_frame.on_open_background_window()
            return True
        elif keycode == ord('P'):
            print("Opening Fitting Window")
            self.main_frame.on_open_fitting_window()
            return True
        elif keycode in [ord('['), ord(']'), ord('9'), ord('0')]:
            self._handle_sheet_navigation(keycode)
            return True
        elif keycode in [ord('-'), ord('=')]:
            self._handle_zoom_keys(keycode)
            return True
        elif keycode in [wx.WXK_LEFT, wx.WXK_RIGHT]:
            self._handle_ctrl_arrow_keys(keycode)
            return True
        elif keycode in [wx.WXK_UP, wx.WXK_DOWN]:
            self._handle_ctrl_up_down_keys(keycode)
            return True

        return False

    def _handle_shift_arrow_keys(self, event, keycode):
        """Handle Shift+Arrow key combinations"""
        if keycode == wx.WXK_LEFT:
            self.main_frame.adjust_plot_limits('high_be', 'increase')
        elif keycode == wx.WXK_RIGHT:
            self.main_frame.adjust_plot_limits('high_be', 'decrease')

    def _handle_special_keys(self, event, keycode):
        """Handle special key combinations"""
        if keycode == wx.WXK_TAB:
            import time
            current_time = time.time()
            # Only show popup if 10+ seconds have passed since last one
            if not self.main_frame.peak_fitting_tab_selected and (current_time - self.main_frame.last_popup_time > 10):
                self.main_frame.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
                self.main_frame.last_popup_time = current_time
            elif self.main_frame.peak_fitting_tab_selected:
                self.main_frame.peak_manipulation.change_selected_peak(1)  # Move to next peak
            return True
        elif keycode == ord('Q'):
            if not self.main_frame.peak_fitting_tab_selected:
                self.main_frame.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
            else:
                self.main_frame.peak_manipulation.change_selected_peak(-1)  # Move to previous peak
            return True

        return False

    def _handle_sheet_navigation(self, keycode):
        """Handle sheet navigation with Ctrl+brackets or Ctrl+9/0"""
        current_index = self.main_frame.sheet_combobox.GetSelection()
        num_sheets = self.main_frame.sheet_combobox.GetCount()

        if keycode == ord('[') or keycode == ord('9'):
            new_index = (current_index - 1) % num_sheets
        else:
            new_index = (current_index + 1) % num_sheets

        self.main_frame.sheet_combobox.SetSelection(new_index)
        new_sheet = self.main_frame.sheet_combobox.GetString(new_index)

        # Call on_sheet_selected with the new sheet name
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.main_frame, new_sheet)
        from libraries.Save import save_state
        save_state(self.main_frame)

    def _handle_zoom_keys(self, keycode):
        """Handle zoom in/out with Ctrl+- and Ctrl+="""
        sheet_name = self.main_frame.sheet_combobox.GetValue()
        limits = self.main_frame.plot_config.get_plot_limits(self.main_frame, sheet_name)

        zoom_factor = 0.2
        if keycode == ord('-'):  # Zoom out
            limits['Xmin'] -= zoom_factor
            limits['Xmax'] += zoom_factor
        else:  # Zoom in
            limits['Xmin'] += zoom_factor
            limits['Xmax'] -= zoom_factor

        # Update the plot limits
        self.main_frame.plot_config.update_plot_limits(self.main_frame, sheet_name,
                                                       x_min=limits['Xmin'],
                                                       x_max=limits['Xmax'])

        # Update the plot
        self.main_frame.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis

        # Update subplot limits if it exists
        if hasattr(self.main_frame, 'residuals_subplot') and self.main_frame.residuals_subplot:
            self.main_frame.residuals_subplot.set_xlim(limits['Xmax'], limits['Xmin'])

        # After zooming, update residuals
        self.main_frame.plot_manager.update_overall_fit_and_residuals(self.main_frame)

        self.main_frame.canvas.draw_idle()

    def _handle_ctrl_arrow_keys(self, keycode):
        """Handle Ctrl+Arrow keys for plot movement"""
        sheet_name = self.main_frame.sheet_combobox.GetValue()
        limits = self.main_frame.plot_config.get_plot_limits(self.main_frame, sheet_name)
        move_factor = 0.1

        if keycode == wx.WXK_LEFT:
            limits['Xmin'] -= move_factor
            limits['Xmax'] -= move_factor
        else:  # Right key
            limits['Xmin'] += move_factor
            limits['Xmax'] += move_factor

        # Update the plot limits
        self.main_frame.plot_config.update_plot_limits(self.main_frame, sheet_name,
                                                       x_min=limits['Xmin'],
                                                       x_max=limits['Xmax'])

        # Update the plot
        self.main_frame.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
        self.main_frame.plot_manager.update_overall_fit_and_residuals(self.main_frame)

        self.main_frame.canvas.draw_idle()

    def _handle_ctrl_up_down_keys(self, keycode):
        """Handle Ctrl+Up/Down keys for intensity adjustment"""
        sheet_name = self.main_frame.sheet_combobox.GetValue()
        limits = self.main_frame.plot_config.get_plot_limits(self.main_frame, sheet_name)
        intensity_factor = 0.05
        max_intensity = max(self.main_frame.y_values)

        if keycode == wx.WXK_DOWN:  # Decrease intensity
            limits['Ymax'] = max(limits['Ymax'] - intensity_factor * max_intensity, limits['Ymin'])
        else:  # Increase intensity
            limits['Ymax'] += intensity_factor * max_intensity

        # Update the plot limits
        self.main_frame.plot_config.update_plot_limits(self.main_frame, sheet_name, y_max=limits['Ymax'])

        # Update the plot
        self.main_frame.ax.set_ylim(limits['Ymin'], limits['Ymax'])

        # Check RSD visibility
        if hasattr(self.main_frame.plot_manager,
                   'residuals_state') and self.main_frame.plot_manager.residuals_state == 1:
            residual_height = 1.07 * max(self.main_frame.y_values)
            if residual_height > limits['Ymax']:
                if hasattr(self.main_frame.plot_manager, 'rsd_text') and self.main_frame.plot_manager.rsd_text:
                    self.main_frame.plot_manager.rsd_text.remove()
                    self.main_frame.plot_manager.rsd_text = None
            else:
                if hasattr(self.main_frame.plot_manager, 'rsd_text') and self.main_frame.plot_manager.rsd_text is None:
                    from libraries.Peak_Functions import PeakFunctions
                    rsd = PeakFunctions.calculate_rsd(self.main_frame.y_values, self.main_frame.background)
                    if rsd is not None:
                        x_min = self.main_frame.ax.get_xlim()[1] + 0.4
                        self.main_frame.plot_manager.rsd_text = self.main_frame.ax.text(x_min, residual_height,
                                                                                        f'RSD: {rsd:.2f}',
                                                                                        horizontalalignment='right',
                                                                                        verticalalignment='center',
                                                                                        fontsize=9,
                                                                                        color=self.main_frame.plot_manager.residual_color,
                                                                                        alpha=self.main_frame.plot_manager.residual_alpha + 0.2,
                                                                                        bbox=dict(facecolor='white',
                                                                                                  edgecolor='none'))

        self.main_frame.canvas.draw_idle()


def setup_key_handlers(main_frame):
    """Setup key event handlers for the main frame"""
    key_handlers = KeyEventHandlers(main_frame)

    # Store reference to handlers in main frame
    main_frame.key_handlers = key_handlers

    # Set up the event bindings
    main_frame.canvas.mpl_connect('key_press_event', key_handlers.on_key_press)
    main_frame.Bind(wx.EVT_CHAR_HOOK, key_handlers.on_key_press_global)

    return key_handlers