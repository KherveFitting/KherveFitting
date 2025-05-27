import wx
import numpy as np
import os
from libraries.Sheet_Operations import on_sheet_selected
from libraries.Save import save_state



class MouseEventHandler:
    def __init__(self, window):
        self.window = window

    def on_mouse_move(self, event):
        if event.inaxes:
            x, y = event.xdata, event.ydata
            if self.window.energy_scale == 'KE':
                self.window.SetStatusText(f"KE: {x:.1f} eV, I: {int(y)} CPS", 1)
                self.window.current_energy_value = x
            else:
                self.window.SetStatusText(f"BE: {x:.1f} eV, I: {int(y)} CPS", 1)
                self.window.current_energy_value = x

    def on_click(self, event):
        if event.inaxes:
            x_click = event.xdata
            if event.button == 1 and event.key == 'shift' and self.window.background_tab_selected:
                self.window.motion_notify_id = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                self.window.button_release_id = self.window.canvas.mpl_connect('button_release_event', self.on_release)

                x_click = event.xdata
                sheet_name = self.window.sheet_combobox.GetValue()
                if self.window.vline1 is not None and self.window.vline2 is not None:
                    vline1_x = self.window.vline1.get_xdata()[0]
                    vline2_x = self.window.vline2.get_xdata()[0]

                    low_be_x = min(vline1_x, vline2_x)
                    high_be_x = max(vline1_x, vline2_x)

                    dist1 = abs(x_click - vline1_x)
                    dist2 = abs(x_click - vline2_x)

                    if dist1 < dist2:
                        raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline1_x))]
                        if vline1_x == low_be_x:
                            self.window.offset_l = event.ydata - raw_y
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l
                            self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                        else:
                            self.window.offset_h = event.ydata - raw_y
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset High'] = self.window.offset_h
                            self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')
                    else:
                        raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                        if vline2_x == low_be_x:
                            self.window.offset_l = event.ydata - raw_y
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l
                            self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                        else:
                            self.window.offset_h = event.ydata - raw_y
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset High'] = self.window.offset_h
                            self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')
                    self.window.plot_manager.plot_background(self.window)
                    return
            elif event.button == 1:
                if event.key == 'shift':
                    if self.window.peak_fitting_tab_selected and self.window.selected_peak_index is not None:
                        row = self.window.selected_peak_index * 2
                        self.window.initial_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                        self.window.initial_x = event.xdata
                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                self.window.peak_manipulation.on_cross_drag)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                 self.window.peak_manipulation.on_cross_release)
                elif self.window.background_tab_selected:
                    self.window.peak_manipulation.deselect_all_peaks()
                    sheet_name = self.window.sheet_combobox.GetValue()
                    if sheet_name in self.window.Data['Core levels']:
                        core_level_data = self.window.Data['Core levels'][sheet_name]
                        if self.window.background_method == "Multi-Regions Smart":
                            if self.window.vline1 is not None and self.window.vline2 is not None:
                                vline1_x = self.window.vline1.get_xdata()[0]
                                vline2_x = self.window.vline2.get_xdata()[0]

                                dist1 = abs(x_click - vline1_x)
                                dist2 = abs(x_click - vline2_x)

                                if dist1 < dist2 and dist1 < self.window.some_threshold:
                                    self.window.moving_vline = self.window.vline1
                                elif dist2 < self.window.some_threshold:
                                    self.window.moving_vline = self.window.vline2
                                else:
                                    self.window.moving_vline = None

                                if self.window.moving_vline is not None:
                                    self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                            self.on_motion)
                                    self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                             self.on_release)
                                    return

                        if self.window.vline1 is None:
                            self.window.vline1 = self.window.ax.axvline(x_click, color='r', linestyle='--')
                            core_level_data['Background']['Bkg Low'] = float(x_click)
                        elif self.window.vline2 is None and abs(
                                x_click - core_level_data['Background']['Bkg Low']) > self.window.some_threshold:
                            self.window.vline2 = self.window.ax.axvline(x_click, color='r', linestyle='--')
                            core_level_data['Background']['Bkg High'] = float(x_click)
                            core_level_data['Background']['Bkg Low'], core_level_data['Background'][
                                'Bkg High'] = sorted([
                                core_level_data['Background']['Bkg Low'],
                                core_level_data['Background']['Bkg High']
                            ])
                        else:
                            self.window.moving_vline = self.window.vline1 if self.window.vline2 is None or abs(
                                x_click - core_level_data['Background']['Bkg Low']) < abs(
                                x_click - core_level_data['Background']['Bkg High']) else self.window.vline2
                            self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                    self.on_motion)
                            self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                     self.on_release)
                elif self.window.noise_tab_selected:
                    if self.window.vline3 is None:
                        self.window.vline3 = self.window.ax.axvline(x_click, color='b', linestyle='--')
                        self.window.noise_min_energy = float(x_click)
                    elif self.window.vline4 is None and abs(
                            x_click - self.window.noise_min_energy) > self.window.some_threshold:
                        self.window.vline4 = self.window.ax.axvline(x_click, color='b', linestyle='--')
                        self.window.noise_max_energy = float(x_click)
                        self.window.noise_min_energy, self.window.noise_max_energy = sorted(
                            [self.window.noise_min_energy, self.window.noise_max_energy])
                    else:
                        self.window.moving_vline = self.window.vline3 if self.window.vline4 is None or abs(
                            x_click - self.window.noise_min_energy) < abs(
                            x_click - self.window.noise_max_energy) else self.window.vline4
                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                 self.on_release)
                elif self.window.peak_fitting_tab_selected:
                    peak_index = self.window.peak_manipulation.get_peak_index_from_position(event.xdata, event.ydata)
                    if peak_index is not None:
                        self.window.selected_peak_index = peak_index
                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                self.window.peak_manipulation.on_cross_drag)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                 self.window.peak_manipulation.on_cross_release)
                        self.window.peak_manipulation.highlight_selected_peak()
                    else:
                        self.window.peak_manipulation.deselect_all_peaks()
                else:
                    self.window.peak_manipulation.deselect_all_peaks()

            self.window.show_hide_vlines()
            self.window.canvas.draw()

    def on_mouse_wheel(self, event):
        shift_currently_pressed = event.key == 'shift'

        if shift_currently_pressed:
            self.window.shift_key_pressed = True
        else:
            self.window.shift_key_pressed = False

        if self.window.shift_key_pressed and self.window.selected_peak_index is not None and self.window.peak_fitting_tab_selected:
            save_state(self.window)
            delta = 0.05 if event.step > 0 else -0.05
            row = self.window.selected_peak_index * 2
            fitting_model = self.window.peak_params_grid.GetCellValue(row, 13)

            if fitting_model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                 "Voigt (Area, L/G, \u03c3, S)"]:
                current_sigma = float(self.window.peak_params_grid.GetCellValue(row, 7))
                new_sigma = max(current_sigma + delta, 0.2)

                self.window.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")

                lg_ratio = float(self.window.peak_params_grid.GetCellValue(row, 5))
                new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                self.window.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
            else:
                current_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                new_fwhm = max(current_fwhm + delta, 0.3)
                self.window.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")

            self.window.recalculate_peak_area(self.window.selected_peak_index)
            self.window.update_linked_fwhm_recursive(self.window.selected_peak_index,
                                                     new_sigma if fitting_model.startswith("Voigt") else new_fwhm)
            self.window.clear_and_replot()
            self.window.peak_manipulation.highlight_selected_peak()

        elif not self.window.shift_key_pressed:
            current_index = self.window.sheet_combobox.GetSelection()
            num_sheets = self.window.sheet_combobox.GetCount()

            if event.step > 0:
                new_index = (current_index - 1) % num_sheets
            else:
                new_index = (current_index + 1) % num_sheets

            if num_sheets > 0:
                self.window.sheet_combobox.SetSelection(new_index)
                new_sheet = self.window.sheet_combobox.GetString(new_index)
                on_sheet_selected(self.window, new_sheet)

    def on_motion(self, event):
        if event.button == 1 and event.key == 'shift' and self.window.background_tab_selected:
            x_click = event.xdata
            sheet_name = self.window.sheet_combobox.GetValue()
            if self.window.vline1 is not None and self.window.vline2 is not None:
                vline1_x = self.window.vline1.get_xdata()[0]
                vline2_x = self.window.vline2.get_xdata()[0]

                low_be_x = min(vline1_x, vline2_x)
                high_be_x = max(vline1_x, vline2_x)

                dist1 = abs(x_click - vline1_x)
                dist2 = abs(x_click - vline2_x)

                if dist1 < dist2:
                    raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline1_x))]
                    if vline1_x == low_be_x:
                        self.window.offset_l = event.ydata - raw_y
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset Low'] = self.window.offset_l
                        self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                    else:
                        self.window.offset_h = event.ydata - raw_y
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset High'] = self.window.offset_h
                        self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')
                else:
                    raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                    if vline2_x == low_be_x:
                        self.window.offset_l = event.ydata - raw_y
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset Low'] = self.window.offset_l
                        self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                    else:
                        self.window.offset_h = event.ydata - raw_y
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset High'] = self.window.offset_h
                        self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')

                self.window.plot_manager.plot_background(self.window)
        elif event.inaxes and self.window.moving_vline is not None:
            x_click = event.xdata
            self.window.moving_vline.set_xdata([x_click])

            sheet_name = self.window.sheet_combobox.GetValue()
            if sheet_name in self.window.Data['Core levels']:
                core_level_data = self.window.Data['Core levels'][sheet_name]

                if self.window.moving_vline in [self.window.vline1, self.window.vline2]:
                    if self.window.moving_vline == self.window.vline1:
                        core_level_data['Background']['Bkg Low'] = float(x_click)
                    else:
                        core_level_data['Background']['Bkg High'] = float(x_click)

                    bkg_low = core_level_data['Background']['Bkg Low']
                    bkg_high = core_level_data['Background']['Bkg High']
                    core_level_data['Background']['Bkg Low'] = min(bkg_low, bkg_high)
                    core_level_data['Background']['Bkg High'] = max(bkg_low, bkg_high)

                elif self.window.moving_vline in [self.window.vline3, self.window.vline4]:
                    if self.window.moving_vline == self.window.vline3:
                        self.window.noise_min_energy = float(x_click)
                    else:
                        self.window.noise_max_energy = float(x_click)

                    self.window.noise_min_energy, self.window.noise_max_energy = sorted(
                        [self.window.noise_min_energy, self.window.noise_max_energy])

            self.window.canvas.draw_idle()

    def on_release(self, event):
        if self.window.moving_vline is not None:
            if hasattr(self.window, 'motion_notify_id'):
                self.window.canvas.mpl_disconnect(self.window.motion_notify_id)
                delattr(self.window, 'motion_notify_id')
            if hasattr(self.window, 'button_release_id'):
                self.window.canvas.mpl_disconnect(self.window.button_release_id)
                delattr(self.window, 'button_release_id')

            self.window.moving_vline = None

            sheet_name = self.window.sheet_combobox.GetValue()
            if sheet_name in self.window.Data['Core levels']:
                core_level_data = self.window.Data['Core levels'][sheet_name]
                if 'Background' in core_level_data:
                    bg_low = core_level_data['Background'].get('Bkg Low')
                    bg_high = core_level_data['Background'].get('Bkg High')
                    if bg_low is not None and bg_high is not None:
                        core_level_data['Background']['Bkg Low'] = min(bg_low, bg_high)
                        core_level_data['Background']['Bkg High'] = max(bg_low, bg_high)

        if self.window.selected_peak_index is not None:
            row = self.window.selected_peak_index * 2
            peak_x = float(self.window.peak_params_grid.GetCellValue(row, 2))
            peak_y = float(self.window.peak_params_grid.GetCellValue(row, 3))
            self.window.update_peak_plot(peak_x, peak_y, remove_old_peaks=False)

            sheet_name = self.window.sheet_combobox.GetValue()
            if sheet_name in self.window.Data['Core levels']:
                core_level_data = self.window.Data['Core levels'][sheet_name]
                if 'Fitting' not in core_level_data:
                    core_level_data['Fitting'] = {}
                if 'Peaks' not in core_level_data['Fitting']:
                    core_level_data['Fitting']['Peaks'] = {}

                peak_label = self.window.peak_params_grid.GetCellValue(row, 1)

                core_level_data['Fitting']['Peaks'][peak_label] = {
                    'Position': peak_x,
                    'Height': peak_y,
                    'FWHM': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 4), 1.6),
                    'L/G': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 5), 20.0),
                    'Area': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 6), 0.0),
                    'Sigma': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 7), 0.5),
                    'Gamma': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 8), 0.5),
                    'Skew': self.window.try_float(self.window.peak_params_grid.GetCellValue(row, 9), 0.1)
                }

        self.window.selected_peak_index = None
        self.window.canvas.draw_idle()

    def on_right_click(self, event):
        if event.button == 3:
            import os
            import tempfile
            from libraries.Save import copy_all_peak_parameters, paste_all_peak_parameters, copy_core_level, \
                paste_core_level

            menu = wx.Menu()
            zoom_in = menu.Append(-1, "Zoom In")
            zoom_out = menu.Append(-1, "Zoom Out")
            drag = menu.Append(-1, "Drag")

            menu.AppendSeparator()

            copy = menu.Append(-1, "Copy Core Level")
            paste = menu.Append(-1, "Paste Core Level")

            clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_clipboard.json')
            has_clipboard_data = os.path.exists(clipboard_file)

            menu.AppendSeparator()

            copy_peak_table = menu.Append(-1, "Copy Peak Table")
            paste_peak_table = menu.Append(-1, "Paste Peak Table")
            peak_clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')

            has_peak_clipboard_data = os.path.exists(peak_clipboard_file)
            has_rows = self.window.peak_params_grid.GetNumberRows() > 0

            paste.Enable(has_clipboard_data)
            paste_peak_table.Enable(has_peak_clipboard_data)
            copy_peak_table.Enable(has_rows)

            self.window.Bind(wx.EVT_MENU, self.window.on_zoom_in_tool, zoom_in)
            self.window.Bind(wx.EVT_MENU, self.window.on_zoom_out, zoom_out)
            self.window.Bind(wx.EVT_MENU, self.window.on_drag_tool, drag)
            self.window.Bind(wx.EVT_MENU, lambda evt: copy_core_level(self.window), copy)
            self.window.Bind(wx.EVT_MENU, lambda evt: paste_core_level(self.window), paste)
            self.window.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self.window), copy_peak_table)
            self.window.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self.window), paste_peak_table)

            self.window.PopupMenu(menu)
            menu.Destroy()

    def on_peak_params_right_click(self, event):
        import tempfile
        row = event.GetRow()
        col = event.GetCol()

        menu = wx.Menu()
        copy_item = menu.Append(wx.ID_ANY, "Copy Peak Table")
        paste_item = menu.Append(wx.ID_ANY, "Paste Peak Table")

        # Add separator and propagate option
        menu.AppendSeparator()

        # Create dynamic propagate text
        propagate_text = "Propagate to column"
        propagate_diff_text = "Propagate difference to column"  # New option

        if col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1:
            # Get the peak letter from the parameter row
            param_row = row - 1
            peak_letter = self.window.peak_params_grid.GetCellValue(param_row, 0)

            # Map column to parameter name
            col_names = {
                2: "Positions", 3: "Heights", 4: "FWHMs", 5: "L/G ratios",
                6: "Areas", 7: "Sigmas", 8: "Gammas", 9: "Skews"
            }

            param_name = col_names.get(col, "values")
            propagate_text = f"Propagate {param_name} from {peak_letter}"

            # Add difference option for FWHM
            if col == 4:  # FWHM column
                propagate_diff_text = f"Propagate FWHM others - {peak_letter}"

        propagate_item = menu.Append(wx.ID_ANY, propagate_text)

        # Only add difference option for FWHM constraint rows
        propagate_diff_item = None
        if col == 4 and row % 2 == 1:  # FWHM constraint row
            propagate_diff_item = menu.Append(wx.ID_ANY, propagate_diff_text)

        clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')
        has_clipboard_data = os.path.exists(clipboard_file)
        has_rows = self.window.peak_params_grid.GetNumberRows() > 0

        copy_item.Enable(has_rows)
        paste_item.Enable(has_clipboard_data)
        propagate_item.Enable(col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1)

        if propagate_diff_item:
            propagate_diff_item.Enable(True)

        from libraries.Save import copy_all_peak_parameters, paste_all_peak_parameters
        from libraries.Utilities import propagate_constraint, propagate_fwhm_difference

        self.window.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self.window), copy_item)
        self.window.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self.window), paste_item)
        self.window.Bind(wx.EVT_MENU, lambda evt: propagate_constraint(self.window, row, col), propagate_item)

        if propagate_diff_item:
            self.window.Bind(wx.EVT_MENU, lambda evt: propagate_fwhm_difference(self.window, row, col),
                             propagate_diff_item)

        self.window.peak_params_grid.PopupMenu(menu, event.GetPosition())
        menu.Destroy()




def setup_mouse_handlers(window):
    """Set up mouse event handlers for the window"""
    mouse_handler = MouseEventHandler(window)

    window.canvas.mpl_connect("button_press_event", mouse_handler.on_click)
    window.canvas.mpl_connect('motion_notify_event', mouse_handler.on_mouse_move)
    window.canvas.mpl_connect('scroll_event', mouse_handler.on_mouse_wheel)
    window.canvas.mpl_connect('button_press_event', mouse_handler.on_right_click)

    # Add peak params grid right click handler
    window.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, mouse_handler.on_peak_params_right_click)

    return mouse_handler