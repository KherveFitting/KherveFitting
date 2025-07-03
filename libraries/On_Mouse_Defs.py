import wx
import numpy as np
import os
from libraries.Sheet_Operations import on_sheet_selected
from libraries.Save import save_state
import platform



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
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            self.window.offset_l = min(calculated_offset, 0)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l
                            self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                        else:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            self.window.offset_h = min(calculated_offset, 0)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset High'] = self.window.offset_h
                            self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')
                    else:
                        raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                        if vline2_x == low_be_x:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            self.window.offset_l = min(calculated_offset, 0)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l
                            self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                        else:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            self.window.offset_h = min(calculated_offset, 0)
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
        self.window.shift_key_pressed = False
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
            if platform.system() == 'Darwin':  # Darwin is macOS
                return

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
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        self.window.offset_l = min(calculated_offset, 0)
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset Low'] = self.window.offset_l
                        self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                    else:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        self.window.offset_h = min(calculated_offset, 0)
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset High'] = self.window.offset_h
                        self.window.fitting_window.offset_h_text.SetValue(f'{self.window.offset_h:.1f}')
                else:
                    raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                    if vline2_x == low_be_x:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        self.window.offset_l = min(calculated_offset, 0)
                        self.window.Data['Core levels'][sheet_name]['Background'][
                            'Bkg Offset Low'] = self.window.offset_l
                        self.window.fitting_window.offset_l_text.SetValue(f'{self.window.offset_l:.1f}')
                    else:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        self.window.offset_h = min(calculated_offset, 0)
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
            # if hasattr(self.window, 'motion_notify_id'):
            #     self.window.canvas.mpl_disconnect(self.window.motion_notify_id)
            #     delattr(self.window, 'motion_notify_id')
            # if hasattr(self.window, 'button_release_id'):
            #     self.window.canvas.mpl_disconnect(self.window.button_release_id)
            #     delattr(self.window, 'button_release_id')

            # Use the correct variable names to disconnect events
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            # Reset the moving vline to None
            self.window.moving_vline = None

            # # Rest of your existing code for updating background data...
            # sheet_name = self.window.sheet_combobox.GetValue()
            # if sheet_name in self.window.Data['Core levels']:
            #     core_level_data = self.window.Data['Core levels'][sheet_name]
            #     if 'Background' in core_level_data:
            #         bg_low = core_level_data['Background']['Bkg Low']
            #         bg_high = core_level_data['Background']['Bkg High']
            #         # Update background if needed
            #         if self.window.background_method == "Multi-Regions Smart":
            #             self.window.plot_manager.plot_background(self.window)
            #
            # self.window.moving_vline = None

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

        # Existing items
        copy_item = menu.Append(wx.ID_ANY, "Copy Peak Table")
        paste_item = menu.Append(wx.ID_ANY, "Paste Peak Table")

        menu.AppendSeparator()

        # New peak operations
        if row % 2 == 0:  # Parameter row - can delete this peak
            peak_index = row // 2
            peak_letter = self.window.peak_params_grid.GetCellValue(row, 0)
            delete_item = menu.Append(wx.ID_ANY, f"Delete Peak {peak_letter}")
        else:
            delete_item = None

        # Add peak submenu
        add_submenu = wx.Menu()
        models = [
            "GL (Area)", "SGL (Area)", "LA (Area, σ/γ, γ)", "Voigt (Area, L/G, σ)",
            "LA (Area, σ, γ)", "DS*G (A, σ, γ, S)", "Voigt (Area, L/G, σ, S)",
            "ExpGauss.(Area, σ, γ)", "Voigt (Area, σ, γ)", "LA*G (Area, σ/γ, γ)",
            "Pseudo-Voigt (Area)", "GL (Height)", "SGL (Height)", "DS (A, σ, γ)"
        ]

        add_items = []
        for model in models:
            add_items.append(add_submenu.Append(wx.ID_ANY, model))

        menu.AppendSubMenu(add_submenu, "Add Peak")

        menu.AppendSeparator()

        # Existing propagate options
        propagate_text = "Propagate to column"
        propagate_diff_text = "Propagate difference to column"

        if col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1:
            param_row = row - 1
            peak_letter = self.window.peak_params_grid.GetCellValue(param_row, 0)
            col_names = {
                2: "Positions", 3: "Heights", 4: "FWHMs", 5: "L/G ratios",
                6: "Areas", 7: "Sigmas", 8: "Gammas", 9: "Skews"
            }
            param_name = col_names.get(col, "values")
            if col == 2:
                propagate_text = f"Constraint all Current {param_name} to {peak_letter}"
            elif col == 3:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
            elif col == 4:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
                propagate_diff_text = (f"Constraint all Current {param_name} to {peak_letter}")
            elif col == 5:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
            elif col == 6:
                propagate_text = f"Constraint all Current {param_name} to {peak_letter}"
            elif col == 7:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
            elif col == 8:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
            elif col == 9:
                propagate_text = f"Constraint all {param_name}: {peak_letter}*1"
            # if col == 4:
            #     propagate_diff_text = f"OR Constraint all {param_name}: {peak_letter} + (current values - {peak_letter})"

        propagate_item = menu.Append(wx.ID_ANY, propagate_text)
        propagate_diff_item = None
        if col == 4 and row % 2 == 1:
            propagate_diff_item = menu.Append(wx.ID_ANY, propagate_diff_text)

        # Enable/disable items
        clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')
        has_clipboard_data = os.path.exists(clipboard_file)
        has_rows = self.window.peak_params_grid.GetNumberRows() > 0

        copy_item.Enable(has_rows)
        paste_item.Enable(has_clipboard_data)
        if delete_item:
            delete_item.Enable(has_rows)
        propagate_item.Enable(col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1)

        # Bind events
        from libraries.Save import copy_all_peak_parameters, paste_all_peak_parameters
        from libraries.Utilities import propagate_constraint, propagate_fwhm_difference

        self.window.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self.window), copy_item)
        self.window.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self.window), paste_item)

        if delete_item:
            self.window.Bind(wx.EVT_MENU, lambda evt: self.delete_peak_at_index(peak_index), delete_item)

        for i, add_item in enumerate(add_items):
            model = models[i]
            self.window.Bind(wx.EVT_MENU, lambda evt, m=model, r=row: self.add_peak_with_model(m, r), add_item)

        self.window.Bind(wx.EVT_MENU, lambda evt: propagate_constraint(self.window, row, col), propagate_item)
        if propagate_diff_item:
            self.window.Bind(wx.EVT_MENU, lambda evt: propagate_fwhm_difference(self.window, row, col),
                             propagate_diff_item)

        self.window.peak_params_grid.PopupMenu(menu, event.GetPosition())
        menu.Destroy()

    def delete_peak_at_index(self, peak_index):
        """Delete a peak and renumber remaining peaks"""
        save_state(self.window)

        sheet_name = self.window.sheet_combobox.GetValue()

        # Remove rows from grid first
        row = peak_index * 2
        self.window.peak_params_grid.DeleteRows(row, 2)
        self.window.peak_count -= 1

        # Update Data structure
        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][sheet_name]:
            if 'Fitting' in self.window.Data['Core levels'][sheet_name] and 'Peaks' in \
                    self.window.Data['Core levels'][sheet_name]['Fitting']:
                peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                peak_keys = list(peaks.keys())

                if peak_index < len(peak_keys):
                    # Store all peak data with their original labels in order
                    all_peak_data = []
                    for key in peak_keys:
                        all_peak_data.append((key, peaks[key]))  # Keep original key/label

                    # Remove the deleted peak
                    all_peak_data.pop(peak_index)

                    # Clear and rebuild peaks dictionary with original labels
                    peaks.clear()
                    for i, (original_label, peak_data) in enumerate(all_peak_data):
                        # Update constraints to reference correct peak letters
                        if 'Constraints' in peak_data:
                            for constraint_key, constraint_value in peak_data['Constraints'].items():
                                if isinstance(constraint_value, str):
                                    peak_data['Constraints'][constraint_key] = self.update_constraint_references(
                                        constraint_value, peak_index)

                        # Store with original label
                        peaks[original_label] = peak_data

        # Update grid letters only (A, B, C, D...)
        for i in range(self.window.peak_params_grid.GetNumberRows() // 2):
            new_letter = chr(65 + i)
            self.window.peak_params_grid.SetCellValue(i * 2, 0, new_letter)

        # Update all constraint references in grid
        for row in range(1, self.window.peak_params_grid.GetNumberRows(), 2):  # Constraint rows only
            for col in range(2, 10):  # Constraint columns
                constraint = self.window.peak_params_grid.GetCellValue(row, col)
                updated_constraint = self.update_constraint_references(constraint, peak_index)
                self.window.peak_params_grid.SetCellValue(row, col, updated_constraint)


        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.window, sheet_name)
        self.window.update_ratios()

        self.window.clear_and_replot()

    def update_constraints_after_deletion(self, deleted_index):
        """Update peak letter references in constraints after deletion"""
        for row in range(1, self.window.peak_params_grid.GetNumberRows(), 2):  # Constraint rows only
            for col in range(2, 10):  # Constraint columns
                constraint = self.window.peak_params_grid.GetCellValue(row, col)
                updated_constraint = self.update_constraint_references(constraint, deleted_index)
                self.window.peak_params_grid.SetCellValue(row, col, updated_constraint)

    def update_constraint_references(self, constraint, deleted_index):
        """Update peak letter references in a constraint string"""
        if not constraint or constraint in ['Fixed', '']:
            return constraint

        import re

        # Pattern to match peak letters (A-P) in constraints
        def replace_letter(match):
            letter = match.group(1)
            letter_index = ord(letter) - 65

            if letter_index > deleted_index:
                # Shift letter down by one (C becomes B, D becomes C, etc.)
                new_letter = chr(65 + letter_index - 1)
                return new_letter
            elif letter_index == deleted_index:
                # Reference to deleted peak - convert to default range
                return "1:1000"
            else:
                # Letters before deleted index stay the same
                return letter

        # Replace all peak letter references
        pattern = r'([A-P])(?=[+\-*/]|$)'
        updated = re.sub(pattern, replace_letter, constraint)

        # Handle cases where constraint becomes invalid
        if updated.startswith('1:1000'):
            # If it was just a letter reference, make it a proper range
            if ':' not in constraint:
                return "1:1000"

        return updated

    def add_peak_with_model(self, model_name, row=None):
        """Add a peak with specified model at specified position"""
        save_state(self.window)

        if row is None:
            row = 0

        # Calculate insert position
        insert_index = row // 2 if row % 2 == 0 else (row + 1) // 2

        sheet_name = self.window.sheet_combobox.GetValue()

        # Check if background exists
        if self.window.bg_min_energy is None or self.window.bg_max_energy is None:
            self.window.show_popup_message2("No Background", "Please create a background first.")
            return

        # Get current peak data
        peaks_data = []
        if sheet_name in self.window.Data['Core levels'] and 'Fitting' in self.window.Data['Core levels'][
            sheet_name] and 'Peaks' in self.window.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            for key, data in peaks.items():
                peaks_data.append((key, data))

        # Set fitting method temporarily
        old_method = self.window.selected_fitting_method
        self.window.selected_fitting_method = model_name

        # Create new peak data manually (similar to add_peak_params but without grid operations)
        self.window.peak_count += 1

        # Calculate peak position from residuals
        if len(peaks_data) == 0:
            residual = self.window.y_values - np.array(
                self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
            peak_y = residual[np.argmax(residual)]
            peak_x = self.window.x_values[np.argmax(residual)]
        else:
            residual = self.window.plot_manager.update_overall_fit_and_residuals(self.window)
            if residual is not None:
                peak_y = residual.max()
                peak_x = self.window.x_values[np.argmax(residual)]
            else:
                peak_y = self.window.y_values.max()
                peak_x = self.window.x_values[np.argmax(self.window.y_values)]

        # Create new peak data based on model
        letter_id = chr(64 + self.window.peak_count)
        new_peak_key = f"{sheet_name} p{self.window.peak_count}"

        # Get constraints range
        x_values = self.window.Data['Core levels'][sheet_name]['B.E.']
        position_constraint = f"{min(x_values):.2f}:{max(x_values):.2f}"

        # Create peak data based on model type
        if model_name in ["ExpGauss.(Area, σ, γ)"]:
            new_peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': round(peak_y * 1.6 * 1.064, 1),
                'Sigma': 0.3,
                'Gamma': 1.2,
                'Skew': 0.64,
                'Fitting Model': model_name,
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
                    'Sigma': "0.01:1",
                    'Gamma': "0.01:3",
                    'Skew': "0.01:2"
                }
            }
        elif model_name in ["LA (Area, σ/γ, γ)", "LA*G (Area, σ/γ, γ)"]:
            new_peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 50,
                'Area': round(peak_y * 1.6 * 1.064, 1),
                'Sigma': 2.7,
                'Gamma': 2.7,
                'Skew': 0.64,
                'Fitting Model': model_name,
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
                    'Sigma': "0.01:10" if "LA*G" not in model_name else "0.01:4",
                    'Gamma': "0.01:10" if "LA*G" not in model_name else "0.01:4",
                    'Skew': "0.01:2"
                }
            }
        elif model_name in ["Voigt (Area, L/G, σ, S)"]:
            new_peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': round(peak_y * 1.6 * 1.064, 1),
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.01,
                'Fitting Model': model_name,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "15:85",
                    'Area': '1:1e7',
                    'Sigma': "0.2:1.5",
                    'Gamma': "0.2:1.5",
                    'Skew': "0.01:0.7"
                }
            }
        elif model_name in ["DS (A, σ, γ)", "DS*G (A, σ, γ, S)"]:
            new_peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.0,
                'L/G': 20,
                'Area': round(peak_y * 1.6 * 1.064, 1),
                'Sigma': 0.5,
                'Gamma': 0.5 if "DS*G" in model_name else 0.0,
                'Skew': 0.0,
                'Fitting Model': model_name,
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
                    'Gamma': "0.1:1.5" if "DS*G" in model_name else "-0.1:1.5",
                    'Skew': "0:0.2" if "DS*G" in model_name else "-0.2:0.2"
                }
            }
        else:
            # Default for GL, SGL, Pseudo-Voigt, etc.
            new_peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': round(peak_y * 1.6 * 1.064, 1),
                'Sigma': 1.0,
                'Gamma': 0.15,
                'Skew': 0.64,
                'Fitting Model': model_name,
                'Bkg Type': self.window.background_method,
                'Bkg Low': self.window.bg_min_energy,
                'Bkg High': self.window.bg_max_energy,
                'Bkg Offset Low': self.window.offset_l,
                'Bkg Offset High': self.window.offset_h,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "5:80",
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:2"
                }
            }

        # Insert at correct position
        peaks_data.insert(insert_index, (new_peak_key, new_peak_data))

        # Update all constraint references for peaks after insert point
        for i, (key, data) in enumerate(peaks_data):
            if i > insert_index and 'Constraints' in data:
                for constraint_key, constraint_value in data['Constraints'].items():
                    if isinstance(constraint_value, str):
                        data['Constraints'][constraint_key] = self.shift_constraint_letters_after_insert(
                            constraint_value, insert_index)

        # Update Data structure
        if 'Fitting' not in self.window.Data['Core levels'][sheet_name]:
            self.window.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.window.Data['Core levels'][sheet_name]['Fitting']:
            self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        new_peaks = {}
        for key, data in peaks_data:
            new_peaks[key] = data
        self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks

        # Rebuild grid
        self.rebuild_grid_after_insert(peaks_data, insert_index)

        # Restore original method
        self.window.selected_fitting_method = old_method
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.window, sheet_name)
        self.window.update_ratios()
        self.window.clear_and_replot()

    def rebuild_grid_from_data(self, sheet_name):
        """Rebuild grid from Data structure with all formatting"""
        # Clear grid
        if self.window.peak_params_grid.GetNumberRows() > 0:
            self.window.peak_params_grid.DeleteRows(0, self.window.peak_params_grid.GetNumberRows())

        # Use the existing sheet selection logic which rebuilds everything correctly
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.window, sheet_name)

    def shift_constraint_letters_after_insert(self, constraint, insert_index):
        """Shift peak letter references in constraints after inserting a peak"""
        if not constraint or constraint in ['Fixed', '']:
            return constraint

        import re

        def replace_letter(match):
            letter = match.group(1)
            letter_index = ord(letter) - 65

            if letter_index >= insert_index:
                # Shift letter up by one (B becomes C, C becomes D, etc.)
                new_letter = chr(65 + letter_index + 1)
                return new_letter
            else:
                # Letters before insert index stay the same
                return letter

        # Replace all peak letter references
        pattern = r'([A-P])(?=[+\-*/]|$|#)'
        updated = re.sub(pattern, replace_letter, constraint)

        return updated

    def rebuild_grid_after_insert(self, peaks_data, insert_index):
        """Rebuild the entire grid with correct peak order and IDs"""
        # Clear existing grid
        if self.window.peak_params_grid.GetNumberRows() > 0:
            self.window.peak_params_grid.DeleteRows(0, self.window.peak_params_grid.GetNumberRows())

        # Add rows for all peaks
        num_peaks = len(peaks_data)
        self.window.peak_params_grid.AppendRows(num_peaks * 2)

        # Populate grid with reordered data
        for i, (key, data) in enumerate(peaks_data):
            row = i * 2

            # Set peak ID letter
            letter_id = chr(65 + i)
            self.window.peak_params_grid.SetCellValue(row, 0, letter_id)
            self.window.peak_params_grid.SetReadOnly(row, 0)

            # Set peak data
            self.window.peak_params_grid.SetCellValue(row, 1, key)
            self.window.peak_params_grid.SetCellValue(row, 2, f"{data.get('Position', 0):.2f}")
            self.window.peak_params_grid.SetCellValue(row, 3, f"{data.get('Height', 1000):.2f}")
            self.window.peak_params_grid.SetCellValue(row, 4, f"{data.get('FWHM', 1.6):.2f}")
            self.window.peak_params_grid.SetCellValue(row, 5, f"{data.get('L/G', 20):.2f}")
            self.window.peak_params_grid.SetCellValue(row, 6, f"{data.get('Area', 1000):.2f}")
            self.window.peak_params_grid.SetCellValue(row, 7, f"{data.get('Sigma', 1.0):.3f}")
            self.window.peak_params_grid.SetCellValue(row, 8, f"{data.get('Gamma', 0.5):.3f}")
            self.window.peak_params_grid.SetCellValue(row, 9, f"{data.get('Skew', 0.1):.3f}")
            self.window.peak_params_grid.SetCellValue(row, 13, data.get('Fitting Model', 'GL (Area)'))
            self.window.peak_params_grid.SetCellValue(row, 14, data.get('Bkg Type', ''))
            self.window.peak_params_grid.SetCellValue(row, 15, str(data.get('Bkg Low', '')))
            self.window.peak_params_grid.SetCellValue(row, 16, str(data.get('Bkg High', '')))
            self.window.peak_params_grid.SetCellValue(row, 17, str(data.get('Bkg Offset Low', '')))
            self.window.peak_params_grid.SetCellValue(row, 18, str(data.get('Bkg Offset High', '')))

            # Set constraints
            if 'Constraints' in data:
                constraints = data['Constraints']
                constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
                for col_idx, constraint_key in enumerate(constraint_keys, 2):
                    constraint_value = constraints.get(constraint_key, '')
                    self.window.peak_params_grid.SetCellValue(row + 1, col_idx, str(constraint_value))

            # Apply colors based on fitting model (same as add_peak_params)
            model = data.get('Fitting Model', 'GL (Area)')

            # Set background colors for constraint row
            for col in range(self.window.peak_params_grid.GetNumberCols()):
                self.window.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200, 245, 228))
                self.window.peak_params_grid.SetCellBackgroundColour(row, col, wx.WHITE)

            # Set text colors based on model (copy from add_peak_params)
            for col in [10, 11, 12]:
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(27, 140, 60))
            for col in [0, 1, 2]:
                self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))

            # Apply model-specific coloring
            if model == "Voigt (Area, L/G, σ)":
                for col in [3, 4, 8]:
                    self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(128, 128, 128))
                    self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
                for col in [5, 6, 7]:
                    self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                    self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
                for col in [9]:
                    self.window.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                    self.window.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
            # Add other model-specific coloring here following the pattern from add_peak_params

        # Apply choice editors and formatting
        self.window.set_model_choice_editors(self.window)
        self.window.peak_params_grid.ForceRefresh()

    def create_peak_data_for_model(self, model_name, peak_x, peak_y, sheet_name):
        """Create peak data structure for specific model"""
        # Get background range
        bg_low = self.window.bg_min_energy
        bg_high = self.window.bg_max_energy
        position_constraint = f"{bg_low:.2f},{bg_high:.2f}"

        # Base peak data
        peak_data = {
            'Position': peak_x,
            'Height': peak_y,
            'FWHM': 1.6,
            'L/G': 20,
            'Area': peak_y * 1.6 * 1.064,
            'Sigma': 1.0,
            'Gamma': 0.5,
            'Skew': 0.1,
            'Fitting Model': model_name,
            'Bkg Type': self.window.background_method,
            'Bkg Low': bg_low,
            'Bkg High': bg_high,
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

        # Model-specific adjustments
        if model_name in ["LA (Area, σ, γ)", "LA (Area, σ/γ, γ)", "LA*G (Area, σ/γ, γ)"]:
            peak_data.update({
                'L/G': 50,
                'Sigma': 2.7,
                'Gamma': 2.7,
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "Fixed",
                    'Sigma': "0.01:10",
                    'Gamma': "0.01:10"
                }
            })
        elif model_name in ["Voigt (Area, L/G, σ, S)"]:
            peak_data.update({
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.01,
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "15:85",
                    'Sigma': "0.2:1.5",
                    'Gamma': "0.2:1.5",
                    'Skew': "0.01:0.7"
                }
            })
        elif model_name in ["DS (A, σ, γ)"]:
            peak_data.update({
                'Sigma': 0.5,
                'Gamma': 0.0,
                'Skew': 0.0,
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "Fixed",
                    'Sigma': "0.3:1.5",
                    'Gamma': "0.1:1.5",
                    'Skew': "-0.2:0.2"
                }
            })
        elif model_name in ["DS*G (A, σ, γ, S)"]:
            peak_data.update({
                'Sigma': 0.5,
                'Gamma': 0.5,
                'Skew': 0.0,
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "Fixed",
                    'Sigma': "0.3:1.5",
                    'Gamma': "0.1:1.5",
                    'Skew': "0:0.2"
                }
            })
        elif model_name == "ExpGauss.(Area, σ, γ)":
            peak_data.update({
                'Sigma': 0.3,
                'Gamma': 1.2,
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "Fixed",
                    'Sigma': "0.01:1",
                    'Gamma': "0.01:3",
                    'Skew': "0.01:2"
                }
            })
        elif model_name in ["Voigt (Area, σ, γ)"]:
            peak_data.update({
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "Fixed"
                }
            })
        elif model_name in ["GL (Area)", "SGL (Area)", "Pseudo-Voigt (Area)"]:
            peak_data.update({
                'Constraints': {
                    **peak_data['Constraints'],
                    'L/G': "5:80"
                }
            })

        return peak_data




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