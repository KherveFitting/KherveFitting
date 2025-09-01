import wx
import numpy as np
import os
from libraries.Sheet_Operations import on_sheet_selected
from libraries.FileMenu.Save import save_state
import platform



class MouseEventHandler:
    def __init__(self, window):
        self.window = window
        # Add these new variables for CTRL+drag functionality
        self.ctrl_drag_active = False
        self.vline_gap = 0.0
        self.ctrl_drag_reference_pos = 0.0

    def on_mouse_move(self, event):
        if event.inaxes:
            x, y = event.xdata, event.ydata
            if self.window.energy_scale == 'KE':
                self.window.SetStatusText(f"KE: {x:.3f} eV, I: {y:.3f} CPS", 1)
                self.window.current_energy_value = x
            else:
                self.window.SetStatusText(f"BE: {x:.3f} eV, I: {y:.3f} CPS", 1)
                self.window.current_energy_value = x

    def insert_cross_core_constraint(self, constraint_ref, row, col):
        """Insert a cross-core-level constraint into the specified cell"""
        # Save state for undo
        from libraries.FileMenu.Save import save_state
        save_state(self.window)

        # Determine if we're on a parameter row or constraint row
        if row % 2 == 0:  # Parameter row - insert into constraint row below
            constraint_row = row + 1
            parameter_row = row
        else:  # Constraint row - use this row
            constraint_row = row
            parameter_row = row - 1

        # Parse the constraint reference (e.g., "C1s_A")
        core_level_name, peak_letter = constraint_ref.split('_')

        # Get the referenced peak data
        if core_level_name in self.window.Data['Core levels']:
            core_level_data = self.window.Data['Core levels'][core_level_name]

            if ('Fitting' in core_level_data and
                    'Peaks' in core_level_data['Fitting']):

                peaks = core_level_data['Fitting']['Peaks']
                peak_keys = list(peaks.keys())
                peak_index = ord(peak_letter) - ord('A')

                if peak_index < len(peak_keys):
                    peak_key = peak_keys[peak_index]
                    ref_peak_data = peaks[peak_key]

                    if col == 2:  # Position constraint
                        # Get current position and reference position
                        current_pos = float(self.window.peak_params_grid.GetCellValue(parameter_row, col))
                        ref_pos = float(ref_peak_data.get('Position', current_pos))

                        # Calculate difference
                        difference = current_pos - ref_pos

                        # Create constraint string
                        if difference >= 0:
                            constraint_value = f"{constraint_ref}+{difference:.2f}#0.1"
                        else:
                            constraint_value = f"{constraint_ref}{difference:.2f}#0.1"  # difference is already negative

                    elif col == 6:  # Area constraint
                        # Get current area and reference area for ratio calculation
                        current_area = float(self.window.peak_params_grid.GetCellValue(parameter_row, col))
                        ref_area = float(ref_peak_data.get('Area', current_area))

                        # Calculate ratio
                        ratio = (current_area / ref_area) if ref_area != 0 else 1

                        if ratio != 1:
                            constraint_value = f"{constraint_ref}*{ratio:.2f}#0.01"
                        else:
                            constraint_value = f"{constraint_ref}*1"

                    else:  # FWHM, Sigma, Gamma, Height, L/G, Skew
                        # For non-position/area parameters, just add *1
                        constraint_value = f"{constraint_ref}*1"

                    # Insert the constraint into the constraint cell
                    self.window.peak_params_grid.SetCellValue(constraint_row, col, constraint_value)

                    # Update the data structure
                    sheet_name = self.window.sheet_combobox.GetValue()
                    if sheet_name in self.window.Data['Core levels']:
                        fitting_data = self.window.Data['Core levels'][sheet_name].get('Fitting', {})
                        peaks_data = fitting_data.get('Peaks', {})

                        current_peak_index = constraint_row // 2  # Use constraint_row for peak index
                        peak_keys_current = list(peaks_data.keys())

                        if current_peak_index < len(peak_keys_current):
                            current_peak_key = peak_keys_current[current_peak_index]

                            if 'Constraints' not in peaks_data[current_peak_key]:
                                peaks_data[current_peak_key]['Constraints'] = {}

                            constraint_names = {
                                2: 'Position', 3: 'Height', 4: 'FWHM', 5: 'L/G',
                                6: 'Area', 7: 'Sigma', 8: 'Gamma', 9: 'Skew'
                            }
                            constraint_name = constraint_names.get(col)
                            if constraint_name:
                                peaks_data[current_peak_key]['Constraints'][constraint_name] = constraint_value

                    # Refresh the grid
                    self.window.peak_params_grid.ForceRefresh()

    def on_click(self, event):
        if event.inaxes:
            x_click = event.xdata
            if self.window.zoom_mode:
                return  # Exit early, don't process vLine clicks during zoom

            # CTRL+DRAG FUNCTIONALITY - Try multiple CTRL key detection methods
            ctrl_detected = (
                    event.key == 'ctrl' or
                    event.key == 'control' or
                    (event.key and 'ctrl' in str(event.key).lower()) or
                    (event.key and 'control' in str(event.key).lower())
            )

            if (event.button == 1 and ctrl_detected and
                    (self.window.background_tab_selected or
                     (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected))):



                # Check if both vlines exist
                if self.window.vline1 is not None and self.window.vline2 is not None:
                    vline1_x = self.window.vline1.get_xdata()[0]
                    vline2_x = self.window.vline2.get_xdata()[0]

                    # Calculate current gap between vlines (maintain sign/order)
                    self.vline_gap = vline2_x - vline1_x
                    self.ctrl_drag_reference_pos = x_click
                    self.ctrl_drag_active = True

                    # Explicitly prevent normal vline movement
                    self.window.moving_vline = None

                    # Clean up any existing handlers first
                    if hasattr(self.window, 'motion_cid'):
                        self.window.canvas.mpl_disconnect(self.window.motion_cid)
                        delattr(self.window, 'motion_cid')
                    if hasattr(self.window, 'release_cid'):
                        self.window.canvas.mpl_disconnect(self.window.release_cid)
                        delattr(self.window, 'release_cid')

                    # Set up motion and release handlers
                    self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                    self.window.release_cid = self.window.canvas.mpl_connect('button_release_event', self.on_release)

                    return  # CRITICAL: Exit here to prevent any other vline logic
                else:
                    print("CTRL detected but vlines don't exist")

            elif event.button == 1 and event.key == 'shift' and self.window.background_tab_selected:
                # Store current vline positions
                current_vline1_pos = None
                current_vline2_pos = None
                if self.window.vline1 is not None:
                    current_vline1_pos = self.window.vline1.get_xdata()[0]
                if self.window.vline2 is not None:
                    current_vline2_pos = self.window.vline2.get_xdata()[0]

                # Clean up any existing handlers first
                if hasattr(self.window, 'motion_cid'):
                    self.window.canvas.mpl_disconnect(self.window.motion_cid)
                    delattr(self.window, 'motion_cid')
                if hasattr(self.window, 'release_cid'):
                    self.window.canvas.mpl_disconnect(self.window.release_cid)
                    delattr(self.window, 'release_cid')

                self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                self.window.release_cid = self.window.canvas.mpl_connect('button_release_event', self.on_release)

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
                            calculated_offset = min(calculated_offset, 0)

                            # REGION-SPECIFIC UPDATE:
                            self.window.fitting_window.offset_l_text.SetValue(f'{calculated_offset:.1f}')
                            if hasattr(self.window.fitting_window,
                                       'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                                offset_h_value = float(self.window.fitting_window.offset_h_text.GetValue())
                                self.window.fitting_window.update_active_range_offsets(offset_h_value,
                                                                                       calculated_offset)
                            self.window.set_offset_l(calculated_offset)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l

                        else:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            calculated_offset = min(calculated_offset, 0)

                            # REGION-SPECIFIC UPDATE:
                            self.window.fitting_window.offset_h_text.SetValue(f'{calculated_offset:.1f}')
                            if hasattr(self.window.fitting_window,
                                       'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                                offset_l_value = float(self.window.fitting_window.offset_l_text.GetValue())
                                self.window.fitting_window.update_active_range_offsets(calculated_offset,
                                                                                       offset_l_value)
                            self.window.set_offset_h(calculated_offset)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset High'] = self.window.offset_h

                    else:
                        raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                        if vline2_x == low_be_x:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            calculated_offset = min(calculated_offset, 0)

                            # REGION-SPECIFIC UPDATE:
                            self.window.fitting_window.offset_l_text.SetValue(f'{calculated_offset:.1f}')
                            if hasattr(self.window.fitting_window,
                                       'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                                offset_h_value = float(self.window.fitting_window.offset_h_text.GetValue())
                                self.window.fitting_window.update_active_range_offsets(offset_h_value,
                                                                                       calculated_offset)
                            self.window.set_offset_l(calculated_offset)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset Low'] = self.window.offset_l

                        else:
                            calculated_offset = event.ydata - raw_y
                            # Ensure offset cannot be positive
                            calculated_offset = min(calculated_offset, 0)

                            # REGION-SPECIFIC UPDATE:
                            self.window.fitting_window.offset_h_text.SetValue(f'{calculated_offset:.1f}')
                            if hasattr(self.window.fitting_window,
                                       'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                                offset_l_value = float(self.window.fitting_window.offset_l_text.GetValue())
                                self.window.fitting_window.update_active_range_offsets(calculated_offset,
                                                                                       offset_l_value)
                            self.window.set_offset_h(calculated_offset)
                            self.window.Data['Core levels'][sheet_name]['Background'][
                                'Bkg Offset High'] = self.window.offset_h

                    self.window.plot_manager.plot_background(self.window)

                    # Force correct legend update for Multi-Regions Smart background
                    if hasattr(self.window.plot_manager, 'update_legend'):
                        self.window.plot_manager.update_legend(self.window)

                    # Restore vlines after plotting
                    if current_vline1_pos is not None and current_vline2_pos is not None:
                        wx.CallAfter(self.restore_vlines_after_plot, current_vline1_pos, current_vline2_pos)
                    return
            elif event.button == 1:
                if event.key == 'shift':
                    if self.window.peak_fitting_tab_selected and self.window.selected_peak_index is not None:
                        row = self.window.selected_peak_index * 2
                        self.window.initial_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                        self.window.initial_x = event.xdata

                        # Clean up existing handlers
                        if hasattr(self.window, 'motion_cid'):
                            self.window.canvas.mpl_disconnect(self.window.motion_cid)
                        if hasattr(self.window, 'release_cid'):
                            self.window.canvas.mpl_disconnect(self.window.release_cid)

                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                self.window.peak_manipulation.on_cross_drag)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                 self.window.peak_manipulation.on_cross_release)
                        # Check if either fitting screen background tab OR area screen is active for vline interaction
                if (self.window.background_tab_selected or
                        (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected)):
                    # Clean up any existing handlers first
                    if hasattr(self.window, 'motion_cid'):
                        self.window.canvas.mpl_disconnect(self.window.motion_cid)
                        delattr(self.window, 'motion_cid')
                    if hasattr(self.window, 'release_cid'):
                        self.window.canvas.mpl_disconnect(self.window.release_cid)
                        delattr(self.window, 'release_cid')

                    # Reset moving vline state
                    self.window.moving_vline = None

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

                                # Calculate adaptive threshold based on plot width
                                x_range = abs(max(self.window.x_values) - min(self.window.x_values))
                                adaptive_threshold = max(self.window.some_threshold, x_range * 0.02)  # 2% of plot width

                                if dist1 < dist2 and dist1 < adaptive_threshold:
                                    self.window.moving_vline = self.window.vline1
                                elif dist2 < adaptive_threshold:
                                    self.window.moving_vline = self.window.vline2
                                else:
                                    self.window.moving_vline = None

                                if self.window.moving_vline is not None:
                                    self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                            self.on_motion)
                                    self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                             self.on_release)
                                    return
                        else:
                            # Standard background mode - improved vline selection logic
                            if self.window.vline1 is not None and self.window.vline2 is not None:
                                # Both vlines exist
                                vline1_x = self.window.vline1.get_xdata()[0]
                                vline2_x = self.window.vline2.get_xdata()[0]

                                dist1 = abs(x_click - vline1_x)
                                dist2 = abs(x_click - vline2_x)

                                # Calculate adaptive threshold
                                x_range = abs(max(self.window.x_values) - min(self.window.x_values))
                                adaptive_threshold = max(self.window.some_threshold, x_range * 0.02)

                                # Check if click is near either vline
                                if dist1 <= adaptive_threshold or dist2 <= adaptive_threshold:
                                    # Select the closest vline
                                    if dist1 < dist2:
                                        self.window.moving_vline = self.window.vline1
                                    else:
                                        self.window.moving_vline = self.window.vline2

                                    self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                            self.on_motion)
                                    self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                             self.on_release)
                                    return
                                else:
                                    # Click not near any vline - move the closest one
                                    if dist1 < dist2:
                                        self.window.moving_vline = self.window.vline1
                                        # Convert display position back to BE for storage
                                        be_position = self.window.convert_energy_from_display(x_click)
                                        core_level_data['Background']['Bkg Low'] = float(be_position)
                                    else:
                                        self.window.moving_vline = self.window.vline2
                                        # Convert display position back to BE for storage
                                        be_position = self.window.convert_energy_from_display(x_click)
                                        core_level_data['Background']['Bkg High'] = float(be_position)

                                    # Update the vline position immediately
                                    self.window.moving_vline.set_xdata([x_click])

                                    # Sort the background limits
                                    bkg_low = core_level_data['Background']['Bkg Low']
                                    bkg_high = core_level_data['Background']['Bkg High']
                                    core_level_data['Background']['Bkg Low'] = min(bkg_low, bkg_high)
                                    core_level_data['Background']['Bkg High'] = max(bkg_low, bkg_high)

                                    self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                            self.on_motion)
                                    self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                             self.on_release)
                                    self.window.canvas.draw_idle()
                                    return

                        # Create vlines if they don't exist
                        if self.window.vline1 is None:
                            self.window.vline1 = self.window.ax.axvline(x_click, color='r', linestyle='--')
                            # Convert display position back to BE for storage
                            be_position = self.window.convert_energy_from_display(x_click)
                            core_level_data['Background']['Bkg Low'] = float(be_position)
                            self.window.canvas.draw_idle()
                        elif self.window.vline2 is None and abs(
                                x_click - self.window.convert_energy_for_display(
                                    core_level_data['Background']['Bkg Low'])) > self.window.some_threshold:
                            self.window.vline2 = self.window.ax.axvline(x_click, color='r', linestyle='--')
                            # Convert display position back to BE for storage
                            be_position = self.window.convert_energy_from_display(x_click)
                            core_level_data['Background']['Bkg High'] = float(be_position)
                            core_level_data['Background']['Bkg Low'], core_level_data['Background'][
                                'Bkg High'] = sorted([
                                core_level_data['Background']['Bkg Low'],
                                core_level_data['Background']['Bkg High']
                            ])
                            self.window.canvas.draw_idle()
                        else:
                            # Both vlines exist but we're here somehow - select closest one
                            if self.window.vline2 is not None:
                                # Convert stored BE values to display coordinates for distance calculation
                                display_low = self.window.convert_energy_for_display(
                                    core_level_data['Background']['Bkg Low'])
                                display_high = self.window.convert_energy_for_display(
                                    core_level_data['Background']['Bkg High'])
                                dist_to_low = abs(x_click - display_low)
                                dist_to_high = abs(x_click - display_high)

                                if dist_to_low < dist_to_high:
                                    self.window.moving_vline = self.window.vline1
                                else:
                                    self.window.moving_vline = self.window.vline2
                            else:
                                self.window.moving_vline = self.window.vline1

                            self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event',
                                                                                    self.on_motion)
                            self.window.release_cid = self.window.canvas.mpl_connect('button_release_event',
                                                                                     self.on_release)

                elif self.window.noise_tab_selected:
                    # Clean up existing handlers
                    if hasattr(self.window, 'motion_cid'):
                        self.window.canvas.mpl_disconnect(self.window.motion_cid)
                    if hasattr(self.window, 'release_cid'):
                        self.window.canvas.mpl_disconnect(self.window.release_cid)

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
                    # Clean up existing handlers
                    if hasattr(self.window, 'motion_cid'):
                        self.window.canvas.mpl_disconnect(self.window.motion_cid)
                    if hasattr(self.window, 'release_cid'):
                        self.window.canvas.mpl_disconnect(self.window.release_cid)

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

        self.window.canvas.draw_idle()

        # Refresh vline text labels after mouse wheel zoom
        self.window.refresh_vline_text_labels()

    def on_motion(self, event):
        # CTRL+DRAG MOTION HANDLING - MUST BE FIRST
        if self.ctrl_drag_active and event.inaxes:
            if self.window.vline1 is not None and self.window.vline2 is not None:

                # Calculate mouse movement delta
                mouse_delta = event.xdata - self.ctrl_drag_reference_pos

                # Get current vline1 position as reference
                current_vline1_x = self.window.vline1.get_xdata()[0]

                # Calculate new positions
                new_vline1_x = current_vline1_x + mouse_delta
                new_vline2_x = new_vline1_x + self.vline_gap  # Maintain exact gap

                # Update both vline positions simultaneously
                self.window.vline1.set_xdata([new_vline1_x, new_vline1_x])
                self.window.vline2.set_xdata([new_vline2_x, new_vline2_x])

                # Update range controls
                if hasattr(self.window, 'fitting_window') and self.window.fitting_window:
                    min_pos = min(new_vline1_x, new_vline2_x)
                    max_pos = max(new_vline1_x, new_vline2_x)

                    try:
                        self.window.fitting_window.updating_range_controls = True
                        self.window.fitting_window.min_range_text.SetValue(f"{min_pos:.2f}")
                        self.window.fitting_window.max_range_text.SetValue(f"{max_pos:.2f}")
                        self.window.fitting_window.updating_range_controls = False
                    except:
                        pass

                # UPDATE VLINE TEXT LABELS - Try multiple methods
                try:
                    # Method 1: Direct update if method exists
                    if hasattr(self.window, 'update_vline_text_labels'):
                        self.window.update_vline_text_labels()

                    # Method 2: Refresh if method exists
                    if hasattr(self.window, 'refresh_vline_text_labels'):
                        self.window.refresh_vline_text_labels()

                    # Method 3: Update text positions directly if text objects exist
                    if hasattr(self.window, 'vline1_text') and self.window.vline1_text:
                        y_pos = self.window.vline1_text.get_position()[1]
                        self.window.vline1_text.set_position((new_vline1_x, y_pos))

                    if hasattr(self.window, 'vline2_text') and self.window.vline2_text:
                        y_pos = self.window.vline2_text.get_position()[1]
                        self.window.vline2_text.set_position((new_vline2_x, y_pos))

                    # Method 4: Update text content if text objects exist
                    if hasattr(self.window, 'vline1_text') and self.window.vline1_text:
                        self.window.vline1_text.set_text(f"{new_vline1_x:.1f}")

                    if hasattr(self.window, 'vline2_text') and self.window.vline2_text:
                        self.window.vline2_text.set_text(f"{new_vline2_x:.1f}")

                    # UPDATE AVERAGING INDICATOR LINES
                    if hasattr(self.window, 'add_averaging_indicator_lines'):
                        self.window.add_averaging_indicator_lines()

                except Exception as e:
                    print(f"Error updating vline text labels: {e}")

                # Update auto-detection for area fitting
                if hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected:
                    if hasattr(self.window.fitting_window, 'auto_detect_area_name'):
                        self.window.fitting_window.auto_detect_area_name(new_vline1_x, new_vline2_x)
                    if hasattr(self.window.fitting_window, 'update_core_level_list_if_open'):
                        self.window.fitting_window.update_core_level_list_if_open()

                # Update reference position for smooth dragging
                self.ctrl_drag_reference_pos = event.xdata

                # Redraw canvas
                self.window.canvas.draw_idle()

                return  # CRITICAL: Exit here to prevent normal motion handling
        elif event.button == 1 and event.key == 'shift' and self.window.background_tab_selected:
            # Store current vline positions
            current_vline1_pos = None
            current_vline2_pos = None
            if self.window.vline1 is not None:
                current_vline1_pos = self.window.vline1.get_xdata()[0]
            if self.window.vline2 is not None:
                current_vline2_pos = self.window.vline2.get_xdata()[0]

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
                        calculated_offset = min(calculated_offset, 0)

                        # REGION-SPECIFIC UPDATE:
                        self.window.fitting_window.offset_l_text.SetValue(f'{calculated_offset:.1f}')
                        if hasattr(self.window.fitting_window,
                                   'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                            offset_h_value = float(self.window.fitting_window.offset_h_text.GetValue())
                            self.window.fitting_window.update_active_range_offsets(offset_h_value, calculated_offset)
                        self.window.set_offset_l(calculated_offset)

                    else:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        calculated_offset = min(calculated_offset, 0)

                        # REGION-SPECIFIC UPDATE:
                        self.window.fitting_window.offset_h_text.SetValue(f'{calculated_offset:.1f}')
                        if hasattr(self.window.fitting_window,
                                   'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                            offset_l_value = float(self.window.fitting_window.offset_l_text.GetValue())
                            self.window.fitting_window.update_active_range_offsets(calculated_offset, offset_l_value)
                        self.window.set_offset_h(calculated_offset)

                else:
                    raw_y = self.window.y_values[np.argmin(np.abs(self.window.x_values - vline2_x))]
                    if vline2_x == low_be_x:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        calculated_offset = min(calculated_offset, 0)

                        # REGION-SPECIFIC UPDATE:
                        self.window.fitting_window.offset_l_text.SetValue(f'{calculated_offset:.1f}')
                        if hasattr(self.window.fitting_window,
                                   'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                            offset_h_value = float(self.window.fitting_window.offset_h_text.GetValue())
                            self.window.fitting_window.update_active_range_offsets(offset_h_value, calculated_offset)
                        self.window.set_offset_l(calculated_offset)

                    else:
                        calculated_offset = event.ydata - raw_y
                        # Ensure offset cannot be positive
                        calculated_offset = min(calculated_offset, 0)

                        # REGION-SPECIFIC UPDATE:
                        self.window.fitting_window.offset_h_text.SetValue(f'{calculated_offset:.1f}')
                        if hasattr(self.window.fitting_window,
                                   'active_range_index') and self.window.fitting_window.active_range_index >= 0:
                            offset_l_value = float(self.window.fitting_window.offset_l_text.GetValue())
                            self.window.fitting_window.update_active_range_offsets(calculated_offset, offset_l_value)
                        self.window.set_offset_h(calculated_offset)

                # Store old 'Bkg Offset' values in window.data
                self.window.Data['Core levels'][sheet_name]['Background']['Bkg Offset Low'] = self.window.offset_l
                self.window.Data['Core levels'][sheet_name]['Background']['Bkg Offset High'] = self.window.offset_h

                self.window.plot_manager.plot_background(self.window)

                # Force correct legend update for Multi-Regions Smart background
                if hasattr(self.window.plot_manager, 'update_legend'):
                    self.window.plot_manager.update_legend(self.window)

                # Restore vlines after plotting
                if current_vline1_pos is not None and current_vline2_pos is not None:
                    wx.CallAfter(self.restore_vlines_after_plot, current_vline1_pos, current_vline2_pos)

        elif event.inaxes and self.window.moving_vline is not None:
            x_click = event.xdata
            self.window.moving_vline.set_xdata([x_click])

            sheet_name = self.window.sheet_combobox.GetValue()
            if sheet_name in self.window.Data['Core levels']:
                core_level_data = self.window.Data['Core levels'][sheet_name]

                if self.window.moving_vline in [self.window.vline1, self.window.vline2]:
                    # Convert display position back to BE for storage
                    be_position = self.window.convert_energy_from_display(x_click)
                    if self.window.moving_vline == self.window.vline1:
                        core_level_data['Background']['Bkg Low'] = float(be_position)
                    else:
                        core_level_data['Background']['Bkg High'] = float(be_position)

                    bkg_low = core_level_data['Background']['Bkg Low']
                    bkg_high = core_level_data['Background']['Bkg High']
                    core_level_data['Background']['Bkg Low'] = min(bkg_low, bkg_high)
                    core_level_data['Background']['Bkg High'] = max(bkg_low, bkg_high)

                    # Update text labels when vlines are moved
                    self.window.update_vline_text_labels()

                    # UPDATE AVERAGING INDICATOR LINES
                    if hasattr(self.window, 'add_averaging_indicator_lines'):
                        self.window.add_averaging_indicator_lines()

                    # Update fitting screen range controls
                    self.window.update_fitting_screen_range_controls()

                    # Update area screen range controls when vlines move
                    if (hasattr(self.window, 'background_window') and self.window.background_window is not None and
                            hasattr(self.window.background_window, 'update_vline_text_labels') and
                            hasattr(self.window.background_window, 'update_range_controls_from_data')):
                        if self.window.moving_vline in [self.window.vline1, self.window.vline2]:
                            self.window.background_window.update_vline_text_labels()
                            self.window.background_window.update_range_controls_from_data()
                            self.window.update_area_screen_range_controls()

                elif self.window.moving_vline in [self.window.vline3, self.window.vline4]:
                    if self.window.moving_vline == self.window.vline3:
                        self.window.noise_min_energy = float(x_click)
                    else:
                        self.window.noise_max_energy = float(x_click)

                    self.window.noise_min_energy, self.window.noise_max_energy = sorted(
                        [self.window.noise_min_energy, self.window.noise_max_energy])

            # Update VBM controls during dragging
            self.update_vbm_controls_from_vlines()

            self.window.canvas.draw_idle()

        # Update vline text labels when moving vlines
        if self.window.moving_vline in [self.window.vline1, self.window.vline2]:
            self.window.update_vline_text_labels()

        # UPDATE AVERAGING INDICATOR LINES
        if hasattr(self.window, 'add_averaging_indicator_lines'):
            self.window.add_averaging_indicator_lines()

    def redraw_all_regions_background_OLD(self):
        """Delete whole background and redraw from region 1, 2, 3... in sequence"""
        if not hasattr(self.window, 'fitting_window') or self.window.fitting_window is None:
            return

        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name not in self.window.Data['Core levels']:
            return

        # Get all recorded ranges from window.data
        ranges = self.window.fitting_window.get_recorded_ranges_from_data()
        if not ranges:
            return

        # Store current vline positions before they get destroyed
        current_vline1_pos = None
        current_vline2_pos = None
        if self.window.vline1 is not None:
            current_vline1_pos = self.window.vline1.get_xdata()[0]
        if self.window.vline2 is not None:
            current_vline2_pos = self.window.vline2.get_xdata()[0]

        # Clear existing background
        x_values = np.array(self.window.Data['Core levels'][sheet_name]['B.E.'], dtype=float)
        y_values = np.array(self.window.Data['Core levels'][sheet_name]['Raw Data'], dtype=float)

        # Initialize background to raw data
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values.tolist()
        current_background = np.array(y_values)

        # Get active region index
        active_region_index = getattr(self.window.fitting_window, 'active_range_index', -1)

        # Apply each region in sequence: region 1, then 2, then 3, etc.
        for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):
            # # Calculate background for this specific region using window.data offset values
            # from libraries.Peak_Functions import BackgroundCalculations
            # current_background = BackgroundCalculations.calculate_adaptive_smart_background(
            #     x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l
            # )

            # Calculate background for this specific region based on selected method
            from libraries.Peak_Functions import BackgroundCalculations
            method = self.window.background_method

            # SPECIAL HANDLING FOR TOUGAARD METHODS - they cannot be applied region-by-region
            if method in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]:
                # Skip region-by-region processing for Tougaard methods
                # They will be handled by the final step outside the loop
                pass
            elif method == "Multi-Regions Smart":
                current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                    x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
            elif method == "Shirley":
                current_background = BackgroundCalculations.calculate_adaptive_shirley_background(
                    x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
            elif method == "Linear":
                current_background = BackgroundCalculations.calculate_adaptive_linear_background(
                    x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
            elif method == "Smart":
                current_background = BackgroundCalculations.calculate_adaptive_single_smart_background(
                    x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
            else:
                # Fallback to smart for unknown methods
                current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                    x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)

            # Handle Tougaard methods AFTER processing all regions
            method = self.window.background_method
            if method in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]:
                print(f"Applying {method} to all regions...")

                # Store original values
                temp_bg_min = self.window.bg_min_energy
                temp_bg_max = self.window.bg_max_energy

                # Set full range for Tougaard calculation
                self.window.bg_min_energy = float(np.min(x_values))
                self.window.bg_max_energy = float(np.max(x_values))

                try:
                    # Calculate Tougaard background for entire spectrum or per-region
                    if method == "U4-Tougaard":
                        full_tougaard_bg = BackgroundCalculations.calculate_tougaard_background(
                            x_values, y_values, sheet_name, self.window)
                    elif method == "U2-Tougaard":
                        # For U2-Tougaard: Calculate backgrounds region by region using CURRENT background state
                        current_background = np.array(y_values)  # Start with raw data

                        for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):
                            print(f"Calculating U2-Tougaard for region {i + 1}: {min_range:.2f} - {max_range:.2f} eV")

                            # Use the current cumulative background state (includes previous regions)
                            # This is critical - each region builds on the previous ones
                            region_vline_range = (min_range, max_range)

                            # Pass the current background state, not raw y_values
                            region_tougaard_bg = BackgroundCalculations.calculate_u2_tougaard_background(
                                x_values, current_background, sheet_name, self.window, region_vline_range)

                            # Apply only to this region
                            region_mask = (x_values >= min_range) & (x_values <= max_range)
                            current_background[region_mask] = region_tougaard_bg[region_mask]

                            # Get the fitted parameters for logging
                            fitted_b = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_B', 0)
                            fitted_c = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_C', 1643)
                            print(f"Region {i + 1} U2-Tougaard: B={fitted_b:.2f}, C={fitted_c:.2f} applied to {min_range:.2f}-{max_range:.2f}")

                        # Don't apply again outside the loop
                        full_tougaard_bg = current_background
                    else:
                        full_tougaard_bg = current_background  # fallback

                    # Apply Tougaard background ONLY in the defined regions (skip for U2 since already done)
                    if method != "U2-Tougaard":
                        current_background = np.array(y_values)  # Start fresh with raw data
                        for offset_h, offset_l, min_range, max_range in ranges:
                            region_mask = (x_values >= min_range) & (x_values <= max_range)
                            current_background[region_mask] = full_tougaard_bg[region_mask]

                finally:
                    # Restore original values
                    self.window.bg_min_energy = temp_bg_min
                    self.window.bg_max_energy = temp_bg_max
        # Update the final background
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = current_background.tolist()
        self.window.background = current_background

        # CRITICAL: Update window offset values to match active region offsets
        # This ensures plot_background() uses correct offsets for the active region
        if active_region_index >= 0 and active_region_index < len(ranges):
            active_offset_h, active_offset_l, _, _ = ranges[active_region_index]
            self.window.offset_h = active_offset_h
            self.window.offset_l = active_offset_l
            # print(f"Updated window offsets to active region values: {active_offset_h:.1f}, {active_offset_l:.1f}")

        # Redraw the plot
        self.window.plot_manager.plot_background(self.window)

        # Force correct legend update for Multi-Regions Smart background
        if hasattr(self.window.plot_manager, 'update_legend'):
            self.window.plot_manager.update_legend(self.window)

        # RESTORE vLines after plotting (they get destroyed by clear_and_replot)
        if current_vline1_pos is not None and current_vline2_pos is not None:
            # Force vline recreation at stored positions
            wx.CallAfter(self.restore_vlines_after_plot, current_vline1_pos, current_vline2_pos)

    def redraw_all_regions_background_OLD2(self):
        """Delete whole background and redraw from region 1, 2, 3... in sequence"""

        from libraries.Peak_Functions import BackgroundCalculations
        if not hasattr(self.window, 'fitting_window') or self.window.fitting_window is None:
            return

        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name not in self.window.Data['Core levels']:
            return

        # Get all recorded ranges from window.data
        ranges = self.window.fitting_window.get_recorded_ranges_from_data()
        if not ranges:
            return

        # Store current vline positions before they get destroyed
        current_vline1_pos = None
        current_vline2_pos = None
        if self.window.vline1 is not None:
            current_vline1_pos = self.window.vline1.get_xdata()[0]
        if self.window.vline2 is not None:
            current_vline2_pos = self.window.vline2.get_xdata()[0]

        # Clear existing background
        x_values = np.array(self.window.Data['Core levels'][sheet_name]['B.E.'], dtype=float)
        y_values = np.array(self.window.Data['Core levels'][sheet_name]['Raw Data'], dtype=float)

        # Initialize background to raw data
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values.tolist()
        current_background = np.array(y_values)

        # Get active region index
        active_region_index = getattr(self.window.fitting_window, 'active_range_index', -1)

        method = self.window.background_method

        # HANDLE TOUGAARD METHODS SEPARATELY (OUTSIDE THE MAIN LOOP)
        if method in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]:
            print(f"Applying {method} to all regions...")

            # Store original values
            temp_bg_min = self.window.bg_min_energy
            temp_bg_max = self.window.bg_max_energy

            # Set full range for Tougaard calculation
            self.window.bg_min_energy = float(np.min(x_values))
            self.window.bg_max_energy = float(np.max(x_values))

            try:
                if method == "U4-Tougaard":
                    full_tougaard_bg = BackgroundCalculations.calculate_tougaard_background(
                        x_values, y_values, sheet_name, self.window)
                    # Apply to all regions
                    for offset_h, offset_l, min_range, max_range in ranges:
                        region_mask = (x_values >= min_range) & (x_values <= max_range)
                        current_background[region_mask] = full_tougaard_bg[region_mask]

                elif method == "U2-Tougaard":
                    # For U2-Tougaard: Calculate each region independently
                    for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):
                        print(f"Calculating U2-Tougaard for region {i + 1}: {min_range:.2f} - {max_range:.2f} eV")

                        region_vline_range = (min_range, max_range)
                        region_tougaard_bg = BackgroundCalculations.calculate_u2_tougaard_background(
                            x_values, current_background, sheet_name, self.window, region_vline_range)

                        # Apply only to this region
                        region_mask = (x_values >= min_range) & (x_values <= max_range)
                        current_background[region_mask] = region_tougaard_bg[region_mask]

                        # Get the fitted parameters for logging
                        fitted_b = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_B', 0)
                        fitted_c = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_C', 1643)
                        print(f"Region {i + 1} U2-Tougaard: B={fitted_b:.2f}, C={fitted_c:.2f} applied to {min_range:.2f}-{max_range:.2f}")
            finally:
                # Restore original values
                self.window.bg_min_energy = temp_bg_min
                self.window.bg_max_energy = temp_bg_max

        else:
            # HANDLE NON-TOUGAARD METHODS WITH MAIN LOOP
            for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):


                if method == "Multi-Regions Smart":
                    current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Shirley":
                    current_background = BackgroundCalculations.calculate_adaptive_shirley_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Linear":
                    current_background = BackgroundCalculations.calculate_adaptive_linear_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Smart":
                    current_background = BackgroundCalculations.calculate_adaptive_single_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                else:
                    # Fallback to smart for unknown methods
                    current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)

        # Update the final background
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = current_background.tolist()
        self.window.background = current_background

        # CRITICAL: Update window offset values to match active region offsets
        if active_region_index >= 0 and active_region_index < len(ranges):
            active_offset_h, active_offset_l, _, _ = ranges[active_region_index]
            self.window.offset_h = active_offset_h
            self.window.offset_l = active_offset_l

        # Redraw the plot
        self.window.plot_manager.plot_background(self.window)

        # Force correct legend update
        if hasattr(self.window.plot_manager, 'update_legend'):
            self.window.plot_manager.update_legend(self.window)

        # RESTORE vLines after plotting
        if current_vline1_pos is not None and current_vline2_pos is not None:
            wx.CallAfter(self.restore_vlines_after_plot, current_vline1_pos, current_vline2_pos)

    def redraw_all_regions_background(self):
        """Delete whole background and redraw from region 1, 2, 3... in sequence"""

        from libraries.Peak_Functions import BackgroundCalculations
        if not hasattr(self.window, 'fitting_window') or self.window.fitting_window is None:
            return

        sheet_name = self.window.sheet_combobox.GetValue()
        if sheet_name not in self.window.Data['Core levels']:
            return

        # Get all recorded ranges from window.data
        ranges = self.window.fitting_window.get_recorded_ranges_from_data()
        if not ranges:
            return

        # Store current vline positions before they get destroyed
        current_vline1_pos = None
        current_vline2_pos = None
        if self.window.vline1 is not None:
            current_vline1_pos = self.window.vline1.get_xdata()[0]
        if self.window.vline2 is not None:
            current_vline2_pos = self.window.vline2.get_xdata()[0]

        # Clear existing background
        x_values = np.array(self.window.Data['Core levels'][sheet_name]['B.E.'], dtype=float)
        y_values = np.array(self.window.Data['Core levels'][sheet_name]['Raw Data'], dtype=float)

        # Initialize background to raw data
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values.tolist()
        current_background = np.array(y_values)

        # Get active region index
        active_region_index = getattr(self.window.fitting_window, 'active_range_index', -1)

        method = self.window.background_method

        # HANDLE TOUGAARD METHODS SEPARATELY (OUTSIDE THE MAIN LOOP)
        if method in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]:
            print(f"Applying {method} to all regions...")

            # Store original values
            temp_bg_min = self.window.bg_min_energy
            temp_bg_max = self.window.bg_max_energy

            # Set full range for Tougaard calculation
            self.window.bg_min_energy = float(np.min(x_values))
            self.window.bg_max_energy = float(np.max(x_values))

            try:
                if method == "U4-Tougaard":
                    full_tougaard_bg = BackgroundCalculations.calculate_tougaard_background(
                        x_values, y_values, sheet_name, self.window)
                    # Apply to all regions
                    for offset_h, offset_l, min_range, max_range in ranges:
                        region_mask = (x_values >= min_range) & (x_values <= max_range)
                        current_background[region_mask] = full_tougaard_bg[region_mask]

                elif method == "U2-Tougaard":
                    # For U2-Tougaard: Calculate each region independently using ONLY region data
                    for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):
                        print(f"Calculating U2-Tougaard for region {i + 1}: {min_range:.2f} - {max_range:.2f} eV")

                        # CRITICAL: Extract only the data within this region's range
                        region_mask = (x_values >= min_range) & (x_values <= max_range)
                        x_region = x_values[region_mask]
                        y_region = y_values[region_mask]

                        if len(x_region) < 3:  # Need minimum points for calculation
                            print(f"Warning: Region {i + 1} has insufficient data points, skipping")
                            continue

                        region_vline_range = (min_range, max_range)

                        # Calculate U2-Tougaard using ONLY the region data
                        region_tougaard_bg = BackgroundCalculations.calculate_u2_tougaard_background(
                            x_region, y_region, sheet_name, self.window, region_vline_range)

                        # Apply the calculated background to this region in the full spectrum
                        current_background[region_mask] = region_tougaard_bg

                        # Get the fitted parameters for logging
                        fitted_b = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_B', 0)
                        fitted_c = self.window.Data['Core levels'][sheet_name]['Background'].get('Fitted_C', 1643)
                        print(f"Region {i + 1} U2-Tougaard: B={fitted_b:.2f}, C={fitted_c:.2f} applied to {min_range:.2f}-{max_range:.2f}")
                        print(f"  Using {len(x_region)} data points from region")
            finally:
                # Restore original values
                self.window.bg_min_energy = temp_bg_min
                self.window.bg_max_energy = temp_bg_max

        else:
            # HANDLE NON-TOUGAARD METHODS WITH MAIN LOOP
            for i, (offset_h, offset_l, min_range, max_range) in enumerate(ranges):
                if method == "Multi-Regions Smart":
                    current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Shirley":
                    current_background = BackgroundCalculations.calculate_adaptive_shirley_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Linear":
                    current_background = BackgroundCalculations.calculate_adaptive_linear_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                elif method == "Smart":
                    current_background = BackgroundCalculations.calculate_adaptive_single_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)
                else:
                    # Fallback to smart for unknown methods
                    current_background = BackgroundCalculations.calculate_adaptive_smart_background(
                        x_values, y_values, (min_range, max_range), current_background, offset_h, offset_l)

        # Update the final background
        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = current_background.tolist()
        self.window.background = current_background

        # CRITICAL: Update window offset values to match active region offsets
        if active_region_index >= 0 and active_region_index < len(ranges):
            active_offset_h, active_offset_l, _, _ = ranges[active_region_index]
            self.window.offset_h = active_offset_h
            self.window.offset_l = active_offset_l

        # Only redraw plot for non-Tougaard methods (Tougaard background is already calculated above)
        if method not in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]:
            self.window.plot_manager.plot_background(self.window)
        else:
            print(f"Skipped plot_background call for {method} - background already calculated")
            # For Tougaard: plot data and background, then replot peaks if they exist
            self.window.plot_manager.plot_data(self.window)

            # Plot the background manually (same as plot_background method)
            x_values = np.array(self.window.Data['Core levels'][sheet_name]['B.E.'])
            self.window.ax.plot(x_values, current_background,
                                color=self.window.plot_manager.background_color,
                                linestyle=self.window.plot_manager.background_linestyle,
                                alpha=self.window.plot_manager.background_alpha,
                                linewidth=self.window.plot_manager.background_thickness,
                                label='Background (U2-Tougaard)' if method == "U2-Tougaard" else f'Background ({method})')

            # Replot peaks if they exist (same as regular plot_background)
            if self.window.peak_params_grid.GetNumberRows() > 0:
                self.window.clear_and_replot()

        # Force correct legend update
        if hasattr(self.window.plot_manager, 'update_legend'):
            self.window.plot_manager.update_legend(self.window)

        # RESTORE vLines after plotting
        if current_vline1_pos is not None and current_vline2_pos is not None:
            wx.CallAfter(self.restore_vlines_after_plot, current_vline1_pos, current_vline2_pos)

    def restore_vlines_after_plot(self, vline1_pos, vline2_pos):
        """Restore vlines at specified positions after plotting"""
        try:
            # FIRST: Remove/destroy any existing vlines
            if self.window.vline1 is not None:
                try:
                    self.window.vline1.remove()
                except:
                    pass
                self.window.vline1 = None

            if self.window.vline2 is not None:
                try:
                    self.window.vline2.remove()
                except:
                    pass
                self.window.vline2 = None

            # Remove any existing text labels
            if hasattr(self.window, 'vline1_text') and self.window.vline1_text is not None:
                try:
                    self.window.vline1_text.remove()
                except:
                    pass
                self.window.vline1_text = None

            if hasattr(self.window, 'vline2_text') and self.window.vline2_text is not None:
                try:
                    self.window.vline2_text.remove()
                except:
                    pass
                self.window.vline2_text = None

            # THEN: Create new vlines at the specified positions
            self.window.vline1 = self.window.ax.axvline(x=vline1_pos, color='red', linestyle='--', alpha=0.7)
            self.window.vline2 = self.window.ax.axvline(x=vline2_pos, color='red', linestyle='--', alpha=0.7)

            # Create new text labels
            if hasattr(self.window, 'add_vline_text_labels'):
                self.window.add_vline_text_labels()
            elif hasattr(self.window, 'update_vline_text_labels'):
                self.window.update_vline_text_labels()

            # UPDATE AVERAGING INDICATOR LINES
            if hasattr(self.window, 'add_averaging_indicator_lines'):
                self.window.add_averaging_indicator_lines()

            # Make sure they're visible
            self.window.show_hide_vlines()

            # Force canvas redraw
            self.window.canvas.draw_idle()

            # print(f"Restored vlines at positions: {vline1_pos:.2f}, {vline2_pos:.2f}")

        except Exception as e:
            print(f"Error restoring vlines: {e}")

    def update_active_region_positions(self):
        """Record new vLine positions in the active region's min/max range and window.data"""
        if (not hasattr(self.window, 'fitting_window') or
                not hasattr(self.window.fitting_window, 'active_range_index') or
                self.window.fitting_window.active_range_index < 0):
            # Still update VBM controls even if no fitting window active
            self.update_vbm_controls_from_vlines()
            return

        # Get current vline positions
        if self.window.vline1 is None or self.window.vline2 is None:
            return

        vline1_x = self.window.vline1.get_xdata()[0]
        vline2_x = self.window.vline2.get_xdata()[0]

        # Round to 2 decimal places before storing
        min_pos = round(min(vline1_x, vline2_x), 2)
        max_pos = round(max(vline1_x, vline2_x), 2)

        # Get current offset values from UI controls instead of stored values
        try:
            current_offset_h = float(self.window.fitting_window.offset_h_text.GetValue())
            current_offset_l = float(self.window.fitting_window.offset_l_text.GetValue())
        except (ValueError, AttributeError):
            current_offset_h = 0.0
            current_offset_l = 0.0

        # Update the active region's positions
        ranges = self.window.fitting_window.get_recorded_ranges_from_data()
        active_idx = self.window.fitting_window.active_range_index

        if active_idx < len(ranges):
            # Use current UI offset values instead of stored ones
            old_offset_h, old_offset_l, old_min, old_max = ranges[active_idx]
            ranges[active_idx] = (current_offset_h, current_offset_l, min_pos, max_pos)

            # Save back to window.data
            sheet_name = self.window.sheet_combobox.GetValue()
            if 'Background' not in self.window.Data['Core levels'][sheet_name]:
                self.window.Data['Core levels'][sheet_name]['Background'] = {}
            self.window.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = ranges

            # Update the UI range controls
            self.window.fitting_window.updating_range_controls = True
            self.window.fitting_window.min_range_text.SetValue(f"{min_pos:.2f}")
            self.window.fitting_window.max_range_text.SetValue(f"{max_pos:.2f}")
            self.window.fitting_window.updating_range_controls = False

        # Update VBM controls if VBM window is open
        self.update_vbm_controls_from_vlines()

    def cleanup_vline_handlers(self):
        """Clean up any existing vline event handlers"""
        if hasattr(self.window, 'motion_cid'):
            self.window.canvas.mpl_disconnect(self.window.motion_cid)
            delattr(self.window, 'motion_cid')
        if hasattr(self.window, 'release_cid'):
            self.window.canvas.mpl_disconnect(self.window.release_cid)
            delattr(self.window, 'release_cid')

        # Reset moving vline state
        self.window.moving_vline = None

    def on_release_OLD(self, event):

        if self.ctrl_drag_active:
            print(f"CTRL+drag ended")

            # Reset CTRL+drag state
            self.ctrl_drag_active = False
            self.vline_gap = 0.0
            self.ctrl_drag_reference_pos = 0.0

            # Ensure moving_vline is None to prevent conflicts
            self.window.moving_vline = None

            # UPDATE BACKGROUND - Use the same logic as single vline dragging
            if (self.window.background_tab_selected and
                    hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None):
                # Update active region positions with new vLine positions (same as single vline)
                self.update_active_region_positions()

                # Redraw all regions in sequence (same as single vline)
                self.redraw_all_regions_background()

            # Alternative background update for other cases
            elif (self.window.background_tab_selected and
                  hasattr(self.window, 'background_method') and
                  self.window.background_method == "Multi-Regions Smart"):
                if hasattr(self.window, 'plot_manager'):
                    print('Helloooooo22222')
                    self.window.plot_manager.plot_background(self.window)

            # CRITICAL: Skip regular background plotting for Tougaard methods with multiple regions
            # The redraw_all_regions_background() call above already handled them properly
            elif (self.window.background_tab_selected and
                  hasattr(self.window, 'background_method') and
                  self.window.background_method in ["U4-Tougaard", "U2-Tougaard", "2x U4-Tougaard", "3x U4-Tougaard"]):
                # Skip regular background plotting - already handled by redraw_all_regions_background()
                print(f"Skipping regular background plot for {self.window.background_method} - already handled by multi-region code")
                pass

            # Save state after movement
            save_state(self.window)

            # Clean up motion and release handlers
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            return  # CRITICAL: Exit here to prevent normal release handling
        elif self.window.moving_vline is not None:
            # Store which vline was moved before resetting to None
            moved_vline = self.window.moving_vline

            # Save state after vline movement and background update
            save_state(self.window)

            # Update VBM controls if VBM window is open
            self.update_vbm_controls_from_vlines()

            # Use the correct variable names to disconnect events
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            # Reset the moving vline to None
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

                # if (moved_vline in [self.window.vline1, self.window.vline2] and
                #         self.window.background_method == "Multi-Regions Smart" and
                #         hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None):
                if (moved_vline in [self.window.vline1, self.window.vline2] and
                            hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None):
                    # Update active region positions with new vLine positions
                    self.update_active_region_positions()
                    print('Helllo')
                    # # Redraw all regions in sequence
                    self.redraw_all_regions_background()

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

    def on_release(self, event):
        if self.ctrl_drag_active:
            print(f"CTRL+drag ended")

            # Reset CTRL+drag state
            self.ctrl_drag_active = False
            self.vline_gap = 0.0
            self.ctrl_drag_reference_pos = 0.0

            # Ensure moving_vline is None to prevent conflicts
            self.window.moving_vline = None

            # UPDATE BACKGROUND - Use the same logic as single vline dragging
            if (self.window.background_tab_selected and
                    hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None):
                # Update active region positions with new vLine positions (same as single vline)
                self.update_active_region_positions()

                # Redraw all regions in sequence (same as single vline)
                self.redraw_all_regions_background()

            # Alternative background update for Multi-Regions Smart only (non-fitting window case)
            elif (self.window.background_tab_selected and
                  hasattr(self.window, 'background_method') and
                  self.window.background_method == "Multi-Regions Smart" and
                  not hasattr(self.window, 'fitting_window')):
                if hasattr(self.window, 'plot_manager'):
                    print('Multi-Regions Smart fallback')
                    self.window.plot_manager.plot_background(self.window)

            # Save state after movement
            save_state(self.window)

            # Clean up motion and release handlers
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            return  # CRITICAL: Exit here to prevent normal release handling

        elif self.window.moving_vline is not None:
            # Store which vline was moved before resetting to None
            moved_vline = self.window.moving_vline

            # Save state after vline movement and background update
            save_state(self.window)

            # Update VBM controls if VBM window is open
            self.update_vbm_controls_from_vlines()

            # Use the correct variable names to disconnect events
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            # Reset the moving vline to None
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

                # Handle background redraw for fitting window cases
                if (moved_vline in [self.window.vline1, self.window.vline2] and
                        hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None):
                    # Update active region positions with new vLine positions
                    self.update_active_region_positions()
                    print('Regular vline drag - redrawing all regions')
                    # Redraw all regions in sequence
                    self.redraw_all_regions_background()

        # Handle peak updates
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
            from libraries.FileMenu.Save import copy_all_peak_parameters, paste_all_peak_parameters, copy_core_level, \
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

        # Add export to results grid option
        export_item = menu.Append(wx.ID_ANY, "Export to Results Grid")
        from libraries.FileMenu.Export import export_results
        self.window.Bind(wx.EVT_MENU, lambda evt: export_results(self.window), export_item)

        menu.AppendSeparator()



        # New peak operations - available on both parameter and constraint rows
        # Determine peak index and letter based on row type
        if row % 2 == 0:  # Parameter row
            peak_index = row // 2
            peak_letter = self.window.peak_params_grid.GetCellValue(row, 0)
        else:  # Constraint row
            peak_index = row // 2  # Integer division gives us the peak index
            peak_letter = self.window.peak_params_grid.GetCellValue(row - 1, 0)  # Get letter from parameter row above

        delete_item = menu.Append(wx.ID_ANY, f"Delete Peak {peak_letter}")

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

        if col in [2, 3, 4, 5, 6, 7, 8, 9]:  # Available on both parameter and constraint rows
            # Determine parameter row based on whether we're on parameter row or constraint row
            if row % 2 == 0:  # Parameter row
                param_row = row
            else:  # Constraint row
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

        # NEW CODE - Add cross-core-level constraint menu
        if col in [2, 3, 4, 5, 6, 7, 8, 9] :
            menu.AppendSeparator()

            # Create cross-core-level constraint submenu
            cross_core_menu = wx.Menu()

            # Get current sheet name for comparison
            current_sheet_name = self.window.sheet_combobox.GetValue()

            # Get all core levels that have fitting data
            available_core_levels = {}
            for core_level_name, core_level_data in self.window.Data['Core levels'].items():
                if ('Fitting' in core_level_data and
                        'Peaks' in core_level_data['Fitting'] and
                        len(core_level_data['Fitting']['Peaks']) > 0):
                    available_core_levels[core_level_name] = core_level_data['Fitting']['Peaks']

            if available_core_levels:
                # Create submenu for each core level
                for core_level_name, peaks in available_core_levels.items():
                    core_level_submenu = wx.Menu()

                    # Add each peak as a menu item
                    peak_keys = list(peaks.keys())
                    for i, peak_key in enumerate(peak_keys):
                        peak_letter = chr(65 + i)  # A, B, C, etc.
                        constraint_ref = f"{core_level_name}_{peak_letter}"

                        # Show just the letter if it's the same core level, otherwise show full reference
                        if core_level_name == current_sheet_name:
                            display_name = peak_letter  # Just "A", "B", "C"
                        else:
                            display_name = constraint_ref  # "C1s_A", "Sr3d_B", etc.

                        # Create menu item for this peak
                        peak_item = core_level_submenu.Append(wx.ID_ANY, display_name)

                        # Bind the menu item to insert the constraint (always use full constraint_ref)
                        self.window.Bind(wx.EVT_MENU,
                                         lambda evt, ref=constraint_ref, r=row, c=col:
                                         self.insert_cross_core_constraint(ref, r, c),
                                         peak_item)

                    # Add the core level submenu to the main cross-core menu
                    cross_core_menu.AppendSubMenu(core_level_submenu, core_level_name)

                # Add the main cross-core menu to the context menu
                menu.AppendSubMenu(cross_core_menu, "Constraint to Other Core Levels")

        # Enable/disable items
        clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')
        has_clipboard_data = os.path.exists(clipboard_file)
        has_rows = self.window.peak_params_grid.GetNumberRows() > 0

        copy_item.Enable(has_rows)
        paste_item.Enable(has_clipboard_data)
        delete_item.Enable(has_rows)
        propagate_item.Enable(col in [2, 3, 4, 5, 6, 7, 8, 9])

        # Bind events
        from libraries.FileMenu.Save import copy_all_peak_parameters, paste_all_peak_parameters
        from libraries.Utilities import propagate_constraint, propagate_fwhm_difference

        self.window.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self.window), copy_item)
        self.window.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self.window), paste_item)


        self.window.Bind(wx.EVT_MENU, lambda evt: self.delete_peak_at_index(peak_index), delete_item)

        for i, add_item in enumerate(add_items):
            model = models[i]
            self.window.Bind(wx.EVT_MENU, lambda evt, m=model, r=row: self.add_peak_with_model(m, r), add_item)

        # Always pass the constraint row to propagate_constraint
        constraint_row = row + 1 if row % 2 == 0 else row
        self.window.Bind(wx.EVT_MENU, lambda evt: propagate_constraint(self.window, constraint_row, col),
                         propagate_item)
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
        if (hasattr(self.window, 'fitting_window') and
                hasattr(self.window.fitting_window, 'get_overall_background_range')):
            bg_low, bg_high = self.window.fitting_window.get_overall_background_range()
        else:
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

    def update_vbm_controls_from_vlines(self):
        """Update VBM controls when vlines move"""
        if (hasattr(self.window, 'vb_measurements_window') and
                self.window.vb_measurements_window is not None):
            try:
                self.window.vb_measurements_window.update_controls_from_vlines()
                # print("VBM controls updated")  # Debug line
            except Exception as e:
                print(f"Error updating VBM controls: {e}")
        # else:
        #     print("No VBM window found")  # Debug line




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