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

        self.center_drag_active = False
        self.center_drag_reference_pos = 0.0

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

                    # Check if vline_center exists before accessing it (only exists in AreaFit_Screen)
                    if (hasattr(self.window, 'vline_center') and self.window.vline_center is not None):
                        center_x = self.window.vline_center.get_xdata()[0]
                        tolerance = (max(self.window.x_values) - min(self.window.x_values)) * 0.02

                        if abs(x_click - center_x) < tolerance:
                            vline1_x = self.window.vline1.get_xdata()[0]
                            vline2_x = self.window.vline2.get_xdata()[0]
                            self.vline_gap = vline2_x - vline1_x
                            self.center_drag_reference_pos = x_click
                            self.center_drag_active = True
                            self.window.moving_vline = None

                            # Clean up existing handlers
                            if hasattr(self.window, 'motion_cid'):
                                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                                delattr(self.window, 'motion_cid')
                            if hasattr(self.window, 'release_cid'):
                                self.window.canvas.mpl_disconnect(self.window.release_cid)
                                delattr(self.window, 'release_cid')

                            # Set up motion and release handlers
                            self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                            self.window.release_cid = self.window.canvas.mpl_connect('button_release_event', self.on_release)

                            return


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
                # Handle center vline dragging (without CTRL key) - CHECK THIS FIRST
                if ((self.window.background_tab_selected or
                     (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected)) and
                        hasattr(self.window, 'vline_center') and self.window.vline_center is not None and
                        self.window.vline1 is not None and self.window.vline2 is not None):

                    center_x = self.window.vline_center.get_xdata()[0]
                    tolerance = (max(self.window.x_values) - min(self.window.x_values)) * 0.01

                    if abs(x_click - center_x) < tolerance:
                        vline1_x = self.window.vline1.get_xdata()[0]
                        vline2_x = self.window.vline2.get_xdata()[0]
                        self.vline_gap = vline2_x - vline1_x
                        self.center_drag_reference_pos = x_click
                        self.center_drag_active = True
                        self.window.moving_vline = None

                        # Clean up existing handlers
                        if hasattr(self.window, 'motion_cid'):
                            self.window.canvas.mpl_disconnect(self.window.motion_cid)
                            delattr(self.window, 'motion_cid')
                        if hasattr(self.window, 'release_cid'):
                            self.window.canvas.mpl_disconnect(self.window.release_cid)
                            delattr(self.window, 'release_cid')

                        # Set up motion and release handlers
                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event', self.on_release)
                        return
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

                # Check for green vline if active (works on all tabs)
                if (hasattr(self.window, 'green_vline_active') and self.window.green_vline_active and
                        hasattr(self.window, 'green_vline') and self.window.green_vline is not None and
                        self.window.green_vline.get_visible()):
                    green_vline_x = self.window.green_vline.get_xdata()[0]
                    dist_green = abs(x_click - green_vline_x)

                    # Calculate adaptive threshold
                    x_range = abs(max(self.window.x_values) - min(self.window.x_values))
                    adaptive_threshold = max(self.window.some_threshold, x_range * 0.02)

                    # Check if click is near green vline
                    if dist_green <= adaptive_threshold:
                        # Clean up any existing handlers first
                        if hasattr(self.window, 'motion_cid'):
                            self.window.canvas.mpl_disconnect(self.window.motion_cid)
                            delattr(self.window, 'motion_cid')
                        if hasattr(self.window, 'release_cid'):
                            self.window.canvas.mpl_disconnect(self.window.release_cid)
                            delattr(self.window, 'release_cid')

                        self.window.moving_vline = self.window.green_vline
                        self.window.moving_vline_type = 'green'
                        self.window.motion_cid = self.window.canvas.mpl_connect('motion_notify_event', self.on_motion)
                        self.window.release_cid = self.window.canvas.mpl_connect('button_release_event', self.on_release)
                        return


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

        # # DEBUG: Print the two conditions
        # print(f"DEBUG: peak_fitting_tab_selected = {self.window.peak_fitting_tab_selected}")
        # print(f"DEBUG: selected_peak_index = {self.window.selected_peak_index}")
        # print(f"DEBUG: Both conditions met = {self.window.peak_fitting_tab_selected and self.window.selected_peak_index is not None}")
        # print("---")

        shift_currently_pressed = False
        shift_currently_pressed = event.key == 'shift'

        if shift_currently_pressed:
            self.window.shift_key_pressed = True
        else:
            self.window.shift_key_pressed = False

        # Check if AreaFit screen is open and active
        if (hasattr(self.window, 'background_window') and self.window.background_window is not None and
                hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected):
            # AreaFit screen: adjust vLine range (closer/further from center)
            if (self.window.vline1 is not None and self.window.vline2 is not None and
                    hasattr(self.window, 'vline_center') and self.window.vline_center is not None):

                save_state(self.window)

                # Get current positions
                vline1_x = self.window.vline1.get_xdata()[0]
                vline2_x = self.window.vline2.get_xdata()[0]
                center_x = self.window.vline_center.get_xdata()[0]

                # Calculate current half-gap
                current_gap = abs(vline1_x - vline2_x)
                half_gap = current_gap / 2

                # Adjust gap based on scroll direction
                delta = 0.5 if event.step > 0 else -0.5
                new_half_gap = max(half_gap + delta, 0.5)  # Minimum gap of 1.0 eV

                # Update vline positions around center
                new_vline1_x = center_x - new_half_gap
                new_vline2_x = center_x + new_half_gap

                # Update vlines
                self.window.vline1.set_xdata([new_vline1_x, new_vline1_x])
                self.window.vline2.set_xdata([new_vline2_x, new_vline2_x])

                # Update range controls if they exist
                if hasattr(self.window.background_window, 'min_range_text') and hasattr(self.window.background_window, 'max_range_text'):
                    min_pos = min(new_vline1_x, new_vline2_x)
                    max_pos = max(new_vline1_x, new_vline2_x)
                    self.window.background_window.min_range_text.SetValue(f"{min_pos:.2f}")
                    self.window.background_window.max_range_text.SetValue(f"{max_pos:.2f}")

                # Update text labels
                if hasattr(self.window.background_window, 'update_vline_text_labels'):
                    self.window.background_window.update_vline_text_labels()

                self.window.canvas.draw_idle()
            return

        # Check if Fitting screen is active and a peak is selected
        elif (self.window.peak_fitting_tab_selected and self.window.selected_peak_index is not None):
            # print("Mouse wheel detected in Peak Fitting screen with a selected peak")
            # # Give focus to the canvas parent frame
            # if hasattr(self.window, 'right_frame'):
            #     self.window.right_frame.SetFocus()

            # Give focus to the canvas so it can receive wheel events
            self.window.canvas.SetFocus()

            # Fitting screen: adjust selected peak width
            save_state(self.window)
            delta = 0.05 if event.step > 0 else -0.05
            row = self.window.selected_peak_index * 2
            fitting_model = self.window.peak_params_grid.GetCellValue(row, 13)
            if fitting_model == "SingleEntity":
                # For SingleEntity, scale the area (vertical resize)
                current_area = float(self.window.peak_params_grid.GetCellValue(row, 6))
                area_delta = 50 if event.step > 0 else -50  # Larger steps for area
                new_area = max(current_area + area_delta, 1.0)
                self.window.peak_params_grid.SetCellValue(row, 6, f"{new_area:.2f}")

                # Store scale factor in sigma column for reference
                sheet_name = self.window.sheet_combobox.GetValue()
                peaks_dict = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                peak_letter = self.window.peak_params_grid.GetCellValue(row, 1)

                for peak_name, data in peaks_dict.items():
                    if data.get('Fitting Model') == 'SingleEntity':
                        position = float(self.window.peak_params_grid.GetCellValue(row, 2))
                        if abs(data.get('Position', 0) - position) < 0.01:
                            # Update Area in Data structure
                            data['Area'] = new_area
                            break
            elif fitting_model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                 "Voigt (Area, L/G, \u03c3, S)"]:
                current_sigma = float(self.window.peak_params_grid.GetCellValue(row, 7))
                new_sigma = max(current_sigma + delta, 0.1)

                self.window.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")

                lg_ratio = float(self.window.peak_params_grid.GetCellValue(row, 5))
                new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                self.window.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
            else:
                current_fwhm = float(self.window.peak_params_grid.GetCellValue(row, 4))
                new_fwhm = max(current_fwhm + delta, 0.1)
                self.window.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")
            if fitting_model != "SingleEntity":
                self.window.recalculate_peak_area(self.window.selected_peak_index)
                self.window.update_linked_fwhm_recursive(self.window.selected_peak_index,
                                                         new_sigma if fitting_model.startswith("Voigt") else new_fwhm)

            self.window.recalculate_peak_area(self.window.selected_peak_index)
            self.window.update_linked_fwhm_recursive(self.window.selected_peak_index,
                                                     new_sigma if fitting_model.startswith("Voigt") else new_fwhm)
            self.window.clear_and_replot()
            self.window.peak_manipulation.highlight_selected_peak()
            self.window.canvas.draw_idle()
            return

        # For all other cases, do nothing (core level changing removed)

        # Refresh vline text labels after mouse wheel zoom
        self.window.refresh_vline_text_labels()

    def on_motion(self, event):

        # Handle center vline dragging motion
        if hasattr(self, 'center_drag_active') and self.center_drag_active and event.xdata is not None:
            # Set the new center directly to the mouse position
            new_center = event.xdata

            # Calculate new vline positions maintaining the gap
            new_vline1_x = new_center - self.vline_gap / 2
            new_vline2_x = new_center + self.vline_gap / 2

            # Update all vline positions
            self.window.vline1.set_xdata([new_vline1_x, new_vline1_x])
            self.window.vline2.set_xdata([new_vline2_x, new_vline2_x])
            self.window.vline_center.set_xdata([new_center, new_center])

            # Update text labels and area detection
            if hasattr(self.window, 'background_window') and self.window.background_window:
                # Call AreaFit_Screen's update method which includes auto-detection
                if hasattr(self.window.background_window, 'update_vline_text_labels'):
                    self.window.background_window.update_vline_text_labels()
            elif hasattr(self.window, 'update_vline_text_labels'):
                # Fallback to main window method
                self.window.update_vline_text_labels()

            # Update center text
            if hasattr(self.window, 'vline_center_text') and self.window.vline_center_text:
                y_pos = self.window.vline_center_text.get_position()[1]
                self.window.vline_center_text.set_position((new_center, y_pos))
                self.window.vline_center_text.set_text(f'{new_center:.2f}')

            # Update range controls
            if hasattr(self.window, 'background_window') and self.window.background_window:
                min_pos = min(new_vline1_x, new_vline2_x)
                max_pos = max(new_vline1_x, new_vline2_x)
                try:
                    self.window.background_window.updating_range_controls = True
                    self.window.background_window.min_range_text.SetValue(f"{min_pos:.2f}")
                    self.window.background_window.max_range_text.SetValue(f"{max_pos:.2f}")
                    self.window.background_window.updating_range_controls = False
                except:
                    pass

            self.window.canvas.draw_idle()
            return

        # CTRL+DRAG MOTION HANDLING - MUST BE FIRST
        elif self.ctrl_drag_active and event.inaxes:
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
            if x_click is None:
                return

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

        # Handle green vline movement
        elif (event.inaxes and hasattr(self.window, 'moving_vline_type') and
              self.window.moving_vline_type == 'green' and
              hasattr(self.window, 'moving_vline') and self.window.moving_vline is not None):
            x_click = event.xdata
            self.window.moving_vline.set_xdata([x_click])

            # Update text label
            from libraries.Widgets_Toolbars import update_green_vline_text_label
            update_green_vline_text_label(self.window)

            self.window.canvas.draw_idle()
            return
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

                    if hasattr(self.window, 'background_method') and self.window.background_method:
                        core_level_data['Background']['Bkg Type'] = self.window.background_method

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
                elif method == "Arctan-XAS":
                    current_background = BackgroundCalculations.calculate_adaptive_arctan_background(
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
            # Clean up center vline
            if hasattr(self.window, 'vline_center') and self.window.vline_center is not None:
                try:
                    self.window.vline_center.remove()
                except:
                    pass
                self.vline_center = None
            if hasattr(self.window, 'vline_center_text') and self.window.vline_center_text is not None:
                try:
                    self.window.vline_center_text.remove()
                except:
                    pass
                self.window.vline_center_text = None

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

    def on_release(self, event):

        # Handle center vline drag release
        if hasattr(self, 'center_drag_active') and self.center_drag_active:
            self.center_drag_active = False
            self.vline_gap = 0.0
            self.center_drag_reference_pos = 0.0
            self.window.moving_vline = None

            # Update data structure with new vline positions
            # ONLY when Fitting_Screen is active, NOT when AreaFit_Screen is active
            if (self.window.vline1 is not None and self.window.vline2 is not None and
                    not (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected) and
                    not (hasattr(self.window, 'background_window') and self.window.background_window is not None)):
                print("Center-drag ended, updating background vLine positions...")
                vline1_x = self.window.vline1.get_xdata()[0]
                vline2_x = self.window.vline2.get_xdata()[0]

                sheet_name = self.window.sheet_combobox.GetValue()
                if sheet_name in self.window.Data['Core levels']:
                    print(f"Updating background vLine positions in window.data for sheet: {sheet_name}")
                    if 'Background' not in self.window.Data['Core levels'][sheet_name]:
                        self.window.Data['Core levels'][sheet_name]['Background'] = {}

                    # Update background section
                    self.window.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = float(min(vline1_x, vline2_x))
                    self.window.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(max(vline1_x, vline2_x))

                    # Save current background method from UI
                    if hasattr(self.window, 'background_method') and self.window.background_method:
                        self.window.Data['Core levels'][sheet_name]['Background']['Bkg Type'] = self.window.background_method

                    # Update all peaks in Fitting section if they exist
                    if ('Fitting' in self.window.Data['Core levels'][sheet_name] and
                            'Peaks' in self.window.Data['Core levels'][sheet_name]['Fitting']):

                        print("Updating all peak background parameters to match new vLine positions...")

                        peaks = self.window.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                        for peak_label, peak_data in peaks.items():
                            peak_data['Bkg Type'] = self.window.background_method
                            peak_data['Bkg Low'] = float(min(vline1_x, vline2_x))
                            peak_data['Bkg High'] = float(max(vline1_x, vline2_x))


                        # Update peak fitting grid if it exists and has data
                        # ONLY when Fitting_Screen is active, NOT when AreaFit_Screen is active
                        if (hasattr(self.window, 'peak_params_grid') and
                                self.window.peak_params_grid.GetNumberRows() > 0 and
                                not (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected) and
                                not (hasattr(self.window, 'background_window') and self.window.background_window is not None)):
                            print("4. Updating peak fitting grid background columns...")

                            num_peaks = self.window.peak_params_grid.GetNumberRows() // 2
                            print(f" 5.  - Number of peaks to update: {num_peaks}")
                            for i in range(num_peaks):
                                print(f"   - Updating peak {i + 1} background info")
                                row = i * 2
                                # Update grid columns: 14=Bkg Type, 15=Bkg Low, 16=Bkg High
                                self.window.peak_params_grid.SetCellValue(row, 14, self.window.background_method)
                                self.window.peak_params_grid.SetCellValue(row, 15, f"{overall_bg_low:.2f}")
                                self.window.peak_params_grid.SetCellValue(row, 16, f"{overall_bg_high:.2f}")

            save_state(self.window)

            # Clean up handlers
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            if hasattr(self.window, 'drag_mode'):
                self.window.drag_mode = False
            if hasattr(self.window, 'show_hide_vlines'):
                self.window.show_hide_vlines()

            return
        elif self.ctrl_drag_active:
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
            if hasattr(self.window, 'drag_mode'):
                self.window.drag_mode = False
            if hasattr(self.window, 'show_hide_vlines'):
                self.window.show_hide_vlines()


            return  # CRITICAL: Exit here to prevent normal release handling

            # Handle green vline release
        elif (hasattr(self.window, 'moving_vline_type') and
              self.window.moving_vline_type == 'green'):
            # Reset green vline movement state
            self.window.moving_vline_type = None
            self.window.moving_vline = None

            # Clean up handlers
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')

            return

        elif self.window.moving_vline is not None:
            print("1. Single vline drag ended, updating background vLine positions...")
            # Store which vline was moved before resetting to None
            moved_vline = self.window.moving_vline

            # Save state after vline movement and background update
            save_state(self.window)

            # Update U2-Tougaard control if Fitting Screen is open and U2 is selected
            if (hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None and
                    hasattr(self.window.fitting_window, 'method_combobox')):
                if self.window.fitting_window.method_combobox.GetValue() == "U2-Tougaard":
                    wx.CallAfter(self.window.fitting_window.update_u2_tougaard_control)

            # Update VBM controls if VBM window is open
            self.update_vbm_controls_from_vlines()

            # Use the correct variable names to disconnect events
            if hasattr(self.window, 'motion_cid'):
                self.window.canvas.mpl_disconnect(self.window.motion_cid)
                delattr(self.window, 'motion_cid')
            if hasattr(self.window, 'release_cid'):
                self.window.canvas.mpl_disconnect(self.window.release_cid)
                delattr(self.window, 'release_cid')
            if hasattr(self.window, 'drag_mode'):
                self.window.drag_mode = False
            # if hasattr(self.window, 'show_hide_vlines'):
            #     self.window.show_hide_vlines()

            # Reset the moving vline to None
            self.window.moving_vline = None

            # Update data structure with background method and overall range from recorded ranges
            sheet_name = self.window.sheet_combobox.GetValue()
            if sheet_name in self.window.Data['Core levels']:
                core_level_data = self.window.Data['Core levels'][sheet_name]

                if 'Background' not in core_level_data:
                    core_level_data['Background'] = {}

                # Save current background method from UI
                if hasattr(self.window, 'background_method') and self.window.background_method:
                    core_level_data['Background']['Bkg Type'] = self.window.background_method

                # Get overall background range from all recorded ranges
                overall_bg_low = None
                overall_bg_high = None

                if (hasattr(self.window, 'fitting_window') and self.window.fitting_window is not None and
                        hasattr(self.window.fitting_window, 'get_overall_background_range')):
                    overall_bg_low, overall_bg_high = self.window.fitting_window.get_overall_background_range()
                else:
                    # Fallback: get from existing background data or vline positions
                    bg_low = core_level_data['Background'].get('Bkg Low')
                    bg_high = core_level_data['Background'].get('Bkg High')
                    if bg_low is not None and bg_high is not None:
                        overall_bg_low = min(bg_low, bg_high)
                        overall_bg_high = max(bg_low, bg_high)
                    elif self.window.vline1 is not None and self.window.vline2 is not None:
                        vline1_x = self.window.vline1.get_xdata()[0]
                        vline2_x = self.window.vline2.get_xdata()[0]
                        overall_bg_low = min(vline1_x, vline2_x)
                        overall_bg_high = max(vline1_x, vline2_x)

                # Update background section with overall range
                # ONLY when Fitting_Screen is active, NOT when AreaFit_Screen is active
                if (overall_bg_low is not None and overall_bg_high is not None and
                        not (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected) and
                        not (hasattr(self.window, 'background_window') and self.window.background_window is not None)):
                    print(f"2. Setting overall background range: {overall_bg_low:.2f} - {overall_bg_high:.2f} eV")
                    core_level_data['Background']['Bkg Low'] = float(overall_bg_low)
                    core_level_data['Background']['Bkg High'] = float(overall_bg_high)

                    # Update all peaks in Fitting section if they exist
                    if ('Fitting' in core_level_data and 'Peaks' in core_level_data['Fitting']):
                        print("3. Updating all peak background parameters to match new overall range...")
                        peaks = core_level_data['Fitting']['Peaks']
                        for peak_label, peak_data in peaks.items():
                            peak_data['Bkg Type'] = self.window.background_method
                            peak_data['Bkg Low'] = float(overall_bg_low)
                            peak_data['Bkg High'] = float(overall_bg_high)


                        # Update peak fitting grid if it exists and has data
                        # ONLY when Fitting_Screen is active, NOT when AreaFit_Screen is active
                        if (hasattr(self.window, 'peak_params_grid') and
                                self.window.peak_params_grid.GetNumberRows() > 0 and
                                not (hasattr(self.window, 'area_tab_selected') and self.window.area_tab_selected) and
                                not (hasattr(self.window, 'background_window') and self.window.background_window is not None)):
                            print("4. Updating peak fitting grid background columns...")

                            num_peaks = self.window.peak_params_grid.GetNumberRows() // 2
                            print(f" 5.  - Number of peaks to update: {num_peaks}")
                            for i in range(num_peaks):
                                print(f"   - Updating peak {i + 1} background info")
                                row = i * 2
                                # Update grid columns: 14=Bkg Type, 15=Bkg Low, 16=Bkg High
                                self.window.peak_params_grid.SetCellValue(row, 14, self.window.background_method)
                                self.window.peak_params_grid.SetCellValue(row, 15, f"{overall_bg_low:.2f}")
                                self.window.peak_params_grid.SetCellValue(row, 16, f"{overall_bg_high:.2f}")

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

    def open_profile_editor(self):
        """Open profile editor window"""
        from libraries.ViewMenu.ProfileEditor import ProfileEditWindow

        profile_editor = ProfileEditWindow(self.window)
        profile_editor.Show()

    def on_right_click(self, event):
        if event.button == 3:
            import os
            import tempfile
            from libraries.FileMenu.Save import copy_all_peak_parameters, paste_all_peak_parameters, copy_core_level, \
                paste_core_level

            menu = wx.Menu()

            sheet_name = self.window.sheet_combobox.GetValue()

            # Different menu for zzProfile sheets
            if sheet_name.startswith('zzProfile'):
                # Regular menu for normal sheets
                zoom_in = menu.Append(-1, "Zoom In")
                zoom_out = menu.Append(-1, "Zoom Out")

                menu.AppendSeparator()

                # Line style options for profiles
                increase_linewidth = menu.Append(wx.ID_ANY, "Increase Linewidth")
                decrease_linewidth = menu.Append(wx.ID_ANY, "Decrease Linewidth")

                menu.AppendSeparator()

                toggle_lines = menu.Append(wx.ID_ANY, "Show/Hide Lines")

                menu.AppendSeparator()

                # Interpolation submenu
                interp_menu = wx.Menu()
                linear_interp = interp_menu.Append(wx.ID_ANY, "Linear (Default)")
                cubic_interp = interp_menu.Append(wx.ID_ANY, "Cubic Spline (Smooth)")
                markers_only = interp_menu.Append(wx.ID_ANY, "Markers Only (No Lines)")
                menu.AppendSubMenu(interp_menu, "Interpolation Style")

                menu.AppendSeparator()

                info = menu.Append(wx.ID_ANY, "Edit Profile Data")

                # Bind profile-specific events
                self.window.Bind(wx.EVT_MENU, self.window.on_zoom_in_tool, zoom_in)
                self.window.Bind(wx.EVT_MENU, self.window.on_zoom_out, zoom_out)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.change_profile_linewidth(1), increase_linewidth)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.change_profile_linewidth(-1), decrease_linewidth)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.toggle_profile_lines(), toggle_lines)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_profile_interpolation('linear'), linear_interp)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_profile_interpolation('cubic'), cubic_interp)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_profile_interpolation('markers'), markers_only)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.open_profile_editor(), info)

            else:
                # Regular menu for normal sheets
                zoom_in = menu.Append(-1, "Zoom In")
                zoom_out = menu.Append(-1, "Zoom Out")

                menu.AppendSeparator()

                copy = menu.Append(-1, "Copy Core Level")
                paste = menu.Append(-1, "Paste Core Level")

                clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_corelevels_clipboard.json')
                has_clipboard_data = os.path.exists(clipboard_file)

                menu.AppendSeparator()

                copy_peak_table = menu.Append(-1, "Copy Peak Table")
                paste_peak_table = menu.Append(-1, "Paste Peak Table")

                # Heatmap submenu
                menu.AppendSeparator()

                # Create main Style menu
                style_menu = wx.Menu()

                # ========== HeatMap submenu ==========
                heatmap_menu = wx.Menu()

                # Colormap submenu
                colormap_submenu = wx.Menu()

                # Sequential colormaps
                seq_menu = wx.Menu()
                seq_maps = ['viridis', 'plasma', 'inferno', 'magma', 'cividis', 'twilight', 'rocket', 'mako']
                for cmap in seq_maps:
                    item = seq_menu.Append(wx.ID_ANY, cmap)
                    self.window.Bind(wx.EVT_MENU, lambda evt, c=cmap: self.window.set_heatmap_colormap(c), item)
                colormap_submenu.AppendSubMenu(seq_menu, "Sequential")

                # Diverging colormaps
                div_menu = wx.Menu()
                div_maps = ['RdBu', 'RdYlBu', 'coolwarm', 'bwr', 'seismic', 'PiYG', 'PRGn']
                for cmap in div_maps:
                    item = div_menu.Append(wx.ID_ANY, cmap)
                    self.window.Bind(wx.EVT_MENU, lambda evt, c=cmap: self.window.set_heatmap_colormap(c), item)
                colormap_submenu.AppendSubMenu(div_menu, "Diverging")

                # Qualitative colormaps
                qual_menu = wx.Menu()
                qual_maps = ['tab10', 'Set1', 'Set2', 'Set3', 'Paired', 'Accent']
                for cmap in qual_maps:
                    item = qual_menu.Append(wx.ID_ANY, cmap)
                    self.window.Bind(wx.EVT_MENU, lambda evt, c=cmap: self.window.set_heatmap_colormap(c), item)
                colormap_submenu.AppendSubMenu(qual_menu, "Qualitative")

                # Other colormaps
                other_menu = wx.Menu()
                other_maps = ['jet', 'hot', 'cool', 'spring', 'summer', 'autumn', 'winter', 'gray', 'bone', 'copper']
                for cmap in other_maps:
                    item = other_menu.Append(wx.ID_ANY, cmap)
                    self.window.Bind(wx.EVT_MENU, lambda evt, c=cmap: self.window.set_heatmap_colormap(c), item)
                colormap_submenu.AppendSubMenu(other_menu, "Other")

                heatmap_menu.AppendSubMenu(colormap_submenu, "Colormap")

                # Smooth options
                smooth_submenu = wx.Menu()
                smooth_light = smooth_submenu.Append(wx.ID_ANY, "Light (sigma=0.5)")
                smooth_medium = smooth_submenu.Append(wx.ID_ANY, "Medium (sigma=1.0)")
                smooth_heavy = smooth_submenu.Append(wx.ID_ANY, "Heavy (sigma=2.0)")
                smooth_custom = smooth_submenu.Append(wx.ID_ANY, "Custom...")

                self.window.Bind(wx.EVT_MENU, lambda evt: self.window.smooth_heatmap(0.5), smooth_light)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.window.smooth_heatmap(1.0), smooth_medium)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.window.smooth_heatmap(2.0), smooth_heavy)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.window.smooth_heatmap_custom(), smooth_custom)

                heatmap_menu.AppendSubMenu(smooth_submenu, "Smooth")

                # Reset option
                reset_item = heatmap_menu.Append(wx.ID_ANY, "Reset to Original")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.window.reset_heatmap(), reset_item)

                # Add HeatMap to Style menu
                heatmap_submenu = style_menu.AppendSubMenu(heatmap_menu, "HeatMap")

                # # Enable only if heatmap is active
                # heatmap_active = hasattr(self.window, 'heatmap_data') and self.window.heatmap_data is not None
                # heatmap_submenu.Enable(heatmap_active)

                # ========== EDX~Plot submenu ==========
                edx_menu = wx.Menu()

                # Default style: Black line, black labels
                default_item = edx_menu.Append(wx.ID_ANY, "Default (Black)")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_edx_plot_style('black'), default_item)

                # Blue background, yellow filled, yellow labels
                blue_yellow_item = edx_menu.Append(wx.ID_ANY, "Blue/Yellow")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_edx_plot_style('blue_yellow'), blue_yellow_item)

                # White background, red filled, black labels
                white_red_item = edx_menu.Append(wx.ID_ANY, "White/Red")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_edx_plot_style('white_red'), white_red_item)

                # White background, blue filled, black labels
                white_blue_item = edx_menu.Append(wx.ID_ANY, "White/Blue")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.set_edx_plot_style('white_blue'), white_blue_item)

                edx_submenu = style_menu.AppendSubMenu(edx_menu, "EDX~Plot")

                # # Enable only if current sheet is EDX~Plot
                # edx_active = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')
                # edx_submenu.Enable(edx_active)

                # ========== Multiplot submenu ==========
                multiplot_menu = wx.Menu()

                # Color palette submenu
                palette_menu = wx.Menu()

                # Categorical palettes submenu
                categorical_menu = wx.Menu()
                categorical_palettes = ['tab10', 'tab20', 'tab20b', 'tab20c', 'Set1', 'Set2', 'Set3',
                                        'Paired', 'Dark2', 'Pastel1', 'Pastel2', 'Accent']
                for palette in categorical_palettes:
                    item = categorical_menu.Append(wx.ID_ANY, palette)
                    self.window.Bind(wx.EVT_MENU, lambda evt, p=palette: self.on_change_multiplot_palette(evt, p), item)
                palette_menu.AppendSubMenu(categorical_menu, "Categorical")

                # Gradients submenu
                gradients_menu = wx.Menu()

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Blue")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Blues_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Green")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Greens_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Red")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Reds_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Purple")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Purples_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Pink")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'RdPu_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Orange")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Oranges_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Yellow")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'YlOrBr_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Brown")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'YlOrBr_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Grey")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'Greys_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Dark Grey")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'gray_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Black")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'binary_r'), item)

                item = gradients_menu.Append(wx.ID_ANY, "Gradient Cyan")
                self.window.Bind(wx.EVT_MENU, lambda evt: self.on_change_multiplot_palette(evt, 'GnBu_r'), item)

                palette_menu.AppendSubMenu(gradients_menu, "Gradients")

                # Rainbow/Multi-color submenu
                rainbow_menu = wx.Menu()
                rainbow_palettes = ['viridis', 'plasma', 'inferno', 'magma', 'turbo', 'rainbow', 'jet', 'hsv']
                for palette in rainbow_palettes:
                    item = rainbow_menu.Append(wx.ID_ANY, palette)
                    self.window.Bind(wx.EVT_MENU, lambda evt, p=palette: self.on_change_multiplot_palette(evt, p), item)
                palette_menu.AppendSubMenu(rainbow_menu, "Rainbow/Multi-color")

                multiplot_menu.AppendSubMenu(palette_menu, "Color Palette")

                # Line width submenu
                width_menu = wx.Menu()
                widths = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0]

                for width in widths:
                    item = width_menu.Append(wx.ID_ANY, f"{width:.1f}")
                    self.window.Bind(wx.EVT_MENU, lambda evt, w=width: self.on_change_multiplot_linewidth(evt, w), item)

                multiplot_menu.AppendSubMenu(width_menu, "Line Width")

                # Legend columns submenu
                legend_menu = wx.Menu()
                legend_cols = [1, 2, 3, 4]

                for ncol in legend_cols:
                    item = legend_menu.Append(wx.ID_ANY, f"{ncol} columns")
                    self.window.Bind(wx.EVT_MENU, lambda evt, n=ncol: self.on_change_multiplot_legend_ncol(evt, n), item)

                multiplot_menu.AppendSubMenu(legend_menu, "Legend Columns")

                # Max legend items submenu
                legend_items_menu = wx.Menu()
                legend_items = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55]

                for max_items in legend_items:
                    item = legend_items_menu.Append(wx.ID_ANY, f"{max_items} items max")
                    self.window.Bind(wx.EVT_MENU, lambda evt, m=max_items: self.on_change_multiplot_max_legend_items(evt, m), item)

                multiplot_menu.AppendSubMenu(legend_items_menu, "Max Legend Items")

                multiplot_submenu = style_menu.AppendSubMenu(multiplot_menu, "Multiplot")

                # # Enable only if file_manager is open
                # multiplot_active = hasattr(self.window, 'file_manager') and self.window.file_manager is not None
                # multiplot_submenu.Enable(multiplot_active)

                # Add Style menu to main menu
                menu.AppendSubMenu(style_menu, "Style")

                menu.AppendSeparator()
                edit_data = menu.Append(wx.ID_ANY, "Edit Data")
                info = menu.Append(wx.ID_ANY, "Info")

                peak_clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')

                has_peak_clipboard_data = os.path.exists(peak_clipboard_file)
                has_rows = self.window.peak_params_grid.GetNumberRows() > 0

                paste.Enable(has_clipboard_data)
                paste_peak_table.Enable(has_peak_clipboard_data)
                copy_peak_table.Enable(has_rows)

                self.window.Bind(wx.EVT_MENU, self.window.on_zoom_in_tool, zoom_in)
                self.window.Bind(wx.EVT_MENU, self.window.on_zoom_out, zoom_out)
                self.window.Bind(wx.EVT_MENU, lambda evt: copy_core_level(self.window), copy)
                self.window.Bind(wx.EVT_MENU, lambda evt: paste_core_level(self.window), paste)
                self.window.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self.window), copy_peak_table)
                self.window.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self.window), paste_peak_table)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.open_edit_data_window(), edit_data)
                self.window.Bind(wx.EVT_MENU, lambda evt: self.open_experimental_description(), info)

            self.window.PopupMenu(menu)
            menu.Destroy()

    def set_edx_plot_style_OLD(self, style):
        """Apply EDX plot style to current EDX~Plot sheet"""
        sheet_name = self.window.sheet_combobox.GetValue()
        if not (sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')):
            return

        # Store style preference
        if 'Core levels' in self.window.Data and sheet_name in self.window.Data['Core levels']:
            self.window.Data['Core levels'][sheet_name]['_EDX_style'] = style

        # Get current data
        if sheet_name not in self.window.Data['Core levels']:
            return

        sheet_data = self.window.Data['Core levels'][sheet_name]
        if 'B.E.' not in sheet_data or 'Raw Data' not in sheet_data:
            return

        import numpy as np
        energy = np.array(sheet_data['B.E.'])
        intensity = np.array(sheet_data['Raw Data'])

        # Clear the plot
        self.window.ax.clear()

        # Apply style-specific plotting
        if style == 'black':
            # Default style: Black line, black labels, white background
            self.window.figure.patch.set_facecolor('white')
            self.window.ax.set_facecolor('white')
            self.window.ax.plot(energy, intensity, 'k-', linewidth=0.8)
            self.window.ax.set_xlabel('Energy (keV)', color='black')
            self.window.ax.set_ylabel('Counts', color='black')
            self.window.ax.tick_params(colors='black', labelcolor='black')
            for spine in self.window.ax.spines.values():
                spine.set_edgecolor('black')
            label_color = 'black'

        elif style == 'blue_yellow':
            # Blue background with yellow filled line and yellow labels
            self.window.figure.patch.set_facecolor('#1e3a5f')
            self.window.ax.set_facecolor('#1e3a5f')
            self.window.ax.fill_between(energy, intensity, color='yellow', alpha=0.8)
            self.window.ax.plot(energy, intensity, 'yellow', linewidth=1.2)
            self.window.ax.set_xlabel('Energy (keV)', color='yellow')
            self.window.ax.set_ylabel('Counts', color='yellow')
            self.window.ax.tick_params(colors='yellow', labelcolor='yellow')
            for spine in self.window.ax.spines.values():
                spine.set_edgecolor('yellow')
            label_color = 'yellow'

        elif style == 'white_red':
            # White background with red filled line and black labels
            self.window.figure.patch.set_facecolor('white')
            self.window.ax.set_facecolor('white')
            self.window.ax.fill_between(energy, intensity, color='red', alpha=0.6)
            self.window.ax.plot(energy, intensity, 'red', linewidth=1.0)
            self.window.ax.set_xlabel('Energy (keV)', color='black')
            self.window.ax.set_ylabel('Counts', color='black')
            self.window.ax.tick_params(colors='black', labelcolor='black')
            for spine in self.window.ax.spines.values():
                spine.set_edgecolor('black')
            label_color = 'black'

        # Set title
        self.window.ax.set_title('EDX Spectrum')

        # Set Y-axis to scientific format
        self.window.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

        # Add peak labels with appropriate color
        if hasattr(self.window, 'edx_window') and self.window.edx_window:
            try:
                # Call add_peak_labels
                self.window.edx_window.add_peak_labels(self.window.ax, energy, intensity)

                # Recolor all text labels to match style
                for text in self.window.ax.texts:
                    text.set_color(label_color)
                    text.set_fontweight('bold')
            except Exception as e:
                print(f"Could not add peak labels: {e}")

        # Set display limits
        display_x_max = sheet_data.get('_EDX_display_max', 20)
        self.window.ax.set_xlim(0, display_x_max)
        self.window.ax.set_ylim(np.min(intensity) * 0.95, np.max(intensity) * 1.1)

        self.window.canvas.draw()

    def set_edx_plot_style(self, style):
        """Apply EDX plot style to ALL EDX~Plot sheets"""
        # Store style preference globally in window
        self.window.edx_plot_style = style

        # Apply to ALL EDX~Plot sheets in Data
        if 'Core levels' in self.window.Data:
            for sheet_name in self.window.Data['Core levels']:
                if sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot'):
                    self.window.Data['Core levels'][sheet_name]['_EDX_style'] = style

        # Save to config
        self.window.save_config()

        # Replot current sheet if it's an EDX~Plot
        current_sheet = self.window.sheet_combobox.GetValue()
        if current_sheet == 'EDX~Plot' or current_sheet.startswith('EDX~Plot'):
            from libraries.Plot_Operations import PlotManager
            if hasattr(self.window, 'plot_manager'):
                self.window.plot_manager.plot_edx_data(self.window, current_sheet)

    def open_edit_data_window(self):
        """Open the Edit Data window for the current sheet"""
        from libraries.ViewMenu.EditDataWindow import EditDataWindow

        # Check if window already exists and close it
        if hasattr(self.window, 'edit_data_window') and self.window.edit_data_window is not None:
            try:
                self.window.edit_data_window.Close()
                self.window.edit_data_window.Destroy()
            except:
                pass

        # Create new window
        self.window.edit_data_window = EditDataWindow(self.window)
        self.window.edit_data_window.Show()

    def open_experimental_description(self):
        """Open the Experimental Description window or EDX Info for the current sheet"""
        sheet_name = self.window.sheet_combobox.GetValue()

        # Check if this is an EDX plot sheet
        if sheet_name.startswith('EDX~Plot'):
            self.show_edx_info(sheet_name)
            return

        from libraries.ViewMenu.FileManager import ExperimentalDescriptionWindow

        # Check if window already exists and close it
        if hasattr(self.window, 'experimental_description_window') and self.window.experimental_description_window is not None:
            try:
                self.window.experimental_description_window.Close()
                self.window.experimental_description_window.Destroy()
            except:
                pass

        # Create and show the window
        self.window.experimental_description_window = ExperimentalDescriptionWindow(self.window, sheet_name)
        self.window.experimental_description_window.Show()

    def show_edx_info(self, sheet_name):
        """Show EDX plot information in a dialog"""
        import wx

        if sheet_name not in self.window.Data.get('Core levels', {}):
            wx.MessageBox("No data available for this sheet.", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        sheet_data = self.window.Data['Core levels'][sheet_name]

        # Build info text
        info_lines = []
        info_lines.append(f"Sheet: {sheet_name}")
        info_lines.append("=" * 40)

        # Save time
        save_time = sheet_data.get('_EDX_save_time', 'Unknown')
        if save_time != 'Unknown':
            try:
                from datetime import datetime
                dt = datetime.fromisoformat(save_time)
                save_time = dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                pass
        info_lines.append(f"Saved: {save_time}")
        info_lines.append("")

        # Selection info
        selection_info = sheet_data.get('_EDX_selection')
        if selection_info:
            sel_type = selection_info.get('type', 'Unknown')
            info_lines.append(f"Selection Type: {sel_type.capitalize()}")
            info_lines.append("-" * 30)

            if sel_type == 'point':
                x = selection_info.get('x', 0)
                y = selection_info.get('y', 0)
                size = selection_info.get('size', 1)
                info_lines.append(f"Location: ({x}, {y})")
                info_lines.append(f"Size: {size}×{size} pixels")

            elif sel_type == 'line':
                x1, y1 = selection_info.get('x1', 0), selection_info.get('y1', 0)
                x2, y2 = selection_info.get('x2', 0), selection_info.get('y2', 0)
                length = np.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                info_lines.append(f"Start: ({x1}, {y1})")
                info_lines.append(f"End: ({x2}, {y2})")
                info_lines.append(f"Length: {length:.1f} pixels")

            elif sel_type == 'rectangle':
                cx = selection_info.get('center_x', 0)
                cy = selection_info.get('center_y', 0)
                w = selection_info.get('width', 0)
                h = selection_info.get('height', 0)
                angle = selection_info.get('angle', 0)
                info_lines.append(f"Center: ({cx:.1f}, {cy:.1f})")
                info_lines.append(f"Size: {w:.1f} × {h:.1f} pixels")
                info_lines.append(f"Rotation: {angle:.1f}°")
                info_lines.append(f"Area: {w * h:.1f} pixels²")

            elif sel_type == 'area':
                x1, y1 = selection_info.get('x1', 0), selection_info.get('y1', 0)
                x2, y2 = selection_info.get('x2', 0), selection_info.get('y2', 0)
                w, h = abs(x2 - x1), abs(y2 - y1)
                info_lines.append(f"Top-Left: ({min(x1, x2)}, {min(y1, y2)})")
                info_lines.append(f"Bottom-Right: ({max(x1, x2)}, {max(y1, y2)})")
                info_lines.append(f"Size: {w} × {h} pixels")
                info_lines.append(f"Area: {w * h} pixels²")

        info_lines.append("")

        # Grid data (quantification)
        grid_data = sheet_data.get('_EDX_grid_data', [])
        if grid_data:
            info_lines.append("Quantification Results:")
            info_lines.append("-" * 30)
            info_lines.append(f"{'Element':<12} {'Energy (eV)':<12} {'Height':<10} {'Area':<12} {'Conc.(%)':<10}")
            info_lines.append("-" * 56)

            for row in grid_data:
                label = row.get('label', '')
                position = row.get('position', '')
                height = row.get('height', '')
                area = row.get('area', '')
                conc = row.get('concentration', '')
                info_lines.append(f"{label:<12} {position:<12} {height:<10} {area:<12} {conc:<10}")

        info_lines.append("")

        # Spectrum info
        energy = sheet_data.get('Energy_keV', [])
        intensity = sheet_data.get('Intensity', [])
        if len(energy) > 0 and len(intensity) > 0:
            info_lines.append("Spectrum Info:")
            info_lines.append("-" * 30)
            info_lines.append(f"Energy Range: {min(energy):.2f} - {max(energy):.2f} keV")
            info_lines.append(f"Data Points: {len(energy)}")
            info_lines.append(f"Max Intensity: {max(intensity):.2f}")

        # Create dialog
        dlg = wx.Dialog(self.window, title=f"EDX Info - {sheet_name}", size=(450, 500))
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Text control with info
        text_ctrl = wx.TextCtrl(panel, value="\n".join(info_lines),
                                style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)
        text_ctrl.SetFont(wx.Font(10, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        sizer.Add(text_ctrl, 1, wx.EXPAND | wx.ALL, 10)

        # Close button
        close_btn = wx.Button(panel, wx.ID_OK, "Close")
        sizer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(sizer)
        dlg.ShowModal()
        dlg.Destroy()

    def change_profile_linewidth(self, delta):
        """Change linewidth of profile lines"""
        sheet_name = self.window.sheet_combobox.GetValue()
        if not sheet_name.startswith('zzProfile'):
            return

        # Get or initialize linewidth
        if not hasattr(self.window, 'profile_linewidth'):
            self.window.profile_linewidth = {}

        current_width = self.window.profile_linewidth.get(sheet_name, 2.0)
        new_width = max(0.5, min(10.0, current_width + delta * 0.5))  # Range: 0.5 to 10
        self.window.profile_linewidth[sheet_name] = new_width

        # Replot
        self.window.plot_manager.plot_data(self.window)

    def toggle_profile_lines(self):
        """Toggle visibility of profile lines"""
        sheet_name = self.window.sheet_combobox.GetValue()
        if not sheet_name.startswith('zzProfile'):
            return

        # Get or initialize lines visibility
        if not hasattr(self.window, 'profile_lines_visible'):
            self.window.profile_lines_visible = {}

        current_state = self.window.profile_lines_visible.get(sheet_name, True)
        self.window.profile_lines_visible[sheet_name] = not current_state

        # Replot
        self.window.plot_manager.plot_data(self.window)

    def set_profile_interpolation(self, interp_type):
        """Set interpolation type for profile"""
        sheet_name = self.window.sheet_combobox.GetValue()
        if not sheet_name.startswith('zzProfile'):
            return

        # Store interpolation type
        if not hasattr(self.window, 'profile_interpolation'):
            self.window.profile_interpolation = {}

        self.window.profile_interpolation[sheet_name] = interp_type

        # Replot
        self.window.plot_manager.plot_data(self.window)



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
        pattern = r'([A-Z])(?=[+\-*/]|$)'
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
        pattern = r'([A-Z])(?=[+\-*/]|$|#)'
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

    def on_change_multiplot_palette(self, event, palette):
        """Change color palette for multiplot"""
        self.window.multiplot_palette = palette

        # Save to preferences
        if hasattr(self.window, 'save_config'):
            try:
                self.window.save_config()
            except Exception as e:
                print(f"Error saving palette config: {e}")

        # Refresh the multiplot if file_manager exists and has a plot
        if hasattr(self.window, 'file_manager') and self.window.file_manager:
            selected_sheets = self.window.file_manager.get_selected_sheet_names()
            if len(selected_sheets) > 1:
                # Determine which plot type to refresh based on last plot
                if hasattr(self.window, 'last_multiplot_type'):
                    if self.window.last_multiplot_type == 'F2':
                        self.window.file_manager.plot_multiple_sheets(selected_sheets)
                    elif self.window.last_multiplot_type == 'F3':
                        self.window.file_manager.plot_multiple_sheets_with_offset(selected_sheets)
                else:
                    # Default to F2 style
                    self.window.file_manager.plot_multiple_sheets(selected_sheets)

    def on_change_multiplot_linewidth(self, event, linewidth):
        """Change line width for multiplot"""
        self.window.multiplot_linewidth = linewidth

        # Save to preferences
        if hasattr(self.window, 'save_config'):
            try:
                self.window.save_config()
            except Exception as e:
                print(f"Error saving linewidth config: {e}")

        # Refresh the multiplot if file_manager exists and has a plot
        if hasattr(self.window, 'file_manager') and self.window.file_manager:
            selected_sheets = self.window.file_manager.get_selected_sheet_names()
            if len(selected_sheets) > 1:
                # Determine which plot type to refresh based on last plot
                if hasattr(self.window, 'last_multiplot_type'):
                    if self.window.last_multiplot_type == 'F2':
                        self.window.file_manager.plot_multiple_sheets(selected_sheets)
                    elif self.window.last_multiplot_type == 'F3':
                        self.window.file_manager.plot_multiple_sheets_with_offset(selected_sheets)
                else:
                    # Default to F2 style
                    self.window.file_manager.plot_multiple_sheets(selected_sheets)

    def on_change_multiplot_legend_ncol(self, event, ncol):
        """Change number of legend columns for multiplot"""
        self.window.multiplot_legend_ncol = ncol

        # Save to preferences
        if hasattr(self.window, 'save_config'):
            try:
                self.window.save_config()
            except Exception as e:
                print(f"Error saving legend ncol config: {e}")

        # Refresh the multiplot if file_manager exists and has a plot
        if hasattr(self.window, 'file_manager') and self.window.file_manager:
            selected_sheets = self.window.file_manager.get_selected_sheet_names()
            if len(selected_sheets) > 1:
                # Determine which plot type to refresh based on last plot
                if hasattr(self.window, 'last_multiplot_type'):
                    if self.window.last_multiplot_type == 'F2':
                        self.window.file_manager.plot_multiple_sheets(selected_sheets)
                    elif self.window.last_multiplot_type == 'F3':
                        self.window.file_manager.plot_multiple_sheets_with_offset(selected_sheets)
                else:
                    # Default to F2 style
                    self.window.file_manager.plot_multiple_sheets(selected_sheets)

    def on_change_multiplot_max_legend_items(self, event, max_items):
        """Change maximum number of legend items for multiplot"""
        self.window.multiplot_max_legend_items = max_items

        # Save to preferences
        if hasattr(self.window, 'save_config'):
            try:
                self.window.save_config()
            except Exception as e:
                print(f"Error saving max legend items config: {e}")

        # Refresh the multiplot if file_manager exists and has a plot
        if hasattr(self.window, 'file_manager') and self.window.file_manager:
            selected_sheets = self.window.file_manager.get_selected_sheet_names()
            if len(selected_sheets) > 1:
                # Determine which plot type to refresh based on last plot
                if hasattr(self.window, 'last_multiplot_type'):
                    if self.window.last_multiplot_type == 'F2':
                        self.window.file_manager.plot_multiple_sheets(selected_sheets)
                    elif self.window.last_multiplot_type == 'F3':
                        self.window.file_manager.plot_multiple_sheets_with_offset(selected_sheets)
                else:
                    # Default to F2 style
                    self.window.file_manager.plot_multiple_sheets(selected_sheets)

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