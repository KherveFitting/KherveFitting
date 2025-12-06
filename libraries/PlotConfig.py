# libraries/PlotConfig.py
import matplotlib.widgets as widgets
import matplotlib.pyplot as plt
import wx
import numpy as np


class PlotConfig:
    def __init__(self):
        self.plot_limits = {}
        self.original_limits = {}  # Store original limits for zoom out

    def on_zoom_in_tool(self, window):
        # Toggle zoom mode
        window.zoom_mode = not window.zoom_mode

        if window.zoom_mode:
            if window.zoom_rect:
                # Activate existing zoom rectangle
                window.zoom_rect.set_active(True)
            else:
                # Create new zoom rectangle
                window.zoom_rect = widgets.RectangleSelector(
                    window.ax,
                    lambda eclick, erelease: self.on_zoom_select(window, eclick, erelease),
                    useblit=True,
                    props=dict(facecolor='green', edgecolor='green', alpha=0.3, fill=True),
                    button=[1],
                    minspanx=5,
                    minspany=5,
                    spancoords='pixels',
                    interactive=False
                )
        else:
            # Deactivate zoom rectangle if it exists
            if window.zoom_rect:
                window.zoom_rect.set_active(False)

        # Disable drag mode if active
        if window.drag_mode:
            self.disable_drag(window)
            window.drag_mode = False

        # Redraw the canvas
        window.canvas.draw_idle()

    def on_zoom_select_OLD(self, window, eclick, erelease):
        if window.zoom_mode:
            # Extract click and release coordinates
            x1, y1 = eclick.xdata, eclick.ydata
            x2, y2 = erelease.xdata, erelease.ydata

            # Determine the selected range
            x_min, x_max = min(x1, x2), max(x1, x2)
            y_min, y_max = min(y1, y2), max(y1, y2)

            # Get current sheet name
            sheet_name = window.sheet_combobox.GetValue()
            is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper() or "Ra_" in sheet_name
            is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')
            is_edx = sheet_name == 'EDX~Plot'

            # Set x-axis limits based on energy scale and data type
            if window.energy_scale == 'KE':
                window.ax.set_xlim(min(x_max, x_min), max(x_max, x_min))
            else:
                if is_raman or is_xas or is_edx or sheet_name.startswith('zzProfile'):
                    window.ax.set_xlim(x_min, x_max)  # Normal direction for Raman and zzProfile
                else:
                    window.ax.set_xlim(max(x_max, x_min), min(x_max, x_min))  # Reverse X-axis for XPS

            # Set y-axis limits
            window.ax.set_ylim(y_min, y_max)

            # Deactivate zoom mode
            window.zoom_mode = False

            # Remove zoom rectangle if it exists
            if window.zoom_rect:
                window.zoom_rect.set_active(False)
                window.zoom_rect = None

            # Redraw the canvas to show updated plot
            window.canvas.draw_idle()

    def on_zoom_select(self, window, eclick, erelease):
        if window.zoom_mode:
            # Extract click and release coordinates
            x1, y1 = eclick.xdata, eclick.ydata
            x2, y2 = erelease.xdata, erelease.ydata

            # Determine the selected range
            x_min, x_max = min(x1, x2), max(x1, x2)
            y_min, y_max = min(y1, y2), max(y1, y2)

            # Get current sheet name
            sheet_name = window.sheet_combobox.GetValue()
            is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper() or "Ra_" in sheet_name
            is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')
            is_edx = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')

            # Set x-axis limits based on energy scale and data type
            if window.energy_scale == 'KE':
                window.ax.set_xlim(min(x_max, x_min), max(x_max, x_min))
            else:
                if is_raman or is_xas or is_edx or sheet_name.startswith('zzProfile'):
                    window.ax.set_xlim(x_min, x_max)  # Normal direction for Raman, XAS, EDX and zzProfile
                else:
                    window.ax.set_xlim(max(x_max, x_min), min(x_max, x_min))  # Reverse X-axis for XPS

            # For EDX plots, store the new X-max for all EDX sheets
            if is_edx and 'Core levels' in window.Data:
                for sname in window.Data['Core levels']:
                    if sname == 'EDX~Plot' or sname.startswith('EDX~Plot'):
                        window.Data['Core levels'][sname]['_EDX_display_max'] = x_max

            # Set y-axis limits
            window.ax.set_ylim(y_min, y_max)

            # Deactivate zoom mode
            window.zoom_mode = False

            # Remove zoom rectangle if it exists
            if window.zoom_rect:
                window.zoom_rect.set_active(False)
                window.zoom_rect = None

            # Redraw the canvas to show updated plot
            window.canvas.draw_idle()
    def on_zoom_out_OLD(self, window):
        """Handle zoom out for both single and multiple plot modes"""
        # Check if current sheet is EDX first
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name == 'EDX~Plot':
            # Handle EDX zoom out
            if 'Core levels' in window.Data and sheet_name in window.Data['Core levels']:
                sheet_data = window.Data['Core levels'][sheet_name]
                if '_EDX_min' in sheet_data and '_EDX_max' in sheet_data:
                    energy_min = sheet_data['_EDX_min']
                    energy_max = sheet_data['_EDX_max']
                    window.ax.set_xlim(energy_min, energy_max)

                    # Also reset Y limits
                    if 'Intensity' in sheet_data:
                        intensity = sheet_data['Intensity']
                        window.ax.set_ylim(0, np.max(intensity) * 1.1)

                    window.canvas.draw()
            return

        # Check if FileManager exists and has multiple sheets selected
        is_multiple_plot_mode = False

        if hasattr(window, 'file_manager') and window.file_manager:
            try:
                # Check if FileManager has multiple selected sheets
                selected_sheets = window.file_manager.get_selected_sheet_names()
                is_multiple_plot_mode = len(selected_sheets) > 1
            except:
                is_multiple_plot_mode = False

        # Alternative check: see if we have a flag set for multiple plot mode
        if hasattr(window, 'multiple_plot_mode'):
            is_multiple_plot_mode = window.multiple_plot_mode

        if is_multiple_plot_mode:
            # For multiple plots: zoom out to show all data properly
            self._zoom_out_multiple_plots(window)
        else:
            # For single plots: use standard zoom out
            sheet_name = window.sheet_combobox.GetValue()

            # Check if EDX plot
            is_edx_plot = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')

            if is_edx_plot:
                # For EDX plots, set to 0-20 keV and full intensity range
                if hasattr(window, 'y_values') and len(window.y_values) > 0:
                    window.ax.set_xlim(0, 20)
                    window.ax.set_ylim(np.min(window.y_values) * 0.95, np.max(window.y_values) * 1.1)
                else:
                    window.ax.set_xlim(0, 20)
                    window.ax.set_ylim(auto=True)

                # Store the reset X-max for all EDX~Plot sheets
                if 'Core levels' in window.Data:
                    for sname in window.Data['Core levels']:
                        if sname == 'EDX~Plot' or sname.startswith('EDX~Plot'):
                            window.Data['Core levels'][sname]['_EDX_display_max'] = 20
            else:
                # Reset plot limits to original values for non-EDX
                self.reset_plot_limits(window, sheet_name)

                # Resize plot with reset limits
                self.resize_plot(window)

        # Deactivate and remove zoom rectangle if it exists
        if window.zoom_rect:
            window.zoom_rect.set_active(False)
            window.zoom_rect = None

        # Disable zoom mode
        window.zoom_mode = False

        # Redraw the canvas
        window.canvas.draw_idle()

        # Disable drag mode if active
        if window.drag_mode:
            window.disable_drag()
            window.drag_mode = False

    def on_zoom_out(self, window):
        """Handle zoom out for both single and multiple plot modes"""
        # Get current sheet name
        sheet_name = window.sheet_combobox.GetValue()

        # Check if EDX plot (including EDX~Plot1, EDX~Plot2, etc.)
        is_edx_plot = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')

        if is_edx_plot:
            # For EDX plots, zoom out to 0-20 keV and full intensity range
            if 'Core levels' in window.Data and sheet_name in window.Data['Core levels']:
                sheet_data = window.Data['Core levels'][sheet_name]

                # Get intensity data
                if 'Intensity' in sheet_data:
                    intensity = np.array(sheet_data['Intensity'])
                    y_min = np.min(intensity) * 0.95
                    y_max = np.max(intensity) * 1.1
                elif hasattr(window, 'y_values') and len(window.y_values) > 0:
                    intensity = np.array(window.y_values)
                    y_min = np.min(intensity) * 0.95
                    y_max = np.max(intensity) * 1.1
                else:
                    y_min = 0
                    y_max = 1000

                # Set limits
                window.ax.set_xlim(0, 20)
                window.ax.set_ylim(y_min, y_max)

                # Store the reset X-max for all EDX~Plot sheets
                for sname in window.Data['Core levels']:
                    if sname == 'EDX~Plot' or sname.startswith('EDX~Plot'):
                        window.Data['Core levels'][sname]['_EDX_display_max'] = 20.00

                print(f"EDX zoom out: X=0.00-20.00 keV, Y={y_min:.2f}-{y_max:.2f}")

                window.canvas.draw_idle()

                # Deactivate and remove zoom rectangle if it exists
                if window.zoom_rect:
                    window.zoom_rect.set_active(False)
                    window.zoom_rect = None

                # Disable zoom mode
                window.zoom_mode = False

                # Disable drag mode if active
                if window.drag_mode:
                    window.disable_drag()
                    window.drag_mode = False

                return

        # Check if FileManager exists and has multiple sheets selected
        is_multiple_plot_mode = False

        if hasattr(window, 'file_manager') and window.file_manager:
            try:
                # Check if FileManager has multiple selected sheets
                selected_sheets = window.file_manager.get_selected_sheet_names()
                is_multiple_plot_mode = len(selected_sheets) > 1
            except:
                is_multiple_plot_mode = False

        # Alternative check: see if we have a flag set for multiple plot mode
        if hasattr(window, 'multiple_plot_mode'):
            is_multiple_plot_mode = window.multiple_plot_mode

        if is_multiple_plot_mode:
            # For multiple plots: zoom out to show all data properly
            self._zoom_out_multiple_plots(window)
        else:
            # For single plots: use standard zoom out
            # Reset plot limits to original values for non-EDX
            self.reset_plot_limits(window, sheet_name)

            # Resize plot with reset limits
            self.resize_plot(window)

        # Deactivate and remove zoom rectangle if it exists
        if window.zoom_rect:
            window.zoom_rect.set_active(False)
            window.zoom_rect = None

        # Disable zoom mode
        window.zoom_mode = False

        # Redraw the canvas
        window.canvas.draw_idle()

        # Disable drag mode if active
        if window.drag_mode:
            window.disable_drag()
            window.drag_mode = False

    def _zoom_out_multiple_plots(self, window):
        """Zoom out for multiple plots to show all data appropriately"""
        # Get all plotted lines
        all_lines = window.ax.lines
        excluded_labels = ['Background', 'Overall Fit', 'Residuals', 'Envelope']

        # Find data lines (excluding fit components and peaks)
        data_lines = []
        for line in all_lines:
            label = line.get_label()
            if (label not in excluded_labels and
                    not label.startswith('Peak') and
                    not label.startswith('_') and  # matplotlib internal lines
                    label != ''):
                data_lines.append(line)

        if not data_lines:
            # Fallback to standard zoom out if no data lines found
            sheet_name = window.sheet_combobox.GetValue()
            self.reset_plot_limits(window, sheet_name)
            self.resize_plot(window)
            return

        # Calculate bounds from all data lines
        all_x_data = []
        all_y_data = []

        for line in data_lines:
            x_data = line.get_xdata()
            y_data = line.get_ydata()
            if len(x_data) > 0 and len(y_data) > 0:
                all_x_data.extend(x_data)
                all_y_data.extend(y_data)

        if all_x_data and all_y_data:
            # Calculate appropriate bounds with padding
            x_min, x_max = min(all_x_data), max(all_x_data)
            y_min, y_max = min(all_y_data), max(all_y_data)

            # Add padding (10% on each side)
            x_range = x_max - x_min
            y_range = y_max - y_min

            x_padding = 0
            y_padding = y_range * 0.1 if y_range > 0 else (y_max * 0.1)



            # Apply bounds (note: X-axis is reversed for XPS data)
            window.ax.set_xlim(x_max + x_padding, x_min - x_padding)
            window.ax.set_ylim(y_min - 0.1*y_padding, y_max + y_padding)

            # Update subplot limits if it exists
            if hasattr(window, 'residuals_subplot') and window.residuals_subplot:
                window.residuals_subplot.set_xlim(x_max + x_padding, x_min - x_padding)

            # Update plot manager residuals
            if hasattr(window, 'plot_manager'):
                window.plot_manager.update_overall_fit_and_residuals(window)

            # Store the calculated limits for the current sheet
            sheet_name = window.sheet_combobox.GetValue()
            if sheet_name not in self.plot_limits:
                self.plot_limits[sheet_name] = {}

            self.plot_limits[sheet_name].update({
                'Xmin': float(f"{x_min - x_padding:.2f}"),
                'Xmax': float(f"{x_max + x_padding:.2f}"),
                'Ymin': float(f"{y_min - y_padding:.2f}"),
                'Ymax': float(f"{y_max + y_padding:.2f}")
            })
        else:
            # Fallback to standard zoom out
            sheet_name = window.sheet_combobox.GetValue()
            self.reset_plot_limits(window, sheet_name)
            self.resize_plot(window)

    def on_drag_tool(self, window):
        # Toggle drag mode
        window.drag_mode = not window.drag_mode
        if window.drag_mode:
            window.enable_drag()
        else:
            window.disable_drag()

    def enable_drag(self, window):
        # Activate pan mode in navigation toolbar
        window.navigation_toolbar.pan()

        # Set cursor to hand
        window.canvas.SetCursor(wx.Cursor(wx.CURSOR_HAND))

        # Connect drag release event
        window.drag_release_cid = window.canvas.mpl_connect('button_release_event',
                                                            lambda event: self.on_drag_release(window, event))

    def disable_drag(self, window):
        # Deactivate pan mode in navigation toolbar
        window.navigation_toolbar.pan()

        # Set cursor back to arrow
        window.canvas.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

        # Disconnect drag release event if connected
        if hasattr(window, 'drag_release_cid'):
            window.canvas.mpl_disconnect(window.drag_release_cid)

        # Ensure drag mode is set to False
        window.drag_mode = False

    def on_drag_release(self, window, event):
        if window.drag_mode:
            # Get current sheet name
            sheet_name = window.sheet_combobox.GetValue()

            # Update plot limits after drag
            self.update_after_drag(window, sheet_name)

            # Disable drag mode
            self.disable_drag(window)

            # Redraw the canvas
            window.canvas.draw_idle()

    def update_plot_limits(self, window, sheet_name, x_min=None, x_max=None, y_min=None, y_max=None):
        # Check if sheet exists in Core levels data
        if sheet_name not in window.Data.get('Core levels', {}):
            print(f"Warning: Sheet '{sheet_name}' not found in Core levels data")
            return

        if sheet_name not in self.plot_limits:
            self.plot_limits[sheet_name] = {}
            self.original_limits[sheet_name] = {}

        limits = self.plot_limits[sheet_name]
        original = self.original_limits[sheet_name]

        # Store original limits if not already stored
        if not original:
            # Check if B.E. and Raw Data exist
            core_level_data = window.Data['Core levels'][sheet_name]
            if 'B.E.' not in core_level_data or 'Raw Data' not in core_level_data:
                print(f"Warning: Sheet '{sheet_name}' missing B.E. or Raw Data")
                return

            x_values = core_level_data['B.E.']
            y_values = core_level_data['Raw Data']

            # Ensure we have valid data
            if not x_values or not y_values:
                print(f"Warning: Sheet '{sheet_name}' has empty B.E. or Raw Data")
                return

            original['Xmin'] = min(x_values)
            original['Xmax'] = max(x_values)
            original['Ymin'] = min(y_values) - 0.015 * max(y_values)
            original['Ymax'] = max(y_values) * 1.2  # Add 20% padding to the top

        # Update current limits
        if x_min is not None: limits['Xmin'] = x_min
        if x_max is not None: limits['Xmax'] = x_max
        if y_min is not None: limits['Ymin'] = y_min
        if y_max is not None: limits['Ymax'] = y_max

        # If any limit is not set, use the original values
        if window.energy_scale == 'KE':
            limits['Xmin'] = limits.get('Xmin', original['Xmin'])
            limits['Xmax'] = limits.get('Xmax', original['Xmax'])
        else:
            limits['Xmin'] = limits.get('Xmin', original['Xmin'])
            limits['Xmax'] = limits.get('Xmax', original['Xmax'])
        limits['Ymin'] = limits.get('Ymin', original['Ymin'])
        limits['Ymax'] = limits.get('Ymax', original['Ymax'])

    def reset_plot_limits(self, window, sheet_name):
        if sheet_name in self.original_limits:
            self.plot_limits[sheet_name] = self.original_limits[sheet_name].copy()
        else:
            self.update_plot_limits(window, sheet_name)

    def adjust_plot_limits(self, window, axis, direction):
        # Check if we're in heatmap mode
        if (hasattr(window, 'heatmap_data') and window.heatmap_data is not None and
                hasattr(window, 'file_manager') and window.file_manager is not None):
            # In heatmap mode - adjust heatmap intensity instead of Y-axis
            if axis in ['high_int', 'low_int']:
                # Both high_int and low_int will adjust heatmap vmax
                if direction == 'increase':
                    window.heatmap_vmax = min(2.0, window.heatmap_vmax + 0.05)
                elif direction == 'decrease':
                    window.heatmap_vmax = max(0.1, window.heatmap_vmax - 0.05)

                # Refresh the heatmap with new intensity
                window.file_manager.refresh_heatmap()
                return
            # For BE adjustments in heatmap mode, do nothing (or handle if needed)
            elif axis in ['high_be', 'low_be']:
                return

        # Normal plot mode - original behavior

        sheet_name = window.sheet_combobox.GetValue()

        # Handle EDX~Plot specially - uses different data structure
        if sheet_name == 'EDX~Plot':
            self._adjust_edx_plot_limits(window, axis, direction)
            return

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        limits = self.plot_limits[sheet_name]

        if axis in ['high_be', 'low_be']:
            # Check if we're on an EDX~Plot sheet
            is_edx_plot = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')

            if is_edx_plot and axis == 'high_be':
                # For EDX plots, high_be controls X-max (energy range)
                if direction == 'increase':
                    new_xmax = min(50, limits['Xmax'] + 1)
                else:  # decrease
                    new_xmax = max(1, limits['Xmax'] - 1)

                # Store for all EDX~Plot sheets
                if 'Core levels' in window.Data:
                    for sname in window.Data['Core levels']:
                        if sname == 'EDX~Plot' or sname.startswith('EDX~Plot'):
                            window.Data['Core levels'][sname]['_EDX_display_max'] = new_xmax

                limits['Xmax'] = new_xmax
                limits['Xmin'] = 0
                print(f"EDX X-max set to {new_xmax:.2f} keV for all EDX~Plot sheets")

            elif not is_edx_plot:
                # Normal XPS behavior for non-EDX plots
                increment = 0.2  # Fixed BE increment of 0.2 eV
                if axis == 'high_be':
                    if direction == 'increase':
                        limits['Xmax'] -= increment
                    elif direction == 'decrease':
                        limits['Xmax'] += increment
                elif axis == 'low_be':
                    if direction == 'increase':
                        limits['Xmin'] += increment
                    elif direction == 'decrease':
                        limits['Xmin'] -= increment
        elif axis in ['high_int', 'low_int']:
            max_intensity = max(window.y_values)
            if axis == 'high_int':
                increment = 0.05 * max_intensity
                if direction == 'increase':
                    limits['Ymax'] += increment
                elif direction == 'decrease':
                    limits['Ymax'] = max(limits['Ymax'] - increment, limits['Ymin'])
            elif axis == 'low_int':
                increment = 0.02 * max_intensity
                if direction == 'increase':
                    limits['Ymin'] = min(limits['Ymin'] + increment, limits['Ymax'])
                elif direction == 'decrease':
                    limits['Ymin'] = max(limits['Ymin'] - increment, 0)

        # Check if EDX plot to use normal axis
        is_edx_plot = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')
        if is_edx_plot:
            window.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal X-axis for EDX
        else:
            window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
            if window.energy_scale == 'KE':
                window.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])
        window.ax.set_ylim(limits['Ymin'], limits['Ymax'])
        window.canvas.draw_idle()

    def _adjust_edx_plot_limits(self, window, axis, direction):
        """Handle plot limit adjustments for EDX~Plot sheet"""
        import numpy as np

        xlim = window.ax.get_xlim()
        ylim = window.ax.get_ylim()

        sheet_data = window.Data.get('Core levels', {}).get('EDX~Plot', {})
        if 'Intensity' in sheet_data:
            max_intensity = np.max(np.array(sheet_data['Intensity']))
        else:
            max_intensity = ylim[1]

        if axis in ['high_be', 'low_be']:
            increment = 0.5  # 0.5 keV for EDX
            if axis == 'high_be':
                if direction == 'increase':
                    xlim = (xlim[0], xlim[1] + increment)
                elif direction == 'decrease':
                    xlim = (xlim[0], max(xlim[1] - increment, xlim[0] + 1))
            elif axis == 'low_be':
                if direction == 'increase':
                    xlim = (max(xlim[0] - increment, 0), xlim[1])
                elif direction == 'decrease':
                    xlim = (min(xlim[0] + increment, xlim[1] - 1), xlim[1])
        elif axis in ['high_int', 'low_int']:
            if axis == 'high_int':
                increment = 0.05 * max_intensity
                if direction == 'increase':
                    ylim = (ylim[0], ylim[1] + increment)
                elif direction == 'decrease':
                    ylim = (ylim[0], max(ylim[1] - increment, ylim[0] + 100))
            elif axis == 'low_int':
                increment = 0.02 * max_intensity
                if direction == 'increase':
                    ylim = (min(ylim[0] + increment, ylim[1] - 100), ylim[1])
                elif direction == 'decrease':
                    ylim = (max(ylim[0] - increment, 0), ylim[1])

        window.ax.set_xlim(xlim[0], xlim[1])
        window.ax.set_ylim(ylim[0], ylim[1])
        window.canvas.draw_idle()


    # This def is also in PLotManager
    def resize_plot(self, window):
        sheet_name = window.sheet_combobox.GetValue()
        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper() or "Ra_" in sheet_name
        is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')
        is_edx = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        limits = self.plot_limits[sheet_name]
        if window.energy_scale == 'KE':
            window.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])
        else:
            if is_raman or is_xas or is_edx or sheet_name.startswith('zzProfile'):
                window.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal direction for Raman
            else:
                window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis for XPS

        window.ax.set_ylim(limits['Ymin'], limits['Ymax'])
        window.canvas.draw_idle()

    def get_plot_limits_OLD(self, window, sheet_name=None):
        if sheet_name is None:
            sheet_name = window.sheet_combobox.GetValue()

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        return self.plot_limits[sheet_name]

    def get_plot_limits(self, window, sheet_name):
        # Check if sheet exists in Core levels data
        if sheet_name not in window.Data.get('Core levels', {}):
            print(f"Warning: Sheet '{sheet_name}' not found in Core levels data")
            return None

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        # Return None if update_plot_limits failed
        if sheet_name not in self.plot_limits:
            return None

        return self.plot_limits[sheet_name]


    def adjust_limit(self, window, limit_type, increment):
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        limits = self.plot_limits[sheet_name]

        if limit_type in ['Xmin', 'Xmax']:
            limits[limit_type] += increment
        elif limit_type == 'Ymax':
            limits[limit_type] = max(limits[limit_type] + increment, limits['Ymin'])
        elif limit_type == 'Ymin':
            limits[limit_type] = max(limits[limit_type] + increment, 0)

        window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
        window.ax.set_ylim(limits['Ymin'], limits['Ymax'])
        window.canvas.draw_idle()

    def get_increment(self, window, limit_type):
        sheet_name = window.sheet_combobox.GetValue()
        limits = self.plot_limits[sheet_name]
        if limit_type in ['Xmin', 'Xmax']:
            return 0.2  # Fixed BE increment of 0.2 eV
        else:  # Ymin or Ymax
            return 0.05 * (limits['Ymax'] - limits['Ymin'])

    def update_after_drag(self, window, sheet_name):
        ax = window.ax
        x_min, x_max = ax.get_xlim()
        y_min, y_max = ax.get_ylim()

        # Ensure x_min is always less than x_max (because of reversed x-axis)
        x_min, x_max = min(x_min, x_max), max(x_min, x_max)


        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper() or "Ra_" in sheet_name
        is_xas = sheet_name.startswith('XAS') or sheet_name.startswith('XAS_')
        is_edx = sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot')
        is_eels = sheet_name.startswith('EELS~')

        if sheet_name not in self.plot_limits:
            self.plot_limits[sheet_name] = {}

        limits = self.plot_limits[sheet_name]
        limits['Xmin'] = x_min
        limits['Xmax'] = x_max
        limits['Ymin'] = y_min
        limits['Ymax'] = y_max




        if window.energy_scale == 'KE':
            window.ax.set_xlim(min(x_max, x_min), max(x_max, x_min))
        else:
            if is_raman or is_eels or sheet_name.startswith('zzProfile') or is_edx:
                window.ax.set_xlim(x_min, x_max)  # Normal direction for Raman, zzProfile, and EDX
            else:
                window.ax.set_xlim(max(x_max, x_min), min(x_max, x_min))  # Reverse X-axis for XPS
        window.ax.set_ylim(y_min, y_max)

        print(f"Updated limits after drag: Xmin={x_min:.2f}, Xmax={x_max:.2f}, Ymin={y_min:.2f}, Ymax={y_max:.2f}")



class CustomRectangleSelector(widgets.RectangleSelector):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.background = None

    def release(self, event):
        super().release(event)
        self.remove_rect()

    def remove_rect(self):
        if self.visible:
            self.visible = False
            self.ax.figure.canvas.draw_idle()


class PlotLimitsWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Plot Limits", size=(280, 200),
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX))
        if 'wxGTK' in wx.PlatformInfo:  # Linux
            self.SetSize((400, 200))
            self.SetMinSize((400, 200))
            self.SetMaxSize((400, 200))
        else:
            self.SetSize((280, 200))
            self.SetMinSize((280, 200))
            self.SetMaxSize((280, 200))
        self.parent = parent
        panel = wx.Panel(self)

        # Center on parent
        main_pos = parent.GetPosition()
        main_size = parent.GetSize()
        win_size = self.GetSize()
        x = main_pos.x + (main_size.width - win_size.width) // 2
        y = main_pos.y + (main_size.height - win_size.height) // 2
        self.SetPosition((x, y))

        # Main sizer
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # BE limits
        be_box = wx.StaticBox(panel, label="Binding Energy Limits")
        be_sizer = wx.StaticBoxSizer(be_box, wx.HORIZONTAL)

        be_sizer.Add(wx.StaticText(panel, label="Max:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        if 'wxGTK' in wx.PlatformInfo:  # Linux
            self.be_max_ctrl = wx.SpinCtrlDouble(panel, min=-1000, max=2000, inc=0.1, size=(120, 30))
        else:
            self.be_max_ctrl = wx.SpinCtrlDouble(panel, min=-1000, max=2000, inc=0.1, size=(80, -1))

        be_sizer.Add(self.be_max_ctrl, 1, wx.RIGHT, 10)

        be_sizer.Add(wx.StaticText(panel, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        if 'wxGTK' in wx.PlatformInfo:  # Linux
            self.be_min_ctrl = wx.SpinCtrlDouble(panel, min=-1000, max=2000, inc=0.1, size=(120, 30))
        else:
            self.be_min_ctrl = wx.SpinCtrlDouble(panel, min=-1000, max=2000, inc=0.1, size=(80, -1))
        be_sizer.Add(self.be_min_ctrl, 1)

        # Intensity limits
        int_box = wx.StaticBox(panel, label="Intensity Limits")
        int_sizer = wx.StaticBoxSizer(int_box, wx.HORIZONTAL)

        int_sizer.Add(wx.StaticText(panel, label="Max:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        if 'wxGTK' in wx.PlatformInfo:  # Linux
            self.int_max_ctrl = wx.SpinCtrlDouble(panel, min=0, max=1e10, inc=100, size=(120, 30))
        else:
            self.int_max_ctrl = wx.SpinCtrlDouble(panel, min=0, max=1e10, inc=100, size=(80, -1))
        int_sizer.Add(self.int_max_ctrl, 1, wx.RIGHT, 10)

        int_sizer.Add(wx.StaticText(panel, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        if 'wxGTK' in wx.PlatformInfo:  # Linux
            self.int_min_ctrl = wx.SpinCtrlDouble(panel, min=-1e10, max=1e10, inc=100, size=(120, 30))
        else:
            self.int_min_ctrl = wx.SpinCtrlDouble(panel, min=-1e10, max=1e10, inc=100, size=(80, -1))

        int_sizer.Add(self.int_min_ctrl, 1)

        # Buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        apply_btn = wx.Button(panel, label="Apply")
        reset_btn = wx.Button(panel, label="Reset")

        btn_sizer.Add(apply_btn, 0, wx.RIGHT, 5)
        btn_sizer.Add(reset_btn, 0)

        # Layout
        main_sizer.Add(be_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(int_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(main_sizer)

        # Bind events
        apply_btn.Bind(wx.EVT_BUTTON, self.on_apply)
        reset_btn.Bind(wx.EVT_BUTTON, self.on_reset)

        # Initialize values
        self.update_values()

    def update_values(self):
        sheet_name = self.parent.sheet_combobox.GetValue()
        limits = self.parent.plot_config.get_plot_limits(self.parent, sheet_name)

        self.be_max_ctrl.SetValue(limits.get('Xmax', 0))
        self.be_min_ctrl.SetValue(limits.get('Xmin', 0))
        self.int_max_ctrl.SetValue(limits.get('Ymax', 0))
        self.int_min_ctrl.SetValue(limits.get('Ymin', 0))

    def on_apply(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()

        be_max = self.be_max_ctrl.GetValue()
        be_min = self.be_min_ctrl.GetValue()
        int_max = self.int_max_ctrl.GetValue()
        int_min = self.int_min_ctrl.GetValue()

        self.parent.plot_config.update_plot_limits(
            self.parent, sheet_name, be_min, be_max, int_min, int_max
        )
        self.parent.plot_config.resize_plot(self.parent)

    def on_reset(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        self.parent.plot_config.reset_plot_limits(self.parent, sheet_name)
        self.parent.plot_config.resize_plot(self.parent)
        self.update_values()