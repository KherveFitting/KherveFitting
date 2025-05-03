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

            # Update plot limits based on the selected range
            self.update_plot_limits(window, sheet_name, x_min, x_max, y_min, y_max)

            # Set x-axis limits based on energy scale and data type
            if window.energy_scale == 'KE':
                window.ax.set_xlim(min(x_max, x_min), max(x_max, x_min))
            else:
                if is_raman:
                    window.ax.set_xlim(x_min, x_max)  # Normal direction for Raman
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

    def on_zoom_out(self, window):
        # Get current sheet name
        sheet_name = window.sheet_combobox.GetValue()

        # Reset plot limits to original values
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
        if sheet_name not in self.plot_limits:
            self.plot_limits[sheet_name] = {}
            self.original_limits[sheet_name] = {}

        limits = self.plot_limits[sheet_name]
        original = self.original_limits[sheet_name]

        # Store original limits if not already stored
        if not original:
            x_values = window.Data['Core levels'][sheet_name]['B.E.']
            y_values = window.Data['Core levels'][sheet_name]['Raw Data']
            original['Xmin'] = min(x_values)
            original['Xmax'] = max(x_values)
            original['Ymin'] = min(y_values) - 0.015* max(y_values)
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
        sheet_name = window.sheet_combobox.GetValue()
        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        limits = self.plot_limits[sheet_name]

        if axis in ['high_be', 'low_be']:
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

        window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
        if window.energy_scale == 'KE':
            window.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])
        window.ax.set_ylim(limits['Ymin'], limits['Ymax'])
        window.canvas.draw_idle()

    # This def is also in PLotManager
    def resize_plot(self, window):
        sheet_name = window.sheet_combobox.GetValue()
        is_raman = sheet_name.startswith('RA') or 'RAMAN' in sheet_name.upper() or "Ra_" in sheet_name

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

        limits = self.plot_limits[sheet_name]
        if window.energy_scale == 'KE':
            window.ax.set_xlim(window.photons - limits['Xmax'], window.photons - limits['Xmin'])
        else:
            if is_raman:
                window.ax.set_xlim(limits['Xmin'], limits['Xmax'])  # Normal direction for Raman
            else:
                window.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis for XPS

        window.ax.set_ylim(limits['Ymin'], limits['Ymax'])
        window.canvas.draw_idle()

    def get_plot_limits(self, window, sheet_name=None):
        if sheet_name is None:
            sheet_name = window.sheet_combobox.GetValue()

        if sheet_name not in self.plot_limits:
            self.update_plot_limits(window, sheet_name)

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

        if sheet_name not in self.plot_limits:
            self.plot_limits[sheet_name] = {}

        limits = self.plot_limits[sheet_name]
        limits['Xmin'] = x_min
        limits['Xmax'] = x_max
        limits['Ymin'] = y_min
        limits['Ymax'] = y_max

        # Set x-axis limits based on energy scale and data type
        if window.energy_scale == 'KE':
            window.ax.set_xlim(min(x_max, x_min), max(x_max, x_min))
        else:
            if is_raman:
                window.ax.set_xlim(x_min, x_max)  # Normal direction for Raman
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

