"""
EDX_SEM_Analysis.py
Module for Energy Dispersive X-ray Spectroscopy (EDX) and Scanning Electron Microscopy (SEM) analysis
Uses HyperSpy/ExSpy library for data import and analysis
"""

import wx
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wx import NavigationToolbar2Wx as NavigationToolbar
from matplotlib.widgets import RectangleSelector, SpanSelector
import matplotlib.patches as patches
import hyperspy.api as hs


class EDXSEMWindow(wx.Frame):
    """Main window for EDX/SEM data analysis"""

    def __init__(self, parent, title="EDX HeatMap"):
        super().__init__(parent, title=title, size=(590, 700), style = wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)



        self.parent = parent
        self.edx_data = None
        self.sem_data = None
        self.current_data = None

        # Selection modes
        self.selection_mode = None  # 'point', 'area', 'line'
        self.rect_selector = None
        self.line_selector = None
        self.selected_points = []
        self.selected_areas = []
        self.selected_lines = []
        self.selected_elements = []

        self.point_size = 1  # Size in pixels (1 = single pixel)

        self.loaded_signals = []
        self.data_browser_window = None
        self.info_window = None
        self.current_colorbar = None

        self.zoom_selector = None

        self.init_ui()
        self.Centre()

        # Bind close event to clear parent reference
        self.Bind(wx.EVT_CLOSE, self.on_close)



    def init_ui(self):
        """Initialize user interface"""
        panel = wx.Panel(self, style=wx.BORDER_RAISED)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create menubar
        self.create_menubar()

        # Toolbar
        toolbar_panel = self.create_toolbar(panel)
        main_sizer.Add(toolbar_panel, 0, wx.EXPAND | wx.ALL, 2)

        # Main content - just the map
        map_panel = wx.Panel(panel)
        map_sizer = wx.BoxSizer(wx.VERTICAL)

        # Map display only
        self.map_figure = plt.Figure(figsize=(4, 4))
        self.map_figure.set_tight_layout(True)
        self.map_figure.subplots_adjust(left=0.1, right=0.95, top=0.95, bottom=0.1)

        self.map_canvas = FigureCanvas(map_panel, -1, self.map_figure)
        self.map_ax = self.map_figure.add_subplot(111)

        map_sizer.Add(self.map_canvas, 1, wx.EXPAND)
        map_panel.SetSizer(map_sizer)

        main_sizer.Add(map_panel, 1, wx.EXPAND)

        panel.SetSizer(main_sizer)
        main_sizer.Layout()
        panel.Layout()

        # Bind canvas events
        self.map_canvas.mpl_connect('button_press_event', self.on_map_click)
        self.map_canvas.mpl_connect('motion_notify_event', self.on_map_motion)
        self.map_canvas.mpl_connect('scroll_event', self.on_map_scroll)
        self._line_preview = None

        # Right-click menu
        self.map_canvas.Bind(wx.EVT_RIGHT_DOWN, self.on_right_click)

        # Store colorbar reference
        self.current_colorbar = None
        self.current_cmap = 'plasma'

        wx.CallAfter(self.Layout)

    def create_menubar(self):
        """Create menu bar"""
        menubar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()

        # Import submenu
        import_menu = wx.Menu()
        import_sem = import_menu.Append(wx.ID_ANY, "SEM Image...", "Import SEM image")
        import_edx_spectrum = import_menu.Append(wx.ID_ANY, "EDX Spectrum...", "Import single EDX spectrum")
        import_edx_map = import_menu.Append(wx.ID_ANY, "EDX Map/Spectrum Image...", "Import EDX map or spectrum image")
        import_menu.AppendSeparator()
        import_bcf = import_menu.Append(wx.ID_ANY, "Bruker BCF File...", "Import Bruker BCF file")

        file_menu.AppendSubMenu(import_menu, "Import")

        # Export submenu
        export_menu = wx.Menu()
        export_excel = export_menu.Append(wx.ID_ANY, "Export to Excel...", "Export data to Excel")
        export_csv = export_menu.Append(wx.ID_ANY, "Export to CSV...", "Export spectrum to CSV")
        export_hdf5 = export_menu.Append(wx.ID_ANY, "Export to HDF5...", "Export to HDF5 format")
        export_menu.AppendSeparator()
        export_map_image = export_menu.Append(wx.ID_ANY, "Export Map as Image...", "Save map as PNG/TIFF")

        file_menu.AppendSubMenu(export_menu, "Export")

        file_menu.AppendSeparator()
        close_item = file_menu.Append(wx.ID_CLOSE, "Close\tCtrl+W", "Close window")

        menubar.Append(file_menu, "&File")

        # Analysis menu
        analysis_menu = wx.Menu()
        quantify_item = analysis_menu.Append(wx.ID_ANY, "Quantification...", "Perform EDX quantification")
        background_item = analysis_menu.Append(wx.ID_ANY, "Background Subtraction...", "Subtract background")

        menubar.Append(analysis_menu, "&Analysis")

        self.SetMenuBar(menubar)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_import_sem, import_sem)
        self.Bind(wx.EVT_MENU, self.on_import_edx_spectrum, import_edx_spectrum)
        self.Bind(wx.EVT_MENU, self.on_import_edx_map, import_edx_map)
        self.Bind(wx.EVT_MENU, self.on_import_bcf, import_bcf)
        self.Bind(wx.EVT_MENU, self.on_export_excel, export_excel)
        self.Bind(wx.EVT_MENU, self.on_export_csv, export_csv)
        self.Bind(wx.EVT_MENU, self.on_export_hdf5, export_hdf5)
        self.Bind(wx.EVT_MENU, self.on_export_map_image, export_map_image)
        self.Bind(wx.EVT_MENU, lambda e: self.Close(), close_item)
        self.Bind(wx.EVT_MENU, self.on_quantification, quantify_item)

    def create_toolbar(self, parent):
        """Create toolbar with icon buttons only"""
        toolbar_panel = wx.Panel(parent)
        toolbar_sizer = wx.BoxSizer(wx.HORIZONTAL)

        btn_size = (25, 25)

        # Get icon path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "Icons")

        # Navigation tools
        self.zoom_in_btn = wx.BitmapToggleButton(toolbar_panel, size=btn_size)
        self.zoom_out_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        self.pan_btn = wx.BitmapToggleButton(toolbar_panel, size=btn_size)

        # Load zoom icons with fallback
        zoom_in_path = os.path.join(icon_path, "ZoomIN-3.png")
        if os.path.exists(zoom_in_path):
            self.zoom_in_btn.SetBitmap(wx.Bitmap(zoom_in_path, wx.BITMAP_TYPE_PNG))
        else:
            self.zoom_in_btn.SetBitmap(self.create_icon_bitmap('zoom_in'))

        zoom_out_path = os.path.join(icon_path, "ZoomOUT-3.png")
        if os.path.exists(zoom_out_path):
            self.zoom_out_btn.SetBitmap(wx.Bitmap(zoom_out_path, wx.BITMAP_TYPE_PNG))
        else:
            self.zoom_out_btn.SetBitmap(self.create_icon_bitmap('zoom_out'))

        pan_path = os.path.join(icon_path, "Drag-25.png")
        if os.path.exists(pan_path):
            self.pan_btn.SetBitmap(wx.Bitmap(pan_path, wx.BITMAP_TYPE_PNG))
        else:
            self.pan_btn.SetBitmap(self.create_icon_bitmap('pan'))

        self.zoom_in_btn.SetToolTip("Zoom In")
        self.zoom_out_btn.SetToolTip("Zoom Out / Home")
        self.pan_btn.SetToolTip("Pan")

        toolbar_sizer.Add(self.zoom_in_btn, 0, wx.ALL, 2)
        toolbar_sizer.Add(self.zoom_out_btn, 0, wx.ALL, 2)
        toolbar_sizer.Add(self.pan_btn, 0, wx.ALL, 2)

        toolbar_sizer.Add(wx.StaticLine(toolbar_panel, style=wx.LI_VERTICAL), 0, wx.EXPAND | wx.ALL, 3)

        # Selection tools - use created icons only
        self.point_btn = wx.BitmapToggleButton(toolbar_panel, size=btn_size)
        self.area_btn = wx.BitmapToggleButton(toolbar_panel, size=btn_size)
        self.line_btn = wx.BitmapToggleButton(toolbar_panel, size=btn_size)
        self.clear_btn = wx.BitmapButton(toolbar_panel, size=btn_size)

        self.point_btn.SetBitmap(self.create_icon_bitmap('point'))
        self.area_btn.SetBitmap(self.create_icon_bitmap('area'))
        self.line_btn.SetBitmap(self.create_icon_bitmap('line'))
        self.clear_btn.SetBitmap(self.create_icon_bitmap('clear'))

        self.point_btn.SetToolTip("Point Selection")
        self.area_btn.SetToolTip("Area Selection")
        self.line_btn.SetToolTip("Line Selection")
        self.clear_btn.SetToolTip("Clear Selections")

        toolbar_sizer.Add(self.point_btn, 0, wx.ALL, 2)
        toolbar_sizer.Add(self.area_btn, 0, wx.ALL, 2)
        toolbar_sizer.Add(self.line_btn, 0, wx.ALL, 2)
        toolbar_sizer.Add(self.clear_btn, 0, wx.ALL, 2)

        toolbar_sizer.Add(wx.StaticLine(toolbar_panel, style=wx.LI_VERTICAL), 0, wx.EXPAND | wx.ALL, 3)

        toolbar_panel.SetSizer(toolbar_sizer)

        # Plot Maps button
        self.plot_maps_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        self.plot_maps_btn.SetBitmap(self.create_icon_bitmap('intensity'))
        self.plot_maps_btn.SetToolTip("Plot Element Maps")
        toolbar_sizer.Add(self.plot_maps_btn, 0, wx.ALL, 2)

        # Intensity Map button
        self.intensity_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        heatmap_path = os.path.join(icon_path, "heatmap-3.png")
        self.intensity_btn.SetBitmap(wx.Bitmap(heatmap_path, wx.BITMAP_TYPE_PNG))
        self.intensity_btn.SetToolTip("Show Intensity Map")
        toolbar_sizer.Add(self.intensity_btn, 0, wx.ALL, 2)

        # Add spacer to push remaining buttons to the right
        toolbar_sizer.AddStretchSpacer()

        # Set Elements button
        self.set_elements_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        id_path = os.path.join(icon_path, "ID-3.png")
        if os.path.exists(id_path):
            self.set_elements_btn.SetBitmap(wx.Bitmap(id_path, wx.BITMAP_TYPE_PNG))
        else:
            self.set_elements_btn.SetBitmap(self.create_icon_bitmap('point'))
        self.set_elements_btn.SetToolTip("Set Elements")
        toolbar_sizer.Add(self.set_elements_btn, 0, wx.ALL, 2)

        # Sensitivity/Display controls button
        self.sensitivity_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        self.sensitivity_btn.SetBitmap(self.create_icon_bitmap('sensitivity'))
        self.sensitivity_btn.SetToolTip("Display & Sensitivity Controls")
        self.sensitivity_btn.Bind(wx.EVT_BUTTON, self.on_sensitivity_controls)
        toolbar_sizer.Add(self.sensitivity_btn, 0, wx.ALL, 2)



        toolbar_panel.SetSizer(toolbar_sizer)

        # Bind events
        self.zoom_in_btn.Bind(wx.EVT_TOGGLEBUTTON, self.on_zoom_in)
        self.zoom_out_btn.Bind(wx.EVT_BUTTON, self.on_zoom_out)
        self.pan_btn.Bind(wx.EVT_TOGGLEBUTTON, self.on_pan)
        self.point_btn.Bind(wx.EVT_TOGGLEBUTTON, self.on_point_mode)
        self.area_btn.Bind(wx.EVT_TOGGLEBUTTON, self.on_area_mode)
        self.line_btn.Bind(wx.EVT_TOGGLEBUTTON, self.on_line_mode)
        self.clear_btn.Bind(wx.EVT_BUTTON, self.on_clear_selections)
        self.set_elements_btn.Bind(wx.EVT_BUTTON, self.on_set_elements)
        self.plot_maps_btn.Bind(wx.EVT_BUTTON, self.on_plot_maps)
        self.intensity_btn.Bind(wx.EVT_BUTTON, self.on_intensity_map)

        return toolbar_panel

    def create_icon_bitmap(self, icon_type):
        """Create simple icon bitmaps"""
        size = 20
        bmp = wx.Bitmap(size, size)
        dc = wx.MemoryDC(bmp)
        dc.SetBackground(wx.Brush(wx.Colour(240, 240, 240)))
        dc.Clear()

        dc.SetPen(wx.Pen(wx.Colour(50, 50, 50), 2))
        dc.SetBrush(wx.Brush(wx.Colour(100, 100, 100)))

        if icon_type == 'zoom_in':
            dc.DrawCircle(10, 10, 6)
            dc.DrawLine(8, 10, 12, 10)
            dc.DrawLine(10, 8, 10, 12)
            dc.DrawLine(14, 14, 18, 18)
        elif icon_type == 'zoom_out':
            dc.DrawCircle(10, 10, 6)
            dc.DrawLine(8, 10, 12, 10)
            dc.DrawLine(14, 14, 18, 18)
        elif icon_type == 'pan':
            dc.DrawLine(10, 3, 10, 17)
            dc.DrawLine(3, 10, 17, 10)
            dc.DrawLine(10, 3, 7, 6)
            dc.DrawLine(10, 3, 13, 6)
            dc.DrawLine(10, 17, 7, 14)
            dc.DrawLine(10, 17, 13, 14)
        elif icon_type == 'point':
            dc.SetBrush(wx.Brush(wx.Colour(79, 190, 159)))
            dc.DrawCircle(10, 10, 4)
        elif icon_type == 'area':
            dc.SetBrush(wx.TRANSPARENT_BRUSH)
            dc.SetPen(wx.Pen(wx.Colour(79, 190, 159), 2))
            dc.DrawRectangle(3, 3, 14, 14)
        elif icon_type == 'line':
            dc.SetPen(wx.Pen(wx.Colour(79, 190, 159), 2))
            dc.DrawLine(3, 17, 17, 3)
        elif icon_type == 'clear':
            dc.SetPen(wx.Pen(wx.Colour(200, 50, 50), 2))
            dc.DrawLine(5, 5, 15, 15)
            dc.DrawLine(15, 5, 5, 15)
        elif icon_type == 'intensity' or icon_type == 'heatmap':
            # Draw gradient-like squares for heatmap
            dc.SetPen(wx.Pen(wx.Colour(50, 50, 50), 1))
            dc.SetBrush(wx.Brush(wx.Colour(0, 0, 255)))
            dc.DrawRectangle(2, 2, 7, 7)
            dc.SetBrush(wx.Brush(wx.Colour(0, 255, 0)))
            dc.DrawRectangle(11, 2, 7, 7)
            dc.SetBrush(wx.Brush(wx.Colour(255, 255, 0)))
            dc.DrawRectangle(2, 11, 7, 7)
            dc.SetBrush(wx.Brush(wx.Colour(255, 0, 0)))
            dc.DrawRectangle(11, 11, 7, 7)
        elif icon_type == 'sensitivity':
            # Slider icon
            dc.SetPen(wx.Pen(wx.Colour(79, 190, 159), 2))
            # Horizontal line (slider track)
            dc.DrawLine(5, 10, 15, 10)
            # Slider knob
            dc.SetBrush(wx.Brush(wx.Colour(79, 190, 159)))
            dc.DrawCircle(10, 10, 3)
            # Small lines above and below for adjustment
            dc.DrawLine(10, 4, 10, 7)
            dc.DrawLine(10, 13, 10, 16)

        dc.SelectObject(wx.NullBitmap)
        return bmp



    # ==================== TOOLBAR HANDLERS ====================

    def on_sensitivity_controls(self, event):
        """Open sensitivity and display controls window"""
        if hasattr(self, 'sensitivity_window') and self.sensitivity_window and not self.sensitivity_window.IsBeingDeleted():
            # Window already exists, bring to front
            self.sensitivity_window.Raise()
        else:
            # Create new window
            self.sensitivity_window = EDXSensitivityWindow(self)
            self.sensitivity_window.Show()

    def on_zoom_in(self, event):
        """Toggle zoom in mode using rectangle selector"""
        if self.zoom_in_btn.GetValue():
            self.pan_btn.SetValue(False)
            self.deactivate_selection_modes()

            # Create zoom rectangle selector
            if not hasattr(self, 'zoom_selector') or self.zoom_selector is None:
                self.zoom_selector = RectangleSelector(
                    self.map_ax,
                    self.on_zoom_select,
                    useblit=True,
                    props=dict(facecolor='blue', edgecolor='blue', alpha=0.2, fill=True),
                    button=[1],
                    minspanx=5,
                    minspany=5,
                    spancoords='pixels',
                    interactive=False
                )
            else:
                self.zoom_selector.set_active(True)
        else:
            if hasattr(self, 'zoom_selector') and self.zoom_selector:
                self.zoom_selector.set_active(False)

    def on_zoom_select(self, eclick, erelease):
        """Handle zoom rectangle selection"""
        x1, y1 = eclick.xdata, eclick.ydata
        x2, y2 = erelease.xdata, erelease.ydata

        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        self.map_ax.set_xlim(x_min, x_max)
        self.map_ax.set_ylim(y_max, y_min)  # Inverted for image
        self.map_canvas.draw()

        # Deselect zoom button after use
        self.zoom_in_btn.SetValue(False)
        if hasattr(self, 'zoom_selector') and self.zoom_selector:
            self.zoom_selector.set_active(False)

    def on_zoom_out(self, event):
        """Zoom out / reset view"""
        self.map_ax.autoscale()
        self.map_canvas.draw()

        # Deselect buttons
        self.zoom_in_btn.SetValue(False)
        self.pan_btn.SetValue(False)
        if hasattr(self, 'zoom_selector') and self.zoom_selector:
            self.zoom_selector.set_active(False)

    def on_pan(self, event):
        """Toggle pan mode"""
        if self.pan_btn.GetValue():
            self.zoom_in_btn.SetValue(False)
            self.deactivate_selection_modes()

            if hasattr(self, 'zoom_selector') and self.zoom_selector:
                self.zoom_selector.set_active(False)

            # Enable pan via mouse drag
            self._pan_start = None
            self._pan_cid_press = self.map_canvas.mpl_connect('button_press_event', self._on_pan_press)
            self._pan_cid_release = self.map_canvas.mpl_connect('button_release_event', self._on_pan_release)
            self._pan_cid_motion = self.map_canvas.mpl_connect('motion_notify_event', self._on_pan_motion)
        else:
            # Disconnect pan events
            if hasattr(self, '_pan_cid_press'):
                self.map_canvas.mpl_disconnect(self._pan_cid_press)
            if hasattr(self, '_pan_cid_release'):
                self.map_canvas.mpl_disconnect(self._pan_cid_release)
            if hasattr(self, '_pan_cid_motion'):
                self.map_canvas.mpl_disconnect(self._pan_cid_motion)

    def _on_pan_press(self, event):
        if event.inaxes == self.map_ax and event.button == 1:
            self._pan_start = (event.xdata, event.ydata)

    def _on_pan_release(self, event):
        self._pan_start = None
        self._add_scale_bar()
        self.map_canvas.draw_idle()

    def _on_pan_motion(self, event):
        if self._pan_start is None or event.inaxes != self.map_ax:
            return

        dx = self._pan_start[0] - event.xdata
        dy = self._pan_start[1] - event.ydata

        xlim = self.map_ax.get_xlim()
        ylim = self.map_ax.get_ylim()

        self.map_ax.set_xlim(xlim[0] + dx, xlim[1] + dx)
        self.map_ax.set_ylim(ylim[0] + dy, ylim[1] + dy)

        self.map_canvas.draw_idle()

    def on_home(self, event):
        """Reset view to home"""
        self.map_toolbar.home()
        self.map_canvas.draw()

    def deactivate_selection_modes(self):
        """Deactivate all selection modes"""
        self.point_btn.SetValue(False)
        self.area_btn.SetValue(False)
        self.line_btn.SetValue(False)
        self.selection_mode = None

        if self.rect_selector:
            self.rect_selector.set_active(False)
        if self.line_selector:
            self.line_selector.set_active(False)

    def on_point_mode(self, event):
        """Toggle point selection mode"""
        if self.point_btn.GetValue():
            self.zoom_in_btn.SetValue(False)
            self.pan_btn.SetValue(False)
            self.area_btn.SetValue(False)
            self.line_btn.SetValue(False)
            self.selection_mode = 'point'

            if hasattr(self, 'zoom_selector') and self.zoom_selector:
                self.zoom_selector.set_active(False)
            if self.rect_selector:
                self.rect_selector.set_active(False)

            # Clear previous selections
            self.clear_selection_markers()
            self.map_canvas.draw()
        else:
            self.selection_mode = None

    def on_area_mode(self, event):
        """Toggle area selection mode with rotatable rectangle"""
        if self.area_btn.GetValue():
            self.zoom_in_btn.SetValue(False)
            self.pan_btn.SetValue(False)
            self.point_btn.SetValue(False)
            self.line_btn.SetValue(False)
            self.selection_mode = 'area'

            if hasattr(self, 'zoom_selector') and self.zoom_selector:
                self.zoom_selector.set_active(False)

            # Clear previous selections
            self.clear_selection_markers()
            self.map_canvas.draw()

            # Deactivate old rectangle selector
            if self.rect_selector:
                self.rect_selector.set_active(False)

            # Create or activate rotatable rectangle
            if not hasattr(self, 'rotatable_rect') or self.rotatable_rect is None:
                self.rotatable_rect = RotatableRectangle(self.map_ax, self.on_rotated_area_complete)

            # Connect click event to start new selection
            if not hasattr(self, '_area_click_cid'):
                self._area_click_cid = self.map_canvas.mpl_connect('button_press_event',
                                                                   lambda e: self.rotatable_rect.start_selection(e) if self.area_btn.GetValue() and e.button == 1 else None)
        else:
            self.selection_mode = None
            if self.rect_selector:
                self.rect_selector.set_active(False)
            if hasattr(self, 'rotatable_rect') and self.rotatable_rect:
                self.rotatable_rect.clear()
            if hasattr(self, '_area_click_cid'):
                self.map_canvas.mpl_disconnect(self._area_click_cid)
                self._area_click_cid = None

    def on_rotated_area_complete(self, center, width, height, angle):
        """Handle completion of rotated area selection with proper pixel extraction"""
        if self.current_data is None:
            return

        data = self.current_data.data
        if len(data.shape) != 3:
            return

        # Get dimensions
        map_height, map_width = data.shape[0], data.shape[1]
        cx, cy = center
        hw, hh = width / 2, height / 2

        # Create rotation matrix
        angle_rad = np.radians(-angle)  # Negative for inverse transform
        cos_a = np.cos(angle_rad)
        sin_a = np.sin(angle_rad)

        # Create mask for pixels inside rotated rectangle
        y_indices, x_indices = np.meshgrid(np.arange(map_height), np.arange(map_width), indexing='ij')

        # Transform all pixels to rectangle's local coordinate system
        local_x = (x_indices - cx) * cos_a - (y_indices - cy) * sin_a
        local_y = (x_indices - cx) * sin_a + (y_indices - cy) * cos_a

        # Create mask for pixels inside rectangle
        mask = (np.abs(local_x) <= hw) & (np.abs(local_y) <= hh)

        # Extract spectra from masked pixels
        masked_spectra = data[mask]

        if len(masked_spectra) == 0:
            print("No pixels in selected area")
            return

        # Sum spectra
        summed_spectrum = np.sum(masked_spectra, axis=0)

        # Get energy axis
        energy_axis = self.current_data.axes_manager[-1]
        energy = energy_axis.axis

        # Plot to parent window
        if self.parent is not None:
            # Create EDX~Plot sheet if it doesn't exist
            if 'EDX~Plot' not in self.parent.Data['Core levels']:
                self.parent.Data['Core levels']['EDX~Plot'] = {}

            # Store data
            self.parent.Data['Core levels']['EDX~Plot']['Energy_keV'] = energy.tolist()
            self.parent.Data['Core levels']['EDX~Plot']['Intensity'] = summed_spectrum.tolist()

            # Update sheet combobox
            self.parent.sheet_combobox.SetValue('EDX~Plot')

            # Plot
            self.parent.ax.clear()
            self.parent.ax.plot(energy, summed_spectrum, 'k-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'EDX Area Spectrum (Rotated, {len(masked_spectra)} pixels)')
            self.parent.ax.grid(False)

            # Add peak labels
            self.add_peak_labels(self.parent.ax, energy, summed_spectrum)

            # Get stored X max or default to 20
            display_x_max = self.parent.Data['Core levels']['EDX~Plot'].get('_EDX_display_max', 20)
            self.parent.ax.set_xlim(0, display_x_max)
            self.parent.ax.set_ylim(0, np.max(summed_spectrum) * 1.1)

            self.parent.canvas.draw()

            # Populate peak fitting grid with quantification
            self._populate_edx_grid(energy, summed_spectrum, f"Rotated Area ({len(masked_spectra)} px)")

        print(f"Rotated area: center=({cx:.0f},{cy:.0f}), size=({width:.0f}x{height:.0f}), angle={angle:.1f}°")
        print(f"Extracted {len(masked_spectra)} pixels")

    def on_line_mode(self, event):
        """Toggle line selection mode"""
        if self.line_btn.GetValue():
            self.zoom_in_btn.SetValue(False)
            self.pan_btn.SetValue(False)
            self.point_btn.SetValue(False)
            self.area_btn.SetValue(False)
            self.selection_mode = 'line'
            self.line_start = None
            self.line_end = None

            if hasattr(self, 'zoom_selector') and self.zoom_selector:
                self.zoom_selector.set_active(False)
            if self.rect_selector:
                self.rect_selector.set_active(False)

            # Clear previous selections
            self.clear_selection_markers()
            self.map_canvas.draw()
        else:
            self.selection_mode = None

    def on_clear_selections(self, event=None):
        """Clear all selections"""
        self.selected_points = []
        self.selected_areas = []
        self.selected_lines = []
        self.line_start = None
        self.line_end = None

        # Clear rotatable rectangle if it exists
        if hasattr(self, 'rotatable_rect') and self.rotatable_rect:
            self.rotatable_rect.clear()

        # Clear selection markers
        self.clear_selection_markers()

        self.map_canvas.draw()

    def on_colormap_change(self, event):
        """Change colormap"""
        if self.current_data is not None:
            self.plot_current_map()


    # =================== RIGHT CLICK HANDLERS ===================

    def on_map_right_click(self, event):
        """Handle right-click on map for context menu"""
        if event.button != 3:  # Right click
            return
        if event.inaxes != self.map_ax:
            return

        # Create popup menu
        menu = wx.Menu()

        save_plot_item = menu.Append(wx.ID_ANY, "Save Current EDX Plot...")
        self.Bind(wx.EVT_MENU, self.on_save_edx_plot, save_plot_item)

        # Show menu at mouse position
        self.PopupMenu(menu)
        menu.Destroy()

    def on_save_edx_plot(self, event):
        """Save current EDX plot to a numbered sheet in Excel and window.Data"""
        if self.parent is None:
            return

        # Check if Excel file exists
        if 'FilePath' not in self.parent.Data or not self.parent.Data['FilePath']:
            wx.MessageBox("No Excel file found. Please ensure the EDX data was imported correctly.",
                          "No File", wx.OK | wx.ICON_WARNING)
            return

        # Find next available EDX~Plot number
        existing_sheets = list(self.parent.Data.get('Core levels', {}).keys())
        plot_num = 1
        while f'EDX~Plot{plot_num}' in existing_sheets:
            plot_num += 1

        sheet_name = f'EDX~Plot{plot_num}'

        # Gather current selection info
        selection_info = self._get_current_selection_info()

        if selection_info is None:
            wx.MessageBox("No selection to save. Please select a point, line, or area first.",
                          "No Selection", wx.OK | wx.ICON_WARNING)
            return

        # Get current spectrum data from parent
        if not hasattr(self.parent, 'ax') or len(self.parent.ax.lines) == 0:
            wx.MessageBox("No EDX plot data to save.", "No Data", wx.OK | wx.ICON_WARNING)
            return

        # Get spectrum data from plot
        line = self.parent.ax.lines[0]
        energy = line.get_xdata()
        intensity = line.get_ydata()

        # Get grid data
        grid_data = self._get_grid_data()

        # Create sheet data - USE SAME STRUCTURE AS XPS SHEETS
        import datetime
        import json
        sheet_data = {
            'Name': sheet_name,
            'B.E.': list(energy),
            'Raw Data': list(intensity),
            '_EDX_display_max': 20,
            '_EDX_type': 'plot',
            '_EDX_selection': selection_info,
            '_EDX_grid_data': grid_data,
            '_EDX_save_time': datetime.datetime.now().isoformat(),
            'Background': {}
        }

        # Save to parent Data
        if 'Core levels' not in self.parent.Data:
            self.parent.Data['Core levels'] = {}

        self.parent.Data['Core levels'][sheet_name] = sheet_data

        # Update sheet combobox
        if sheet_name not in [self.parent.sheet_combobox.GetString(i)
                              for i in range(self.parent.sheet_combobox.GetCount())]:
            self.parent.sheet_combobox.Append(sheet_name)

        # ========== Save to Excel file ==========
        try:
            import pandas as pd
            import openpyxl

            file_path = self.parent.Data['FilePath']

            # Create DataFrame for this plot
            edx_df = pd.DataFrame({
                'Energy (keV)': [f"{v:.2f}" for v in energy],
                'Intensity': [f"{v:.2f}" for v in intensity]
            })

            # Append to Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                edx_df.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"Saved {sheet_name} to Excel: {file_path}")

            # Also update JSON
            json_path = file_path.replace('.xlsx', '.json')
            if os.path.exists(json_path):
                with open(json_path, 'r') as f:
                    json_data = json.load(f)

                if 'Core levels' not in json_data:
                    json_data['Core levels'] = {}

                json_data['Core levels'][sheet_name] = {
                    'Name': sheet_name,
                    'B.E.': [float(f"{v:.2f}") for v in energy],
                    'Raw Data': [float(f"{v:.2f}") for v in intensity],
                    '_EDX_display_max': 20,
                    '_EDX_type': 'plot',
                    '_EDX_selection': selection_info,
                    '_EDX_save_time': datetime.datetime.now().isoformat()
                }

                with open(json_path, 'w') as f:
                    json.dump(json_data, f, indent=2)

                print(f"Updated JSON: {json_path}")

            wx.MessageBox(f"EDX plot saved as '{sheet_name}'\n\nSaved to Excel and JSON files.",
                          "Saved", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Saved to memory but error saving to Excel:\n{str(e)}\n\nUse 'Save' from toolbar to save all data.",
                          "Partial Save", wx.OK | wx.ICON_WARNING)
            import traceback
            traceback.print_exc()

    def _get_current_selection_info(self):
        """Get information about current selection"""
        info = {}

        # Check for point selection
        if self.selected_points:
            point = self.selected_points[-1]
            info['type'] = 'point'
            info['x'] = point[0]
            info['y'] = point[1]
            info['size'] = point[2] if len(point) > 2 else 1
            return info

        # Check for line selection
        if hasattr(self, 'line_start') and self.line_start and hasattr(self, 'line_end') and self.line_end:
            info['type'] = 'line'
            info['x1'] = self.line_start[0]
            info['y1'] = self.line_start[1]
            info['x2'] = self.line_end[0]
            info['y2'] = self.line_end[1]
            return info

        # Check for rotatable rectangle
        if hasattr(self, 'rotatable_rect') and self.rotatable_rect and self.rotatable_rect.active:
            rect = self.rotatable_rect
            info['type'] = 'rectangle'
            info['center_x'] = rect.center[0]
            info['center_y'] = rect.center[1]
            info['width'] = rect.width
            info['height'] = rect.height
            info['angle'] = rect.angle
            return info

        # Check for area selection
        if self.selected_areas:
            area = self.selected_areas[-1]
            info['type'] = 'area'
            info['x1'] = area[0]
            info['y1'] = area[1]
            info['x2'] = area[2]
            info['y2'] = area[3]
            return info

        return None

    def _get_grid_data(self):
        """Get current peak fitting grid data"""
        if self.parent is None or not hasattr(self.parent, 'peak_params_grid'):
            return []

        grid = self.parent.peak_params_grid
        data = []

        for row in range(0, grid.GetNumberRows(), 2):  # Only data rows, skip constraint rows
            row_data = {
                'id': grid.GetCellValue(row, 0),
                'label': grid.GetCellValue(row, 1),
                'position': grid.GetCellValue(row, 2),
                'height': grid.GetCellValue(row, 3),
                'fwhm': grid.GetCellValue(row, 4),
                'area': grid.GetCellValue(row, 6),
                'concentration': grid.GetCellValue(row, 10),
            }
            data.append(row_data)

        return data

    def draw_saved_selection(self, selection_info):
        """Draw a saved selection on the map"""
        # Clear previous saved selection markers
        self._clear_saved_selection()

        if selection_info is None:
            return

        sel_type = selection_info.get('type')

        if sel_type == 'point':
            x, y = selection_info['x'], selection_info['y']
            size = selection_info.get('size', 1)

            if size == 1:
                marker, = self.map_ax.plot(x, y, 'k+', markersize=15, markeredgewidth=2)
                marker._is_saved_selection = True
            else:
                from matplotlib.patches import Rectangle
                half_size = size / 2
                rect = Rectangle((x - half_size, y - half_size), size, size,
                                 linewidth=2, edgecolor='black', facecolor='none')
                rect._is_saved_selection = True
                self.map_ax.add_patch(rect)

        elif sel_type == 'line':
            line, = self.map_ax.plot([selection_info['x1'], selection_info['x2']],
                                     [selection_info['y1'], selection_info['y2']],
                                     'k-', linewidth=2)
            line._is_saved_selection = True

        elif sel_type == 'rectangle':
            from matplotlib.patches import Rectangle
            from matplotlib.transforms import Affine2D

            cx, cy = selection_info['center_x'], selection_info['center_y']
            w, h = selection_info['width'], selection_info['height']
            angle = selection_info['angle']

            rect = Rectangle((-w / 2, -h / 2), w, h,
                             linewidth=2, edgecolor='black', facecolor='none')
            t = Affine2D().rotate_deg(angle).translate(cx, cy) + self.map_ax.transData
            rect.set_transform(t)
            rect._is_saved_selection = True
            self.map_ax.add_patch(rect)

        elif sel_type == 'area':
            from matplotlib.patches import Rectangle
            x1, y1 = selection_info['x1'], selection_info['y1']
            x2, y2 = selection_info['x2'], selection_info['y2']
            rect = Rectangle((x1, y1), x2 - x1, y2 - y1,
                             linewidth=2, edgecolor='black', facecolor='none')
            rect._is_saved_selection = True
            self.map_ax.add_patch(rect)

        self.map_canvas.draw_idle()

    def _clear_saved_selection(self):
        """Clear saved selection markers"""
        for artist in self.map_ax.patches[:]:
            if hasattr(artist, '_is_saved_selection') and artist._is_saved_selection:
                try:
                    artist.remove()
                except:
                    pass
        for artist in self.map_ax.lines[:]:
            if hasattr(artist, '_is_saved_selection') and artist._is_saved_selection:
                try:
                    artist.remove()
                except:
                    pass

    # ==================== SELECTION HANDLERS ====================

    def on_map_scroll(self, event):
        """Handle mouse scroll on map"""
        if event.inaxes != self.map_ax:
            return

        if self.selection_mode == 'point':
            # Adjust point size with scroll wheel
            if event.button == 'up':
                self.point_size = min(50, self.point_size + 1)
            elif event.button == 'down':
                self.point_size = max(1, self.point_size - 1)

            # Update point preview if we have a position
            self._update_point_preview(event.xdata, event.ydata)

    def _update_point_preview(self, x, y):
        """Update point size preview rectangle"""
        # Remove existing preview
        self._clear_point_preview()

        if x is None or y is None:
            return

        if self.point_size == 1:
            # Single pixel - show as cross
            marker, = self.map_ax.plot(x, y, 'g+', markersize=15, markeredgewidth=2, alpha=0.5)
            marker._is_point_preview = True
        else:
            # Show rectangle preview
            from matplotlib.patches import Rectangle
            half_size = self.point_size / 2
            rect = Rectangle((x - half_size, y - half_size), self.point_size, self.point_size,
                             linewidth=2, edgecolor='lime', facecolor='green', alpha=0.3)
            rect._is_point_preview = True
            self.map_ax.add_patch(rect)

        self.map_canvas.draw_idle()

    def _clear_point_preview(self):
        """Clear point preview"""
        for artist in self.map_ax.patches[:]:
            if hasattr(artist, '_is_point_preview') and artist._is_point_preview:
                try:
                    artist.remove()
                except:
                    pass
        for artist in self.map_ax.lines[:]:
            if hasattr(artist, '_is_point_preview') and artist._is_point_preview:
                try:
                    artist.remove()
                except:
                    pass

    def on_map_click(self, event):
        """Handle click on map"""
        if event.inaxes != self.map_ax:
            return
        if self.current_data is None:
            return

        x, y = int(event.xdata), int(event.ydata)

        # Validate coordinates
        data_shape = self.current_data.data.shape
        if len(data_shape) < 2:
            return

        max_x = data_shape[1] if len(data_shape) >= 2 else data_shape[0]
        max_y = data_shape[0]

        if x < 0 or x >= max_x or y < 0 or y >= max_y:
            return

        if self.selection_mode == 'point':
            # Clear previous markers and preview
            self.clear_selection_markers()
            self._clear_point_preview()

            # Store point info with size
            self.selected_points = [(x, y, self.point_size)]

            if self.point_size == 1:
                # Single pixel - show as cross
                marker, = self.map_ax.plot(x, y, 'g+', markersize=15, markeredgewidth=2)
                marker._is_selection_marker = True
            else:
                # Show rectangle for multi-pixel selection
                from matplotlib.patches import Rectangle
                half_size = self.point_size / 2
                rect = Rectangle((x - half_size, y - half_size), self.point_size, self.point_size,
                                 linewidth=2, edgecolor='lime', facecolor='green', alpha=0.3)
                rect._is_selection_marker = True
                self.map_ax.add_patch(rect)

            self.map_canvas.draw()

            # Extract and plot spectrum
            self.plot_point_spectrum(x, y)

        elif self.selection_mode == 'line':
            if self.line_start is None:
                # First click - start of line
                # Clear any previous line markers first (for starting a new line)
                self.clear_selection_markers()
                self._clear_line_preview()

                self.line_start = (x, y)
                # Draw start marker
                marker, = self.map_ax.plot(x, y, 'go', markersize=8, markeredgecolor='darkgreen', markeredgewidth=2)
                marker._is_selection_marker = True
                self.map_canvas.draw()
            else:
                # Second click - end of line
                self.line_end = (x, y)

                # Clear preview line
                self._clear_line_preview()

                # Clear previous markers and redraw final line
                self.clear_selection_markers()

                # Draw final line
                line, = self.map_ax.plot([self.line_start[0], self.line_end[0]],
                                         [self.line_start[1], self.line_end[1]],
                                         'g-', linewidth=2)
                line._is_selection_marker = True

                # Draw end points
                marker1, = self.map_ax.plot(self.line_start[0], self.line_start[1], 'go',
                                            markersize=8, markeredgecolor='darkgreen', markeredgewidth=2)
                marker1._is_selection_marker = True
                marker2, = self.map_ax.plot(self.line_end[0], self.line_end[1], 'go',
                                            markersize=8, markeredgecolor='darkgreen', markeredgewidth=2)
                marker2._is_selection_marker = True

                self.map_canvas.draw()

                # Plot line spectrum
                self.plot_line_spectrum(self.line_start[0], self.line_start[1],
                                        self.line_end[0], self.line_end[1])

                # Reset for next line
                self.line_start = None
                self.line_end = None

    def on_map_motion(self, event):
        """Handle mouse motion on map for line preview"""
        if self.selection_mode != 'line' or self.line_start is None:
            return
        if event.inaxes != self.map_ax:
            return
        if event.xdata is None or event.ydata is None:
            return

        x, y = int(event.xdata), int(event.ydata)

        # Update line preview
        self._clear_line_preview()
        self._line_preview, = self.map_ax.plot([self.line_start[0], x],
                                               [self.line_start[1], y],
                                               'g--', linewidth=1.5, alpha=0.7)
        self._line_preview._is_selection_marker = True
        self.map_canvas.draw_idle()

    def _clear_line_preview(self):
        """Clear the line preview"""
        if self._line_preview is not None:
            try:
                self._line_preview.remove()
            except (ValueError, AttributeError):
                pass
            self._line_preview = None

    def clear_selection_markers(self):
        """Clear all selection markers from map"""
        # Remove marked lines and points
        for artist in self.map_ax.lines[:]:
            if hasattr(artist, '_is_selection_marker') and artist._is_selection_marker:
                artist.remove()
        for artist in self.map_ax.patches[:]:
            if hasattr(artist, '_is_selection_marker') and artist._is_selection_marker:
                artist.remove()

    def on_area_select(self, eclick, erelease):
        """Handle area selection"""
        if self.current_data is None:
            return

        x1, y1 = int(eclick.xdata), int(eclick.ydata)
        x2, y2 = int(erelease.xdata), int(erelease.ydata)

        # Ensure proper order
        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        self.selected_areas.append((x_min, y_min, x_max, y_max))

        # Plot summed spectrum from area
        self.plot_area_spectrum(x_min, y_min, x_max, y_max)

    def on_line_select(self, xmin, xmax):
        """Handle line selection (horizontal profile)"""
        if self.current_data is None:
            return

        x1, x2 = int(xmin), int(xmax)
        self.selected_lines.append((x1, x2))

        # Plot line profile spectrum
        self.plot_line_spectrum(x1, x2)

    # ==================== SPECTRUM PLOTTING ====================

    def plot_point_spectrum(self, x, y):
        """Plot spectrum from point or area around point"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D (x, y, energy) for point spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        # Get point size
        point_size = getattr(self, 'point_size', 1)

        if point_size == 1:
            # Single pixel
            spectrum = data[y, x, :]
            title = f'EDX Spectrum at Point ({x}, {y})'
        else:
            # Average over area
            half_size = point_size // 2
            y1 = max(0, y - half_size)
            y2 = min(data.shape[0], y + half_size + 1)
            x1 = max(0, x - half_size)
            x2 = min(data.shape[1], x + half_size + 1)

            spectrum = np.mean(data[y1:y2, x1:x2, :], axis=(0, 1))
            title = f'EDX Spectrum at Point ({x}, {y}) [{point_size}×{point_size} px]'

        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self.parent.ax.clear()
            self.parent.ax.plot(energy, spectrum, 'k-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(title)
            self.parent.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

            # Add element peak labels
            self.add_peak_labels(self.parent.ax, energy, spectrum)

            # Get stored X max or default to 20
            current_sheet = self.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                if current_sheet in self.parent.Data['Core levels']:
                    display_x_max = self.parent.Data['Core levels'][current_sheet].get('_EDX_display_max', 20)
                else:
                    display_x_max = 20
            else:
                display_x_max = 20

            self.parent.ax.set_xlim(0, display_x_max)
            self.parent.ax.set_ylim(np.min(spectrum) * 0.95, np.max(spectrum) * 1.1)

            self.parent.canvas.draw()

            # Populate peak fitting grid with quantification
            self._populate_edx_grid(energy, spectrum, title)

    def plot_area_spectrum(self, x1, y1, x2, y2):
        """Plot summed spectrum from area in KherveFitting main window"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D (x, y, energy) for area spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        # Sum spectrum over area
        spectrum = np.sum(data[y1:y2 + 1, x1:x2 + 1, :], axis=(0, 1))
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self.parent.ax.clear()
            self.parent.ax.plot(energy, spectrum, 'k-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'Summed EDX Spectrum - Area: {(x2 - x1 + 1) * (y2 - y1 + 1)} pixels')

            # Set Y-axis to scientific format
            self.parent.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

            # Add element peak labels
            self.add_peak_labels(self.parent.ax, energy, spectrum)

            # Get stored X max or default to 20
            current_sheet = self.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                if current_sheet in self.parent.Data['Core levels']:
                    display_x_max = self.parent.Data['Core levels'][current_sheet].get('_EDX_display_max', 20)
                else:
                    display_x_max = 20
            else:
                display_x_max = 20

            self.parent.ax.set_xlim(0, display_x_max)
            self.parent.ax.set_ylim(np.min(spectrum) * 0.95, np.max(spectrum) * 1.1)

            self.parent.canvas.draw()

            # Populate peak fitting grid with quantification
            self._populate_edx_grid(energy, spectrum, f"Area ({x2 - x1 + 1}×{y2 - y1 + 1})")

    def plot_line_spectrum(self, x1, y1, x2, y2):
        """Plot summed spectrum along line from (x1,y1) to (x2,y2) in KherveFitting main window"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D (x, y, energy) for line spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        # Get points along line using Bresenham-like algorithm
        num_points = max(abs(x2 - x1), abs(y2 - y1)) + 1
        x_coords = np.linspace(x1, x2, num_points).astype(int)
        y_coords = np.linspace(y1, y2, num_points).astype(int)

        # Sum spectrum along line
        spectrum = np.zeros(data.shape[2])
        for px, py in zip(x_coords, y_coords):
            if 0 <= px < data.shape[1] and 0 <= py < data.shape[0]:
                spectrum += data[py, px, :]

        energy = self.get_energy_axis()
        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self.parent.ax.clear()
            self.parent.ax.plot(energy, spectrum, 'k-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'EDX Line Spectrum ({x1},{y1}) to ({x2},{y2}) - {num_points} pts')

            # Set Y-axis to scientific format
            self.parent.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

            # Add element peak labels
            self.add_peak_labels(self.parent.ax, energy, spectrum)

            # Get stored X max or default to 20
            current_sheet = self.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                if current_sheet in self.parent.Data['Core levels']:
                    display_x_max = self.parent.Data['Core levels'][current_sheet].get('_EDX_display_max', 20)
                else:
                    display_x_max = 20
            else:
                display_x_max = 20

            self.parent.ax.set_xlim(0, display_x_max)
            self.parent.ax.set_ylim(np.min(spectrum) * 0.95, np.max(spectrum) * 1.1)

            self.parent.canvas.draw()

            # Populate peak fitting grid with quantification
            self._populate_edx_grid(energy, spectrum, f"Line ({num_points} pts)")

    def _add_scale_bar(self):
        """Add scale bar to the map based on current view and axes scale information"""
        if self.current_data is None:
            return

        # Remove existing scale bar elements
        self._remove_scale_bar()

        # Get scale from navigation axes
        scale_per_pixel = None
        scale_unit = 'nm'

        if hasattr(self.current_data, 'axes_manager'):
            nav_axes = self.current_data.axes_manager.navigation_axes
            if len(nav_axes) > 0:
                axis = nav_axes[0]
                scale_per_pixel = axis.scale
                scale_unit = axis.units if hasattr(axis, 'units') and axis.units else 'nm'

        if scale_per_pixel is None or scale_per_pixel <= 0:
            return

        # Get current view limits (for zoomed view)
        xlim = self.map_ax.get_xlim()
        ylim = self.map_ax.get_ylim()

        # Calculate visible width in pixels and nm
        visible_width_pixels = abs(xlim[1] - xlim[0])
        visible_height_pixels = abs(ylim[1] - ylim[0])
        visible_width_nm = visible_width_pixels * scale_per_pixel

        # Choose a nice round number for scale bar (~20% of visible width)
        target_length_nm = visible_width_nm * 0.2
        nice_lengths = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000]
        scale_bar_nm = min(nice_lengths, key=lambda x: abs(x - target_length_nm))

        # Convert to pixels
        scale_bar_pixels = scale_bar_nm / scale_per_pixel

        # Position: bottom right corner of current view with margin
        margin_x = visible_width_pixels * 0.05
        margin_y = visible_height_pixels * 0.05

        # Handle inverted Y axis
        y_min, y_max = min(ylim), max(ylim)
        x_min, x_max = min(xlim), max(xlim)

        x_start = x_max - margin_x - scale_bar_pixels
        y_pos = y_max - margin_y  # Bottom of view

        # Draw scale bar
        bar_height = visible_height_pixels * 0.02
        from matplotlib.patches import Rectangle
        import matplotlib.patheffects as path_effects

        # Black outline
        outline = Rectangle((x_start - 1, y_pos - bar_height / 2 - 1),
                            scale_bar_pixels + 2, bar_height + 2,
                            facecolor='black', edgecolor='none', zorder=100)
        outline._is_scale_bar = True
        self.map_ax.add_patch(outline)

        # White bar
        bar = Rectangle((x_start, y_pos - bar_height / 2),
                        scale_bar_pixels, bar_height,
                        facecolor='white', edgecolor='none', zorder=101)
        bar._is_scale_bar = True
        self.map_ax.add_patch(bar)

        # Format scale label
        if scale_bar_nm >= 1000:
            label = f'{scale_bar_nm / 1000:.0f} µm'
        else:
            label = f'{scale_bar_nm:.0f} {scale_unit}'

        # Add label above scale bar
        text = self.map_ax.text(x_start + scale_bar_pixels / 2, y_pos - bar_height - margin_y * 0.5,
                                label, ha='center', va='top', fontsize=9, fontweight='bold',
                                color='white', zorder=102,
                                path_effects=[path_effects.withStroke(linewidth=2, foreground='black')])
        text._is_scale_bar = True

    def _remove_scale_bar(self):
        """Remove existing scale bar elements"""
        # Remove patches
        for artist in self.map_ax.patches[:]:
            if hasattr(artist, '_is_scale_bar') and artist._is_scale_bar:
                artist.remove()
        # Remove texts
        for artist in self.map_ax.texts[:]:
            if hasattr(artist, '_is_scale_bar') and artist._is_scale_bar:
                artist.remove()

    def _update_scale_bar(self):
        """Update scale bar after view change"""
        self._add_scale_bar()
        self.map_canvas.draw_idle()


    def get_energy_axis(self):
        """Get energy axis from current data"""
        if self.current_data is None:
            return None

        if hasattr(self.current_data, 'axes_manager'):
            # Find the signal axis (energy)
            if len(self.current_data.axes_manager.signal_axes) > 0:
                return self.current_data.axes_manager.signal_axes[0].axis

        # Fallback: create channel numbers
        if len(self.current_data.data.shape) == 3:
            return np.arange(self.current_data.data.shape[2])
        elif len(self.current_data.data.shape) == 1:
            return np.arange(len(self.current_data.data))

        return None

    # ==================== FILE IMPORT ====================

    def on_import_sem(self, event):
        """Import SEM image"""
        with wx.FileDialog(self, "Open SEM Image",
                          wildcard="All files (*.*)|*.*|TIFF files (*.tif;*.tiff)|*.tif;*.tiff|"
                                  "HDF5 files (*.hdf5;*.h5)|*.hdf5;*.h5|"
                                  "BMP files (*.bmp)|*.bmp|PNG files (*.png)|*.png",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            file_path = fileDialog.GetPath()

        self.load_file(file_path, 'sem')

    def on_import_edx_spectrum(self, event):
        """Import single EDX spectrum"""
        with wx.FileDialog(self, "Open EDX Spectrum",
                          wildcard="All files (*.*)|*.*|EMSA files (*.emsa;*.ems)|*.emsa;*.ems|"
                                  "SPD files (*.spd)|*.spd|MSA files (*.msa)|*.msa",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            file_path = fileDialog.GetPath()

        self.load_file(file_path, 'edx_spectrum')

    def on_import_edx_map(self, event):
        """Import EDX map or spectrum image"""
        wildcard = "All supported files|*.hdf5;*.h5;*.bcf;*.rpl;*.raw;*.emd;*.ser|" \
                   "HDF5 files (*.hdf5)|*.hdf5|" \
                   "HDF5 files (*.h5)|*.h5|" \
                   "Bruker BCF files (*.bcf)|*.bcf|" \
                   "All files (*.*)|*.*"

        with wx.FileDialog(self, "Open EDX Map file",
                           wildcard=wildcard,
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                file_path = dlg.GetPath()
                self.load_file(file_path, 'EDX Map')

                # Create excel file and add to window.data
                if self.current_data is not None:
                    self.create_edx_map_output(file_path)

    def create_edx_map_output(self, file_path):
        """Create JSON and HDF5 files, and add EDX map data to parent window.Data"""
        import openpyxl
        from openpyxl.drawing.image import Image as OpenpyxlImage
        from io import BytesIO
        import shutil
        import json

        try:
            # ========== Initialize parent Data structure if needed ==========
            if self.parent is not None:
                if not hasattr(self.parent, 'Data'):
                    from libraries.ConfigFile import Init_Measurement_Data
                    self.parent.Data = Init_Measurement_Data(self.parent)

                if 'Core levels' not in self.parent.Data:
                    self.parent.Data['Core levels'] = {}

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            excel_path = os.path.join(os.path.dirname(file_path), f"{base_name}_EDX.xlsx")
            json_path = os.path.join(os.path.dirname(file_path), f"{base_name}_EDX.json")
            hdf5_copy_path = os.path.join(os.path.dirname(file_path), f"{base_name}_EDX.hdf5")

            # Set FilePath EARLY so it's available for other operations
            if self.parent is not None and hasattr(self.parent, 'Data'):
                self.parent.Data['FilePath'] = excel_path

                # Update current_file_path
                if hasattr(self.parent, 'current_file_path'):
                    self.parent.current_file_path = excel_path

                # Update Working_directory
                if hasattr(self.parent, 'Working_directory'):
                    self.parent.Working_directory = os.path.dirname(excel_path)

            # Copy original HDF5 file
            if file_path.lower().endswith(('.hdf5', '.h5')):
                shutil.copy2(file_path, hdf5_copy_path)
                print(f"HDF5 copy saved to: {hdf5_copy_path}")

            # Create workbook
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            # Get energy range
            energy_axis = self.get_energy_axis()
            if energy_axis is not None:
                energy_min = f"{np.min(energy_axis):.2f}"
                energy_max = f"{np.max(energy_axis):.2f}"
                energy_range = f"{energy_min} - {energy_max} keV"
            else:
                energy_range = "N/A"

            # Create sum spectrum
            sum_signal = self.current_data.sum()
            spectrum_data = sum_signal.data

            # EDX~Plot sheet
            ws_plot = wb.create_sheet("EDX~Plot")
            ws_plot.append(['Energy (keV)', 'Intensity', f'Range: {energy_range}'])

            for i, intensity in enumerate(spectrum_data):
                if energy_axis is not None:
                    ws_plot.append([f"{energy_axis[i]:.2f}", f"{intensity:.2f}"])
                else:
                    ws_plot.append([f"{i:.2f}", f"{intensity:.2f}"])

            # EDX~Map sheet
            ws_map = wb.create_sheet("EDX~Map")
            map_data = np.sum(self.current_data.data, axis=2)

            ws_map.append([f'EDX Intensity Map - Range: {energy_range}'])
            ws_map.append([''] * (map_data.shape[1] + 1))

            for row in map_data:
                ws_map.append([f"{val:.2f}" for val in row])

            # Create and save map image at 100 dpi
            fig, ax = plt.subplots(figsize=(map_data.shape[1] / 100, map_data.shape[0] / 100), dpi=100)
            im = ax.imshow(map_data, cmap=self.current_cmap)
            ax.set_title(f'EDX Map - {energy_range}')
            plt.colorbar(im, ax=ax)
            ax.axis('off')

            img_buffer = BytesIO()
            fig.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
            img_buffer.seek(0)
            plt.close(fig)

            img = OpenpyxlImage(img_buffer)
            ws_map.add_image(img, f'A{map_data.shape[0] + 5}')

            # Save Excel file
            wb.save(excel_path)
            print(f"EDX data exported to: {excel_path}")

            # ========== Add to parent window.Data['Core levels'] ==========
            if self.parent is not None and hasattr(self.parent, 'Data'):
                if 'Core levels' not in self.parent.Data:
                    self.parent.Data['Core levels'] = {}

                energy_values = energy_axis if energy_axis is not None else np.arange(len(spectrum_data))

                # Add EDX~Plot
                self.parent.Data['Core levels']['EDX~Plot'] = {
                    'Name': 'EDX~Plot',
                    'B.E.': list(energy_values),
                    'Raw Data': list(spectrum_data),
                    '_EDX_display_max': 20,
                    '_EDX_type': 'plot',
                    'Background': {}
                }

                # Add EDX~Map
                self.parent.Data['Core levels']['EDX~Map'] = {
                    'Name': 'EDX~Map',
                    'Map_Intensity': map_data.tolist(),
                    'Map_Shape': list(map_data.shape),
                    'Energy_Range': energy_range,
                    '_EDX_type': 'map',
                    '_HDF5_Path': hdf5_copy_path if os.path.exists(hdf5_copy_path) else file_path
                }

                # FilePath was already set at the beginning, now update UI
                # Update status bar
                if hasattr(self.parent, 'SetStatusText'):
                    self.parent.SetStatusText(f"Working Directory: {os.path.dirname(excel_path)}", 0)

                # Update window title
                if hasattr(self.parent, 'SetTitle'):
                    self.parent.SetTitle(f"KherveFitting - {os.path.basename(excel_path)}")

                # Add sheets to combobox
                for sheet_name in ['EDX~Plot', 'EDX~Map']:
                    if sheet_name not in [self.parent.sheet_combobox.GetString(i)
                                          for i in range(self.parent.sheet_combobox.GetCount())]:
                        self.parent.sheet_combobox.Append(sheet_name)



            # ========== Create JSON file ==========
            json_data = {
                'FilePath': excel_path,
                'Core levels': {}
            }

            # Add EDX~Plot to JSON
            json_data['Core levels']['EDX~Plot'] = {
                'Name': 'EDX~Plot',
                'B.E.': [float(f"{v:.2f}") for v in energy_values],
                'Raw Data': [float(f"{v:.2f}") for v in spectrum_data],
                '_EDX_display_max': 20,
                '_EDX_type': 'plot'
            }

            # Add EDX~Map to JSON
            json_data['Core levels']['EDX~Map'] = {
                'Name': 'EDX~Map',
                'Map_Intensity': [[float(f"{val:.2f}") for val in row] for row in map_data],
                'Map_Shape': list(map_data.shape),
                'Energy_Range': energy_range,
                '_EDX_type': 'map',
                '_HDF5_Path': hdf5_copy_path if os.path.exists(hdf5_copy_path) else file_path
            }

            with open(json_path, 'w') as jf:
                json.dump(json_data, jf, indent=2)
            print(f"JSON data saved to: {json_path}")

            wx.MessageBox(f"EDX data exported to:\n{excel_path}\n\nJSON: {json_path}\n" +
                          (f"HDF5: {hdf5_copy_path}" if os.path.exists(hdf5_copy_path) else ""),
                          "Export Complete", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error creating EDX output:\n{str(e)}",
                          "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_import_bcf(self, event):
        """Import Bruker BCF file"""
        with wx.FileDialog(self, "Open Bruker BCF File",
                          wildcard="BCF files (*.bcf)|*.bcf",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            file_path = fileDialog.GetPath()

        self.load_file(file_path, 'bcf')

    def load_file(self, file_path, data_type):
        """Generic file loader with automatic reader selection"""
        try:
            loaded_data = None

            # List of readers to try
            readers_to_try = ['HSPY', 'USID', 'Delmic', 'EMSA', 'Bruker']

            # Try each reader in order
            for reader in readers_to_try:
                try:
                    loaded_data = hs.load(file_path, reader=reader)
                    print(f"Successfully loaded with {reader} reader")
                    break
                except Exception as e:
                    print(f"Failed with {reader} reader: {e}")
                    continue

            # If all readers failed, try without specifying reader
            if loaded_data is None:
                try:
                    loaded_data = hs.load(file_path)
                    print("Successfully loaded with default reader")
                except Exception as e:
                    wx.MessageBox(f"Could not load file with any available reader.\nError: {str(e)}",
                                  "Load Error", wx.OK | wx.ICON_ERROR)
                    return

            # Handle list of signals
            if isinstance(loaded_data, list):
                if len(loaded_data) == 1:
                    loaded_data = loaded_data[0]
                else:
                    self.add_multiple_signals_to_tree(loaded_data, file_path, data_type)
                    return

            # Store signal
            self.loaded_signals.append({
                'data': loaded_data,
                'filename': os.path.basename(file_path),
                'path': file_path,
                'type': data_type
            })

            self.current_data = loaded_data

            # Update data browser if open
            if hasattr(self, 'data_browser_window') and self.data_browser_window is not None:
                self.data_browser_window.refresh_tree()

            self.plot_current_map()
            self.update_info_text(loaded_data)



            # Extract elements if available
            self.extract_elements(loaded_data)

        except Exception as e:
            wx.MessageBox(f"Error loading file:\n{str(e)}",
                         "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def add_signal_to_tree(self, signal, file_path, data_type):
        """Add a signal to the data tree"""
        filename = os.path.basename(file_path)

        # Get signal title if available
        title = filename
        if hasattr(signal, 'metadata') and hasattr(signal.metadata, 'General'):
            if hasattr(signal.metadata.General, 'title') and signal.metadata.General.title:
                title = signal.metadata.General.title

        # Create tree item
        item_text = f"{title} ({signal.data.shape})"
        item = self.data_tree.AppendItem(self.tree_root, item_text)
        self.data_tree.SetItemData(item, {'type': data_type, 'data': signal, 'path': file_path})

        # Add metadata as children
        self.add_metadata_to_tree(item, signal)

        # Expand tree
        self.data_tree.Expand(self.tree_root)
        self.data_tree.SelectItem(item)

    def add_multiple_signals_to_tree(self, signals, file_path, data_type):
        """Add multiple signals to tree"""
        filename = os.path.basename(file_path)

        # Create parent item for file
        file_item = self.data_tree.AppendItem(self.tree_root, filename)
        self.data_tree.SetItemData(file_item, {'type': 'container', 'path': file_path})

        for i, signal in enumerate(signals):
            title = f"Signal {i}"
            if hasattr(signal, 'metadata') and hasattr(signal.metadata, 'General'):
                if hasattr(signal.metadata.General, 'title') and signal.metadata.General.title:
                    title = signal.metadata.General.title

            item_text = f"{title} ({signal.data.shape})"
            item = self.data_tree.AppendItem(file_item, item_text)
            self.data_tree.SetItemData(item, {'type': data_type, 'data': signal, 'path': file_path})

            # Add metadata
            self.add_metadata_to_tree(item, signal)

        self.data_tree.Expand(self.tree_root)
        self.data_tree.Expand(file_item)

        # Select first signal
        if signals:
            self.current_data = signals[0]
            self.plot_current_map()
            self.update_info_text(signals[0])
            self.extract_elements(signals[0])

    def add_metadata_to_tree(self, parent_item, signal):
        """Add signal metadata as tree children"""
        if not hasattr(signal, 'metadata'):
            return

        # Add shape info
        shape_item = self.data_tree.AppendItem(parent_item, f"Shape: {signal.data.shape}")

        # Add data type
        dtype_item = self.data_tree.AppendItem(parent_item, f"Dtype: {signal.data.dtype}")

        # Add axes info
        if hasattr(signal, 'axes_manager'):
            axes_item = self.data_tree.AppendItem(parent_item, "Axes")
            for i, axis in enumerate(signal.axes_manager.signal_axes):
                axis_text = f"Signal {i}: {axis.name} [{axis.units}] - {axis.size} pts"
                self.data_tree.AppendItem(axes_item, axis_text)
            for i, axis in enumerate(signal.axes_manager.navigation_axes):
                axis_text = f"Nav {i}: {axis.name} [{axis.units}] - {axis.size} pts"
                self.data_tree.AppendItem(axes_item, axis_text)

        # Add elements if available
        if hasattr(signal, 'metadata') and hasattr(signal.metadata, 'Sample'):
            if hasattr(signal.metadata.Sample, 'elements'):
                elements = signal.metadata.Sample.elements
                elem_item = self.data_tree.AppendItem(parent_item, f"Elements: {', '.join(elements)}")

    def extract_elements(self, signal):
        """Extract element information from signal"""
        if hasattr(signal, 'metadata') and hasattr(signal.metadata, 'Sample'):
            if hasattr(signal.metadata.Sample, 'elements'):
                self.selected_elements = list(signal.metadata.Sample.elements)
            else:
                self.selected_elements = []
        else:
            self.selected_elements = []

    def update_info_text(self, signal):
        """Update info window if open"""
        if hasattr(self, 'info_window') and self.info_window is not None:
            try:
                self.info_window.update_info(signal)
            except:
                pass

    # ==================== MAP PLOTTING ====================

    def plot_current_map(self):
        """Plot the current data as a map"""
        if self.current_data is None:
            return

        # Clear previous colorbar
        if self.current_colorbar is not None:
            try:
                self.current_colorbar.remove()
            except:
                pass
            self.current_colorbar = None

        self.map_ax.clear()
        data = self.current_data.data
        cmap = self.current_cmap

        if len(data.shape) == 2:
            sum_image = data
            title = 'SEM/EDX Image'
        elif len(data.shape) == 3:
            sum_image = np.sum(data, axis=2)
            title = f'Sum Image ({data.shape[0]}×{data.shape[1]} px, {data.shape[2]} ch)'
        elif len(data.shape) == 4:
            sum_image = np.sum(data, axis=(2, 3))
            title = f'Sum Image (4D: {data.shape})'
        else:
            self.map_ax.text(0.5, 0.5, f'Unsupported shape: {data.shape}',
                             ha='center', va='center', transform=self.map_ax.transAxes)
            self.map_canvas.draw()
            return

        im = self.map_ax.imshow(sum_image, cmap=cmap, origin='upper')

        # Remove axis labels and title - use scale bar instead
        self.map_ax.set_xlabel('')
        self.map_ax.set_ylabel('')
        self.map_ax.set_title('')
        self.map_ax.set_xticks([])
        self.map_ax.set_yticks([])

        # Add scale bar if scale information is available
        # self._add_scale_bar(sum_image.shape)
        self._add_scale_bar()

        # Add colorbar matching heatmap style (full height on right side)
        from mpl_toolkits.axes_grid1 import make_axes_locatable
        divider = make_axes_locatable(self.map_ax)
        cax = divider.append_axes("right", size="5%", pad=0.05)
        self.current_colorbar = self.map_figure.colorbar(im, cax=cax)

        self.map_figure.tight_layout(pad=0.5)

        # Add scale bar
        self._add_scale_bar()

        self.map_canvas.draw()

        # Plot sum spectrum in parent KherveFitting window
        self.plot_sum_spectrum_to_parent()

        # # Populate peak fitting grid with quantification
        # self._populate_edx_grid(energy, spectrum, "Sum Spectrum")

        self.reinitialize_selectors()

    def plot_sum_spectrum(self):
        """Plot sum spectrum from current data"""
        if self.current_data is None:
            return

        data = self.current_data.data
        if len(data.shape) != 3:
            return

        # Sum over spatial dimensions
        spectrum = np.sum(data, axis=(0, 1))
        energy = self.get_energy_axis()

        self.spectrum_ax.clear()
        self.spectrum_ax.plot(energy, spectrum, 'b-', linewidth=0.8, label='Sum Spectrum')
        self.spectrum_ax.set_xlabel('Energy (keV)')
        self.spectrum_ax.set_ylabel('Counts')
        self.spectrum_ax.set_title('Total Sum Spectrum')
        self.spectrum_ax.grid(True, alpha=0.3)
        self.spectrum_figure.tight_layout()
        self.spectrum_canvas.draw()

    def on_tree_select(self, event):
        """Handle tree selection"""
        item = event.GetItem()
        data = self.data_tree.GetItemData(item)

        if data is None:
            return

        if data.get('type') == 'container':
            return

        signal = data.get('data')
        if signal is not None:
            self.current_data = signal
            self.plot_current_map()
            self.update_info_text(signal)
            self.extract_elements(signal)

    def on_set_elements(self, event):
        """Set elements using periodic table dialog"""
        # Get currently selected elements
        current_elements = getattr(self, 'selected_elements', [])

        dlg = PeriodicTableDialog(self, current_elements)

        if dlg.ShowModal() == wx.ID_OK:
            self.selected_elements = dlg.get_selected_elements()

        dlg.Destroy()

    def on_plot_maps(self, event):
        """Plot element maps as RGB composite overlay"""
        if not hasattr(self, 'selected_elements') or not self.selected_elements:
            wx.MessageBox("No elements selected.\nUse 'Set Elements' first.",
                          "No Selection", wx.OK | wx.ICON_WARNING)
            return

        if self.current_data is None:
            wx.MessageBox("No data loaded", "No Data", wx.OK | wx.ICON_WARNING)
            return

        selected_elements = self.selected_elements

        if len(selected_elements) > 7:
            wx.MessageBox("Maximum 7 elements can be displayed.\nPlease select fewer elements.",
                          "Too Many Elements", wx.OK | wx.ICON_WARNING)
            return

        try:
            element_maps = {}
            for element in selected_elements:
                element_map = self.get_element_map(element)
                if element_map is not None:
                    element_maps[element] = element_map

            if not element_maps:
                wx.MessageBox("Could not extract any element maps", "Error", wx.OK | wx.ICON_ERROR)
                return

            self.create_composite_map(element_maps)

        except Exception as e:
            wx.MessageBox(f"Error plotting element maps:\n{str(e)}",
                          "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def get_element_map(self, element):
        """Get element map using HyperSpy or manual integration"""
        element_map = None

        # Method 1: Try HyperSpy's get_lines_intensity
        if hasattr(self.current_data, 'get_lines_intensity'):
            try:
                # Set element if not already set
                if hasattr(self.current_data, 'set_elements'):
                    current_elements = []
                    if hasattr(self.current_data.metadata, 'Sample') and \
                            hasattr(self.current_data.metadata.Sample, 'elements'):
                        current_elements = list(self.current_data.metadata.Sample.elements)

                    if element not in current_elements:
                        current_elements.append(element)
                        self.current_data.set_elements(current_elements)
                        self.current_data.add_lines()

                # Try different X-ray lines
                for line in ['Ka', 'La', 'Ma', 'Kb', 'Lb']:
                    try:
                        line_name = f"{element}_{line}"
                        intensity_maps = self.current_data.get_lines_intensity([line_name])
                        if intensity_maps and len(intensity_maps) > 0:
                            element_map = intensity_maps[0].data
                            break
                    except:
                        continue
            except Exception as e:
                print(f"HyperSpy method failed for {element}: {e}")

        # Method 2: Manual energy window integration
        if element_map is None and len(self.current_data.data.shape) == 3:
            element_map = self.get_element_map_manual(element)

        return element_map

    def get_element_map_manual(self, element):
        """Manually integrate around element's characteristic X-ray energy"""
        xray_energies = {
            'C': 0.277, 'N': 0.392, 'O': 0.525, 'F': 0.677, 'Na': 1.041, 'Mg': 1.254,
            'Al': 1.487, 'Si': 1.740, 'P': 2.013, 'S': 2.307, 'Cl': 2.622, 'K': 3.313,
            'Ca': 3.691, 'Ti': 4.510, 'V': 4.952, 'Cr': 5.414, 'Mn': 5.898, 'Fe': 6.403,
            'Co': 6.930, 'Ni': 7.478, 'Cu': 8.048, 'Zn': 8.638, 'Ga': 9.251, 'Ge': 9.886,
            'As': 10.543, 'Se': 11.222, 'Br': 11.924, 'Rb': 13.395, 'Sr': 14.165,
            'Y': 14.958, 'Zr': 15.775, 'Nb': 16.615, 'Mo': 17.479, 'Ag': 22.163,
            'Cd': 23.174, 'In': 24.210, 'Sn': 25.271, 'Sb': 26.359, 'Te': 27.472,
            'I': 28.612, 'Ba': 32.194, 'La': 33.442, 'Ce': 34.720, 'W': 59.318,
            'Pt': 66.831, 'Au': 68.803, 'Pb': 74.969, 'Bi': 77.107
        }

        if element not in xray_energies:
            return None

        target_energy = xray_energies[element]
        energy_axis = self.get_energy_axis()
        if energy_axis is None:
            return None

        # Find energy window (±0.15 keV around peak)
        window = 0.15
        mask = (energy_axis >= target_energy - window) & (energy_axis <= target_energy + window)

        if not np.any(mask):
            window = 0.3
            mask = (energy_axis >= target_energy - window) & (energy_axis <= target_energy + window)

        if not np.any(mask):
            return None

        data = self.current_data.data
        element_map = np.sum(data[:, :, mask], axis=2)

        return element_map

    def create_composite_map(self, element_maps):
        """Create RGB composite map with legend - no colorbar"""
        # Clear previous colorbar
        if self.current_colorbar is not None:
            try:
                self.current_colorbar.remove()
            except:
                pass
            self.current_colorbar = None

        colors = [
            (1.0, 0.0, 0.0),  # Red
            (0.0, 1.0, 0.0),  # Green
            (0.0, 0.0, 1.0),  # Blue
            (1.0, 1.0, 0.0),  # Yellow
            (1.0, 0.0, 1.0),  # Magenta
            (0.0, 1.0, 1.0),  # Cyan
            (1.0, 0.5, 0.0),  # Orange
        ]

        elements = list(element_maps.keys())
        first_map = list(element_maps.values())[0]
        height, width = first_map.shape

        rgb_image = np.zeros((height, width, 3), dtype=np.float64)
        self.element_colors = {}

        for idx, element in enumerate(elements):
            element_map = element_maps[element]
            color = colors[idx % len(colors)]
            self.element_colors[element] = color

            map_min = np.min(element_map)
            map_max = np.max(element_map)
            if map_max > map_min:
                normalized_map = (element_map - map_min) / (map_max - map_min)
            else:
                normalized_map = np.zeros_like(element_map)

            rgb_image[:, :, 0] += normalized_map * color[0]
            rgb_image[:, :, 1] += normalized_map * color[1]
            rgb_image[:, :, 2] += normalized_map * color[2]

        for c in range(3):
            channel_max = np.max(rgb_image[:, :, c])
            if channel_max > 0:
                rgb_image[:, :, c] /= channel_max

        rgb_image = np.clip(rgb_image, 0, 1)

        self.map_ax.clear()
        self.map_ax.imshow(rgb_image, origin='upper')

        # Remove axis labels and title - use scale bar instead
        self.map_ax.set_xlabel('')
        self.map_ax.set_ylabel('')
        self.map_ax.set_xticks([])
        self.map_ax.set_yticks([])

        # Add scale bar
        self._add_scale_bar(rgb_image.shape[:2])

        # Legend only - no colorbar for composite
        legend_elements = []
        for element, color in self.element_colors.items():
            from matplotlib.patches import Patch
            legend_elements.append(Patch(facecolor=color, edgecolor='black', label=element))

        self.map_ax.legend(handles=legend_elements, loc='upper right', fontsize=8, framealpha=0.8)
        self.map_ax.set_title(f'Composite: {", ".join(elements)}')

        self.map_figure.tight_layout()
        self.map_canvas.draw()

        self.reinitialize_selectors()

    def plot_sum_spectrum_to_parent(self):
        """Plot sum spectrum in KherveFitting main window with element labels"""
        if self.current_data is None:
            return

        data = self.current_data.data
        if len(data.shape) != 3:
            return

        spectrum = np.sum(data, axis=(0, 1))
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self.parent.ax.clear()
            self.parent.ax.plot(energy, spectrum, 'k-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title('EDX Sum Spectrum')

            # Set Y-axis to scientific format
            self.parent.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))

            # Add element peak labels
            self.add_peak_labels(self.parent.ax, energy, spectrum)

            # Get stored X max or default to 20
            current_sheet = self.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                if current_sheet in self.parent.Data['Core levels']:
                    display_x_max = self.parent.Data['Core levels'][current_sheet].get('_EDX_display_max', 20)
                else:
                    display_x_max = 20
            else:
                display_x_max = 20

            self.parent.ax.set_xlim(0, display_x_max)
            self.parent.ax.set_ylim(np.min(spectrum) * 0.95, np.max(spectrum) * 1.1)

            self.parent.canvas.draw()

            # Populate peak fitting grid with quantification
            self._populate_edx_grid(energy, spectrum, "Sum Spectrum")

    def add_peak_labels_OLD(self, ax, energy, spectrum):
        """Add element peak labels above peaks using ExSpy's find_peaks1D_ohaver"""
        try:
            # Create a temporary HyperSpy signal from the spectrum data
            import hyperspy.api as hs
            from exspy.material import elements

            # Create signal with proper energy axis
            temp_signal = hs.signals.Signal1D(spectrum)
            temp_signal.axes_manager[0].scale = energy[1] - energy[0] if len(energy) > 1 else 0.01
            temp_signal.axes_manager[0].offset = energy[0]
            temp_signal.axes_manager[0].units = 'keV'
            temp_signal.axes_manager[0].name = 'Energy'

            # Set signal type to EDS
            temp_signal.set_signal_type("EDS_SEM")

            # Find peaks using ExSpy's find_peaks1D_ohaver
            from exspy.utils.eds import get_xray_lines_near_energy

            # Get threshold from multiple sources (priority order)
            threshold = 0.02  # Default 2%

            # 1. Check if stored in parent's Data structure
            if hasattr(self, 'parent') and self.parent and hasattr(self.parent, 'Data'):
                current_sheet = self.parent.sheet_combobox.GetValue()
                if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                    if current_sheet in self.parent.Data['Core levels']:
                        stored_threshold = self.parent.Data['Core levels'][current_sheet].get('_EDX_sensitivity')
                        if stored_threshold is not None:
                            threshold = stored_threshold
                            print(f"Using stored threshold from Data: {threshold:.4f}")

            # 2. Check if set directly on self (overrides stored value)
            if hasattr(self, 'peak_threshold'):
                threshold = self.peak_threshold
                print(f"Using direct threshold from self: {threshold:.4f}")

            # Calculate amp_thresh from sensitivity threshold
            amp_thresh = np.max(spectrum) * threshold
            print(f"Peak detection with threshold: {threshold:.4f}, amp_thresh: {amp_thresh:.2f}")

            peak_data = temp_signal.find_peaks1D_ohaver(
                maxpeakn=100,
                medfilt_radius=5,
                peakgroup=10,
                amp_thresh=amp_thresh,
                slope_thresh=0
            )

            # Get y-axis range for positioning labels
            y_min, y_max = ax.get_ylim()
            y_range = y_max - y_min

            # For each detected peak, find possible X-ray lines
            labeled_positions = []

            # Check if peak_data is valid
            if peak_data is not None and isinstance(peak_data, np.ndarray) and len(peak_data) > 0:
                # Extract the actual peaks array
                peaks = peak_data[0] if len(peak_data.shape) > 1 or peak_data.dtype == object else peak_data

                print(f"Found {len(peaks)} peaks")

                for peak_tuple in peaks:
                    try:
                        # Each peak_tuple is (position, height, width)
                        peak_energy = float(peak_tuple[0])
                        peak_height_val = float(peak_tuple[1])
                        peak_width = float(peak_tuple[2])

                        # Skip if too close to already labeled position
                        too_close = False
                        for labeled_energy in labeled_positions:
                            if abs(peak_energy - labeled_energy) < 0.1:
                                too_close = True
                                break

                        if too_close:
                            continue

                        # Find X-ray lines near this energy
                        try:
                            lines_list = get_xray_lines_near_energy(peak_energy, only_lines=None, width=0.15)

                            if not lines_list:
                                continue

                            # lines_list contains strings like 'Si_Ka', 'Fe_La', etc.
                            # Find the closest matching line by looking up their energies
                            closest_line = None
                            min_diff = float('inf')

                            for line_str in lines_list:
                                # Parse the line string (e.g., 'Si_Ka' -> element='Si', line='Ka')
                                parts = line_str.split('_')
                                if len(parts) != 2:
                                    continue

                                element_symbol = parts[0]
                                line_type = parts[1]

                                # Get the element and look up the line energy
                                try:
                                    element_obj = elements[element_symbol]

                                    # Get line energy from element's X-ray lines
                                    if hasattr(element_obj, 'Atomic_properties') and \
                                            hasattr(element_obj.Atomic_properties, 'Xray_lines'):
                                        xray_lines = element_obj.Atomic_properties.Xray_lines

                                        if line_type in xray_lines:
                                            line_energy = xray_lines[line_type].energy_keV
                                            diff = abs(line_energy - peak_energy)

                                            if diff < min_diff:
                                                min_diff = diff
                                                closest_line = (element_symbol, line_type)
                                except (KeyError, AttributeError):
                                    continue

                            if closest_line and min_diff < 0.15:
                                element, line_type = closest_line

                                # Format line type with Greek letters
                                line_type = str(line_type)
                                line_type = line_type.replace('Ka', 'Kα').replace('Kb', 'Kβ')
                                line_type = line_type.replace('La', 'Lα').replace('Lb', 'Lβ')
                                line_type = line_type.replace('Ma', 'Mα').replace('Mb', 'Mβ')

                                label = f"{element}\n{line_type}"

                                # Get peak height at this energy from spectrum
                                idx = np.argmin(np.abs(energy - peak_energy))
                                spectrum_height = spectrum[idx]

                                # Position label as percentage of Y-axis (closer to peak top)
                                # Calculate peak position as percentage of plot range
                                peak_pct = (spectrum_height - y_min) / y_range
                                # Place label at peak percentage + 2% of plot height
                                label_y = y_min + (peak_pct + 0.02) * y_range

                                # Add label
                                ax.text(peak_energy, label_y, label,
                                        rotation=0,
                                        verticalalignment='bottom',
                                        horizontalalignment='center',
                                        fontsize=7,
                                        color='black',
                                        alpha=1.0)

                                labeled_positions.append(peak_energy)
                                print(f"Labeled peak: {element} {line_type} at {peak_energy:.2f} keV")

                        except Exception as e:
                            print(f"Could not identify line at {peak_energy:.2f} keV: {e}")
                            continue

                    except Exception as e:
                        print(f"Error processing peak: {e}")
                        continue

                print(f"Total peaks labeled: {len(labeled_positions)}")

        except Exception as e:
            print(f"Error in peak labeling: {e}")
            import traceback
            traceback.print_exc()

    def add_peak_labels(self, ax, energy, spectrum):
        """Add element peak labels above peaks with improved identification"""
        try:
            import hyperspy.api as hs
            from exspy.material import elements
            from exspy.utils.eds import get_xray_lines_near_energy

            # Create signal with proper energy axis
            temp_signal = hs.signals.Signal1D(spectrum)
            temp_signal.axes_manager[0].scale = energy[1] - energy[0] if len(energy) > 1 else 0.01
            temp_signal.axes_manager[0].offset = energy[0]
            temp_signal.axes_manager[0].units = 'keV'
            temp_signal.axes_manager[0].name = 'Energy'
            temp_signal.set_signal_type("EDS_SEM")

            # Get threshold
            threshold = 0.02  # Default 2%
            if hasattr(self, 'parent') and self.parent and hasattr(self.parent, 'Data'):
                current_sheet = self.parent.sheet_combobox.GetValue()
                if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.Data:
                    if current_sheet in self.parent.Data['Core levels']:
                        stored_threshold = self.parent.Data['Core levels'][current_sheet].get('_EDX_sensitivity')
                        if stored_threshold is not None:
                            threshold = stored_threshold

            if hasattr(self, 'peak_threshold'):
                threshold = self.peak_threshold

            amp_thresh = np.max(spectrum) * threshold
            print(f"Peak detection with threshold: {threshold:.4f}, amp_thresh: {amp_thresh:.2f}")

            peak_data = temp_signal.find_peaks1D_ohaver(
                maxpeakn=100,
                medfilt_radius=5,
                peakgroup=10,
                amp_thresh=amp_thresh,
                slope_thresh=0
            )

            y_min, y_max = ax.get_ylim()
            y_range = y_max - y_min

            if peak_data is None or not isinstance(peak_data, np.ndarray) or len(peak_data) == 0:
                return

            peaks = peak_data[0] if len(peak_data.shape) > 1 or peak_data.dtype == object else peak_data
            print(f"Found {len(peaks)} peaks")

            # Build peak database with energies and heights
            peak_list = []
            for peak_tuple in peaks:
                try:
                    peak_energy = float(peak_tuple[0])
                    peak_height = float(peak_tuple[1])
                    peak_list.append((peak_energy, peak_height))
                except (TypeError, ValueError, IndexError):
                    continue

            # Sort by intensity (highest first)
            peak_list.sort(key=lambda x: x[1], reverse=True)

            # Identify peaks using intensity-aware matching
            identified_peaks = self._identify_peaks_with_ratios(peak_list, elements)

            # Add labels to plot
            for peak_energy, element, line_type in identified_peaks:
                # Format line type with Greek letters
                line_display = str(line_type)
                line_display = line_display.replace('Ka', 'Kα').replace('Kb', 'Kβ')
                line_display = line_display.replace('La', 'Lα').replace('Lb', 'Lβ')
                line_display = line_display.replace('Ma', 'Mα').replace('Mb', 'Mβ')

                label = f"{element}\n{line_display}"

                # Get peak height at this energy
                idx = np.argmin(np.abs(energy - peak_energy))
                spectrum_height = spectrum[idx]

                # Position label
                peak_pct = (spectrum_height - y_min) / y_range
                label_y = y_min + (peak_pct + 0.02) * y_range

                ax.text(peak_energy, label_y, label,
                        rotation=0,
                        verticalalignment='bottom',
                        horizontalalignment='center',
                        fontsize=7,
                        color='black',
                        fontweight='bold')

        except Exception as e:
            print(f"Error adding peak labels: {e}")
            import traceback
            traceback.print_exc()

    def _identify_peaks_with_ratios(self, peak_list, elements):
        """Identify peaks considering expected intensity ratios and common elements"""
        from exspy.utils.eds import get_xray_lines_near_energy

        identified = []
        used_energies = set()

        # Common elements in EDX analysis - prioritize these heavily
        common_elements = ['O', 'C', 'N', 'F', 'Na', 'Mg', 'Al', 'Si', 'P', 'S', 'Cl', 'K', 'Ca',
                           'Ti', 'V', 'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn', 'Ga', 'Ge', 'As',
                           'Se', 'Br', 'Rb', 'Sr', 'Y', 'Zr', 'Nb', 'Mo', 'Ag', 'Cd', 'In', 'Sn',
                           'Sb', 'Te', 'I', 'Ba', 'W', 'Pt', 'Au', 'Pb', 'Bi']

        # Expected K alpha/beta intensity ratios by atomic number range
        k_ratios = {
            'light': 8.0,  # Z < 20
            'medium': 7.0,  # 20 <= Z < 50
            'heavy': 5.0  # Z >= 50
        }

        # Expected L alpha/beta ratio
        l_ratio = 3.0

        # For each peak, find ALL possible element matches
        peak_candidates = {}  # peak_energy -> [(element, line, score), ...]

        for peak_energy, peak_height in peak_list:
            peak_candidates[peak_energy] = []

            try:
                lines_list = get_xray_lines_near_energy(peak_energy, only_lines=None, width=0.15)
                if not lines_list:
                    continue

                for line_str in lines_list:
                    parts = line_str.split('_')
                    if len(parts) != 2:
                        continue

                    element_symbol = parts[0]
                    line_type = parts[1]

                    try:
                        element_obj = elements[element_symbol]
                        if not (hasattr(element_obj, 'Atomic_properties') and
                                hasattr(element_obj.Atomic_properties, 'Xray_lines')):
                            continue

                        xray_lines = element_obj.Atomic_properties.Xray_lines
                        if line_type not in xray_lines:
                            continue

                        line_energy = xray_lines[line_type].energy_keV
                        energy_diff = abs(line_energy - peak_energy)

                        if energy_diff < 0.15:
                            # Base score on energy match
                            energy_score = 1.0 - (energy_diff / 0.15)

                            # Bonus for common elements
                            if element_symbol in common_elements:
                                energy_score *= 1.5  # 50% bonus

                            peak_candidates[peak_energy].append({
                                'element': element_symbol,
                                'line': line_type,
                                'energy': line_energy,
                                'score': energy_score,
                                'Z': element_obj.General_properties.Z
                            })

                    except (KeyError, AttributeError):
                        continue
            except:
                continue

        # Now find element pairs (alpha + beta)
        element_pairs = {}  # element -> {Ka: peak_energy, Kb: peak_energy, ...}

        for element_symbol in set([cand['element'] for peak_cands in peak_candidates.values() for cand in peak_cands]):
            element_pairs[element_symbol] = {'Ka': None, 'Kb': None, 'La': None, 'Lb': None}

            # Find Ka and Kb
            for peak_energy, peak_cands in peak_candidates.items():
                for cand in peak_cands:
                    if cand['element'] == element_symbol:
                        line = cand['line']
                        if line in ['Ka', 'Kb', 'La', 'Lb']:
                            if element_pairs[element_symbol][line] is None:
                                element_pairs[element_symbol][line] = {
                                    'peak_energy': peak_energy,
                                    'peak_height': next(h for e, h in peak_list if abs(e - peak_energy) < 0.001),
                                    'expected_energy': cand['energy'],
                                    'score': cand['score'],
                                    'Z': cand['Z']
                                }

        # Score element pairs
        pair_scores = []

        for element_symbol, lines in element_pairs.items():
            # Check K pair
            if lines['Ka'] and lines['Kb']:
                ka = lines['Ka']
                kb = lines['Kb']

                Z = ka['Z']
                if Z < 20:
                    expected_ratio = k_ratios['light']
                elif Z < 50:
                    expected_ratio = k_ratios['medium']
                else:
                    expected_ratio = k_ratios['heavy']

                measured_ratio = ka['peak_height'] / kb['peak_height']
                ratio_error = abs(measured_ratio - expected_ratio) / expected_ratio

                # Ratio match score (most important - 60%)
                ratio_score = max(0, 1.0 - ratio_error)

                # Energy match scores (40%)
                energy_score = (ka['score'] + kb['score']) / 2.0

                # Total score
                total_score = 0.6 * ratio_score + 0.4 * energy_score

                # Extra bonus if common element
                if element_symbol in common_elements:
                    total_score *= 1.3

                pair_scores.append({
                    'element': element_symbol,
                    'type': 'K',
                    'score': total_score,
                    'alpha': ka,
                    'beta': kb,
                    'ratio': measured_ratio,
                    'expected_ratio': expected_ratio
                })

                print(f"{element_symbol} K-pair: measured_ratio={measured_ratio:.2f}, expected={expected_ratio:.1f}, score={total_score:.3f}")

            # Check L pair
            if lines['La'] and lines['Lb']:
                la = lines['La']
                lb = lines['Lb']

                measured_ratio = la['peak_height'] / lb['peak_height']
                ratio_error = abs(measured_ratio - l_ratio) / l_ratio

                ratio_score = max(0, 1.0 - ratio_error)
                energy_score = (la['score'] + lb['score']) / 2.0
                total_score = 0.6 * ratio_score + 0.4 * energy_score

                if element_symbol in common_elements:
                    total_score *= 1.3

                pair_scores.append({
                    'element': element_symbol,
                    'type': 'L',
                    'score': total_score,
                    'alpha': la,
                    'beta': lb,
                    'ratio': measured_ratio,
                    'expected_ratio': l_ratio
                })

                print(f"{element_symbol} L-pair: measured_ratio={measured_ratio:.2f}, expected={l_ratio:.1f}, score={total_score:.3f}")

        # Sort pairs by score and add best matches
        pair_scores.sort(key=lambda x: x['score'], reverse=True)

        for pair in pair_scores:
            if pair['score'] < 0.4:  # Minimum threshold
                continue

            alpha_peak = pair['alpha']['peak_energy']
            beta_peak = pair['beta']['peak_energy']

            # Check if these peaks are already used
            if alpha_peak in used_energies or beta_peak in used_energies:
                continue

            # Add this pair
            element = pair['element']
            if pair['type'] == 'K':
                identified.append((alpha_peak, element, 'Ka'))
                identified.append((beta_peak, element, 'Kb'))
            else:
                identified.append((alpha_peak, element, 'La'))
                identified.append((beta_peak, element, 'Lb'))

            used_energies.add(alpha_peak)
            used_energies.add(beta_peak)

            print(f"Identified {element} {pair['type']}-pair with score {pair['score']:.3f}")

        # Second pass: unpaired significant peaks
        for peak_energy, peak_height in peak_list:
            if peak_energy in used_energies:
                continue

            # Only significant peaks
            if peak_height < 0.1 * peak_list[0][1]:
                continue

            # Find best candidate (prefer alpha lines and common elements)
            best_candidate = None
            best_score = 0

            if peak_energy in peak_candidates:
                for cand in peak_candidates[peak_energy]:
                    # Skip beta lines - they should have alphas
                    if cand['line'] in ['Kb', 'Lb', 'Mb']:
                        continue

                    score = cand['score']

                    # Strong preference for alpha lines
                    if cand['line'] in ['Ka', 'La', 'Ma']:
                        score *= 1.2

                    if score > best_score:
                        best_score = score
                        best_candidate = cand

            if best_candidate and best_score > 0.6:
                identified.append((peak_energy, best_candidate['element'], best_candidate['line']))
                used_energies.add(peak_energy)
                print(f"Second pass: {best_candidate['element']} {best_candidate['line']} at {peak_energy:.2f} keV")

        # Sort by energy
        identified.sort(key=lambda x: x[0])

        return identified



    def reinitialize_selectors(self):
        """Reinitialize selection tools after axes change"""
        # Reset selectors
        self.rect_selector = None
        self.line_selector = None

        # Recreate if mode is active
        if self.selection_mode == 'area':
            self.rect_selector = RectangleSelector(
                self.map_ax,
                self.on_area_select,
                useblit=True,
                props=dict(facecolor='white', edgecolor='white', alpha=0.3, fill=True),
                button=[1],
                minspanx=5,
                minspany=5,
                spancoords='pixels',
                interactive=True
            )
        elif self.selection_mode == 'line':
            self.line_selector = SpanSelector(
                self.map_ax,
                self.on_line_select,
                'horizontal',
                useblit=True,
                props=dict(facecolor='white', alpha=0.3),
                interactive=True
            )

    # ==================== EXPORT ====================

    def on_export_excel(self, event):
        """Export to Excel"""
        if self.current_data is None:
            wx.MessageBox("No data to export", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Export to Excel", wildcard="Excel files (*.xlsx)|*.xlsx",
                          style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            file_path = dlg.GetPath()

        self.export_to_excel(file_path)

    def on_export_csv(self, event):
        """Export to CSV"""
        if self.current_data is None:
            wx.MessageBox("No data to export", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Export to CSV", wildcard="CSV files (*.csv)|*.csv",
                          style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            file_path = dlg.GetPath()

        self.export_to_csv(file_path)

    def on_export_hdf5(self, event):
        """Export to HDF5"""
        if self.current_data is None:
            wx.MessageBox("No data to export", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Export to HDF5", wildcard="HDF5 files (*.hdf5)|*.hdf5",
                          style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            file_path = dlg.GetPath()

        try:
            self.current_data.save(file_path)
            wx.MessageBox(f"Data exported to:\n{file_path}", "Success", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Error exporting:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_export_map_image(self, event):
        """Export current map as image"""
        with wx.FileDialog(self, "Export Map Image",
                          wildcard="PNG files (*.png)|*.png|TIFF files (*.tiff)|*.tiff",
                          style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            file_path = dlg.GetPath()

        try:
            self.map_figure.savefig(file_path, dpi=300, bbox_inches='tight')
            wx.MessageBox(f"Map exported to:\n{file_path}", "Success", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Error exporting:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def export_to_excel(self, file_path):
        """Export data to Excel"""
        import pandas as pd

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                data = self.current_data.data

                if len(data.shape) == 3:
                    # Export sum spectrum
                    spectrum = np.sum(data, axis=(0, 1))
                    energy = self.get_energy_axis()

                    df = pd.DataFrame({
                        'Energy (keV)': [f"{e:.2f}" for e in energy],
                        'Intensity (Counts)': [f"{i:.2f}" for i in spectrum]
                    })
                    df.to_excel(writer, sheet_name='Sum_Spectrum', index=False)

                    # Export sum image stats
                    sum_image = np.sum(data, axis=2)
                    stats_df = pd.DataFrame({
                        'Parameter': ['Width (px)', 'Height (px)', 'Channels', 'Min', 'Max', 'Mean'],
                        'Value': [
                            f"{data.shape[1]:.2f}",
                            f"{data.shape[0]:.2f}",
                            f"{data.shape[2]:.2f}",
                            f"{np.min(sum_image):.2f}",
                            f"{np.max(sum_image):.2f}",
                            f"{np.mean(sum_image):.2f}"
                        ]
                    })
                    stats_df.to_excel(writer, sheet_name='Statistics', index=False)

                elif len(data.shape) == 2:
                    stats_df = pd.DataFrame({
                        'Parameter': ['Width (px)', 'Height (px)', 'Min', 'Max', 'Mean', 'Std'],
                        'Value': [
                            f"{data.shape[1]:.2f}",
                            f"{data.shape[0]:.2f}",
                            f"{np.min(data):.2f}",
                            f"{np.max(data):.2f}",
                            f"{np.mean(data):.2f}",
                            f"{np.std(data):.2f}"
                        ]
                    })
                    stats_df.to_excel(writer, sheet_name='Statistics', index=False)

            wx.MessageBox(f"Data exported to:\n{file_path}", "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error exporting:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def export_to_csv(self, file_path):
        """Export spectrum to CSV"""
        import pandas as pd

        try:
            data = self.current_data.data

            if len(data.shape) == 3:
                spectrum = np.sum(data, axis=(0, 1))
            elif len(data.shape) == 1:
                spectrum = data
            else:
                wx.MessageBox("Cannot export 2D image to CSV spectrum format",
                             "Error", wx.OK | wx.ICON_ERROR)
                return

            energy = self.get_energy_axis()

            df = pd.DataFrame({
                'Energy_keV': [f"{e:.2f}" for e in energy],
                'Intensity_Counts': [f"{i:.2f}" for i in spectrum]
            })
            df.to_csv(file_path, index=False)

            wx.MessageBox(f"Spectrum exported to:\n{file_path}", "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error exporting:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_right_click(self, event):
        """Show right-click context menu"""
        menu = wx.Menu()

        # Save EDX Plot option at top
        save_plot_item = menu.Append(wx.ID_ANY, "Save Current EDX Plot...")
        self.Bind(wx.EVT_MENU, self.on_save_edx_plot, save_plot_item)

        menu.AppendSeparator()

        # HeatMap submenu
        heatmap_menu = wx.Menu()
        colormaps = ['plasma', 'viridis', 'inferno', 'magma', 'hot', 'cool', 'gray', 'jet',
                     'rainbow', 'turbo', 'cividis', 'Spectral', 'coolwarm', 'RdYlBu', 'RdBu']

        for cmap in colormaps:
            item = heatmap_menu.AppendRadioItem(wx.ID_ANY, cmap)
            if cmap == self.current_cmap:
                item.Check(True)
            self.Bind(wx.EVT_MENU, lambda evt, c=cmap: self.on_change_colormap(c), item)

        menu.AppendSubMenu(heatmap_menu, "HeatMap")

        menu.AppendSeparator()

        data_browser_item = menu.Append(wx.ID_ANY, "Data Browser...")
        info_item = menu.Append(wx.ID_ANY, "Info...")
        menu.AppendSeparator()
        clear_item = menu.Append(wx.ID_ANY, "Clear Selections")

        self.Bind(wx.EVT_MENU, self.on_show_data_browser, data_browser_item)
        self.Bind(wx.EVT_MENU, self.on_show_info, info_item)
        self.Bind(wx.EVT_MENU, self.on_clear_selections, clear_item)

        self.PopupMenu(menu)
        menu.Destroy()

    def on_quantification(self, event):
        """Perform EDX quantification and display atomic percentages"""
        if self.current_data is None:
            wx.MessageBox("No EDX data loaded.", "Quantification Error", wx.OK | wx.ICON_WARNING)
            return

        if not self.selected_elements:
            wx.MessageBox("Please select elements first using the periodic table.",
                          "Quantification Error", wx.OK | wx.ICON_WARNING)
            return

        try:
            # Get sum spectrum
            data = self.current_data.data
            if len(data.shape) == 3:
                spectrum = np.sum(data, axis=(0, 1))
            elif len(data.shape) == 1:
                spectrum = data
            else:
                wx.MessageBox("Unsupported data shape for quantification.",
                              "Quantification Error", wx.OK | wx.ICON_WARNING)
                return

            energy = self.get_energy_axis()
            if energy is None:
                wx.MessageBox("Could not get energy axis.",
                              "Quantification Error", wx.OK | wx.ICON_WARNING)
                return

            # Calculate atomic percentages using peak intensities
            results = self._calculate_atomic_percent(energy, spectrum, self.selected_elements)

            if results:
                # Display results in a dialog
                self._show_quantification_results(results)

                # Add to peak fitting grid in parent window if available
                if self.parent is not None and hasattr(self.parent, 'peak_params_grid'):
                    self._add_results_to_grid(results)

        except Exception as e:
            wx.MessageBox(f"Quantification error: {str(e)}",
                          "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def _calculate_atomic_percent(self, energy, spectrum, elements):
        """Calculate atomic percentages from peak intensities"""
        from exspy.material import elements as exspy_elements

        results = []
        total_intensity = 0
        element_intensities = {}

        for element in elements:
            try:
                elem_obj = exspy_elements[element]
                if hasattr(elem_obj, 'Atomic_properties') and hasattr(elem_obj.Atomic_properties, 'Xray_lines'):
                    xray_lines = elem_obj.Atomic_properties.Xray_lines

                    # Find the strongest line (Ka preferred, then La)
                    best_line = None
                    best_energy = None
                    for line_type in ['Ka', 'La', 'Ma']:
                        if line_type in xray_lines:
                            best_line = line_type
                            best_energy = xray_lines[line_type].energy_keV
                            break

                    if best_energy is not None:
                        # Find peak intensity at this energy
                        idx = np.argmin(np.abs(energy - best_energy))

                        # Integrate around peak (simple approach)
                        window = 5  # channels
                        start_idx = max(0, idx - window)
                        end_idx = min(len(spectrum), idx + window)
                        peak_intensity = np.sum(spectrum[start_idx:end_idx])

                        # Get atomic weight for normalization
                        atomic_weight = elem_obj.General_properties.atomic_weight

                        element_intensities[element] = {
                            'intensity': peak_intensity,
                            'line': best_line,
                            'energy': best_energy,
                            'atomic_weight': atomic_weight
                        }
                        total_intensity += peak_intensity

            except (KeyError, AttributeError) as e:
                print(f"Could not process element {element}: {e}")
                continue

        # Calculate atomic percentages (simplified - without k-factors)
        if total_intensity > 0:
            for element, data in element_intensities.items():
                # Simple normalized intensity (not true atomic %)
                # For accurate results, k-factors would be needed
                normalized = (data['intensity'] / total_intensity) * 100
                results.append({
                    'element': element,
                    'line': data['line'],
                    'energy': data['energy'],
                    'intensity': data['intensity'],
                    'atomic_percent': normalized,
                    'atomic_weight': data['atomic_weight']
                })

        return results

    def _show_quantification_results(self, results):
        """Display quantification results in a dialog"""
        dlg = wx.Dialog(self, title="EDX Quantification Results", size=(400, 300))
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Header
        header = wx.StaticText(panel, label="Element Composition (Relative %)")
        header.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(header, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Note about simplified calculation
        note = wx.StaticText(panel, label="Note: Simplified calculation without k-factors.\nFor accurate results, use instrument-specific k-factors.")
        note.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(note, 0, wx.ALL, 5)

        # Results list
        list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_HRULES)
        list_ctrl.InsertColumn(0, "Element", width=80)
        list_ctrl.InsertColumn(1, "Line", width=60)
        list_ctrl.InsertColumn(2, "Energy (keV)", width=100)
        list_ctrl.InsertColumn(3, "Rel. At.%", width=80)

        for i, res in enumerate(results):
            list_ctrl.InsertItem(i, res['element'])
            list_ctrl.SetItem(i, 1, res['line'])
            list_ctrl.SetItem(i, 2, f"{res['energy']:.2f}")
            list_ctrl.SetItem(i, 3, f"{res['atomic_percent']:.2f}")

        sizer.Add(list_ctrl, 1, wx.EXPAND | wx.ALL, 10)

        # Close button
        close_btn = wx.Button(panel, wx.ID_OK, "Close")
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        panel.SetSizer(sizer)
        dlg.ShowModal()
        dlg.Destroy()

    def _add_results_to_grid(self, results):
        """Add quantification results to the peak fitting grid"""
        if self.parent is None or not hasattr(self.parent, 'peak_params_grid'):
            return

        grid = self.parent.peak_params_grid

        # Clear existing grid
        if grid.GetNumberRows() > 0:
            grid.DeleteRows(0, grid.GetNumberRows())

        # Add results
        for i, res in enumerate(results):
            grid.AppendRows(2)  # Data row + constraint row
            row = i * 2

            # Format line name with Greek letters
            line_type = res['line']
            line_type = line_type.replace('Ka', 'Kα').replace('Kb', 'Kβ')
            line_type = line_type.replace('La', 'Lα').replace('Lb', 'Lβ')
            line_type = line_type.replace('Ma', 'Mα').replace('Mb', 'Mβ')

            # Peak label
            label = f"{res['element']} {line_type}"
            grid.SetCellValue(row, 1, label)

            # Position (energy)
            grid.SetCellValue(row, 2, f"{res['energy']:.2f}")

            # Intensity
            grid.SetCellValue(row, 3, f"{res['intensity']:.2f}")

            # Atomic percent in a relevant column (using column 6 for Area or similar)
            grid.SetCellValue(row, 6, f"{res['atomic_percent']:.2f}")

        self.parent.canvas.draw()

    def _populate_edx_grid(self, energy, spectrum, title="EDX"):
        """Calculate quantification and populate peak fitting grid"""
        if self.parent is None or not hasattr(self.parent, 'peak_params_grid'):
            return

        # Always detect elements fresh from the current spectrum
        elements = self._detect_elements_from_spectrum(energy, spectrum)

        # If no elements detected, try using selected elements
        if not elements and self.selected_elements:
            elements = self.selected_elements

        if not elements:
            # Clear grid if no elements
            grid = self.parent.peak_params_grid
            if grid.GetNumberRows() > 0:
                grid.DeleteRows(0, grid.GetNumberRows())
            grid.ForceRefresh()
            return

        # Calculate quantification results
        results = self._calculate_quantification(energy, spectrum, elements)

        if not results:
            # Clear grid if no results
            grid = self.parent.peak_params_grid
            if grid.GetNumberRows() > 0:
                grid.DeleteRows(0, grid.GetNumberRows())
            grid.ForceRefresh()
            return

        # Populate grid
        self._update_peak_params_grid(results, title)

    def _detect_elements_from_spectrum(self, energy, spectrum):
        """Auto-detect elements from spectrum peaks"""
        try:
            from exspy.utils.eds import get_xray_lines_near_energy
            from scipy.signal import find_peaks

            detected_elements = []
            detected_energies = {}

            # Find peaks in spectrum
            threshold = np.max(spectrum) * 0.03  # 3% threshold
            peaks, properties = find_peaks(spectrum, height=threshold, distance=10, prominence=threshold * 0.5)

            for peak_idx in peaks:
                if peak_idx >= len(energy):
                    continue
                peak_energy = energy[peak_idx]
                peak_height = spectrum[peak_idx]

                try:
                    lines_list = get_xray_lines_near_energy(peak_energy, width=0.12)
                    if lines_list:
                        # Get the first (best) match
                        line_str = lines_list[0]
                        parts = line_str.split('_')
                        if len(parts) == 2:
                            element = parts[0]
                            # Only add if not already detected or if this peak is stronger
                            if element not in detected_elements:
                                detected_elements.append(element)
                                detected_energies[element] = (peak_energy, peak_height)
                            elif peak_height > detected_energies[element][1]:
                                detected_energies[element] = (peak_energy, peak_height)
                except Exception:
                    continue

            # Limit to top 15 elements by peak height
            if len(detected_elements) > 15:
                sorted_elements = sorted(detected_elements,
                                         key=lambda e: detected_energies.get(e, (0, 0))[1],
                                         reverse=True)
                detected_elements = sorted_elements[:15]

            return detected_elements

        except Exception as e:
            print(f"Element detection error: {e}")
            return []

    def _calculate_quantification(self, energy, spectrum, elements):
        """Calculate quantification with peak fitting parameters"""
        try:
            from exspy.material import elements as exspy_elements

            results = []
            total_intensity = 0

            for element in elements:
                try:
                    elem_obj = exspy_elements[element]
                    if not hasattr(elem_obj, 'Atomic_properties'):
                        continue
                    if not hasattr(elem_obj.Atomic_properties, 'Xray_lines'):
                        continue

                    xray_lines = elem_obj.Atomic_properties.Xray_lines

                    # Find the strongest line (Ka preferred, then La, Ma)
                    best_line = None
                    best_energy_val = None
                    for line_type in ['Ka', 'La', 'Ma']:
                        if line_type in xray_lines:
                            best_line = line_type
                            best_energy_val = xray_lines[line_type].energy_keV
                            break

                    if best_energy_val is None:
                        continue

                    # Find peak in spectrum near this energy
                    idx = np.argmin(np.abs(energy - best_energy_val))

                    # Define integration window (±0.15 keV typical for EDX)
                    energy_window = 0.15  # keV
                    mask = (energy >= best_energy_val - energy_window) & (energy <= best_energy_val + energy_window)

                    if not np.any(mask):
                        continue

                    # Extract peak region
                    peak_energy = energy[mask]
                    peak_spectrum = spectrum[mask]

                    if len(peak_spectrum) == 0:
                        continue

                    # Calculate peak parameters
                    peak_height = float(np.max(peak_spectrum))

                    # Estimate background as minimum in the window
                    background = float(np.min(peak_spectrum))

                    # Net peak height
                    net_height = peak_height - background

                    # Estimate FWHM from the peak shape
                    half_max = background + net_height / 2
                    above_half = peak_spectrum >= half_max
                    if np.any(above_half):
                        indices = np.where(above_half)[0]
                        if len(indices) > 1:
                            fwhm_kev = peak_energy[indices[-1]] - peak_energy[indices[0]]
                        else:
                            fwhm_kev = 0.1  # Default 100 eV
                    else:
                        fwhm_kev = 0.1

                    # Calculate area using Gaussian approximation: Area ≈ Height × FWHM × 1.064
                    # This is more accurate than trapezoid integration for peaks
                    peak_area = net_height * (fwhm_kev * 1000) * 1.064  # Convert FWHM to eV for area calc

                    # Get atomic weight
                    atomic_weight = elem_obj.General_properties.atomic_weight

                    results.append({
                        'element': element,
                        'line': best_line,
                        'energy': best_energy_val,
                        'height': net_height,
                        'area': peak_area,
                        'fwhm': fwhm_kev * 1000,  # Store in eV
                        'atomic_weight': atomic_weight
                    })
                    total_intensity += peak_area

                except (KeyError, AttributeError) as e:
                    print(f"Could not process element {element}: {e}")
                    continue

            # Calculate atomic percentages
            if total_intensity > 0:
                for res in results:
                    res['concentration'] = (res['area'] / total_intensity) * 100

            return results

        except Exception as e:
            print(f"Quantification error: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _update_peak_params_grid(self, results, title="EDX"):
        """Update peak fitting grid with EDX results"""
        if self.parent is None or not hasattr(self.parent, 'peak_params_grid'):
            return

        grid = self.parent.peak_params_grid

        # Clear existing grid completely
        if grid.GetNumberRows() > 0:
            grid.DeleteRows(0, grid.GetNumberRows())

        # Add results
        for i, res in enumerate(results):
            grid.AppendRows(2)  # Data row + constraint row
            row = i * 2
            constraint_row = row + 1

            # Format line name with Greek letters
            line_type = res['line']
            line_display = line_type.replace('Ka', 'Kα').replace('Kb', 'Kβ')
            line_display = line_display.replace('La', 'Lα').replace('Lb', 'Lβ')
            line_display = line_display.replace('Ma', 'Mα').replace('Mb', 'Mβ')

            # Column 0: ID (letter)
            letter_id = chr(65 + i)  # A, B, C, ...
            grid.SetCellValue(row, 0, letter_id)

            # Column 1: Peak label
            label = f"{res['element']} {line_display}"
            grid.SetCellValue(row, 1, label)

            # Column 2: Position (energy in keV displayed as eV * 1000)
            grid.SetCellValue(row, 2, f"{res['energy'] * 1000:.2f}")
            grid.SetCellValue(constraint_row, 2, "fixed")

            # Column 3: Height
            grid.SetCellValue(row, 3, f"{res['height']:.2f}")
            grid.SetCellValue(constraint_row, 3, "fixed")

            # Column 4: FWHM (in eV)
            grid.SetCellValue(row, 4, f"{res['fwhm']:.2f}")
            grid.SetCellValue(constraint_row, 4, "fixed")

            # Column 6: Area
            grid.SetCellValue(row, 6, f"{res['area']:.2f}")
            grid.SetCellValue(constraint_row, 6, "fixed")

            # Column 10: Concentration (%)
            grid.SetCellValue(row, 10, f"{res.get('concentration', 0):.2f}")

            # Apply formatting - Data row (white background, read-only)
            for col in range(grid.GetNumberCols()):
                grid.SetCellBackgroundColour(row, col, wx.WHITE)
                grid.SetReadOnly(row, col, True)

            # Constraint row - light green background, read-only
            for col in range(grid.GetNumberCols()):
                grid.SetCellBackgroundColour(constraint_row, col, wx.Colour(200, 245, 228))
                grid.SetReadOnly(constraint_row, col, True)

        grid.ForceRefresh()

    def on_change_colormap(self, cmap):
        """Change colormap and replot"""
        self.current_cmap = cmap
        self.plot_current_map()

    def on_show_data_browser(self, event):
        """Show data browser window"""
        if not hasattr(self, 'data_browser_window') or self.data_browser_window is None:
            self.data_browser_window = DataBrowserWindow(self)
        self.data_browser_window.Show()
        self.data_browser_window.Raise()

    def on_show_info(self, event):
        """Show info window"""
        if self.current_data is None:
            wx.MessageBox("No data loaded", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        if not hasattr(self, 'info_window') or self.info_window is None:
            self.info_window = InfoWindow(self)
        self.info_window.update_info(self.current_data)
        self.info_window.Show()
        self.info_window.Raise()

    def on_intensity_map(self, event):
        """Show intensity/sum map"""
        if self.current_data is None:
            wx.MessageBox("No data loaded", "No Data", wx.OK | wx.ICON_WARNING)
            return
        self.plot_current_map()

    def on_close(self, event):
        """Clear parent reference when closing"""
        if hasattr(self.parent, 'edx_window'):
            self.parent.edx_window = None
        event.Skip()  # Allow normal close


class DataBrowserWindow(wx.Frame):
    """Separate window for data browser"""

    def __init__(self, parent):
        super().__init__(parent, title="Data Browser", size=(400, 500))
        self.parent = parent

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Data tree
        self.data_tree = wx.TreeCtrl(panel, style=wx.TR_DEFAULT_STYLE | wx.TR_HAS_BUTTONS)
        self.tree_root = self.data_tree.AddRoot("Loaded Data")
        sizer.Add(self.data_tree, 1, wx.EXPAND | wx.ALL, 5)

        # Refresh from parent
        self.refresh_tree()

        panel.SetSizer(sizer)

        self.data_tree.Bind(wx.EVT_TREE_SEL_CHANGED, self.on_tree_select)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def refresh_tree(self):
        """Refresh tree from parent data"""
        self.data_tree.DeleteChildren(self.tree_root)

        if hasattr(self.parent, 'loaded_signals'):
            for item in self.parent.loaded_signals:
                signal = item['data']
                filename = item['filename']

                title = filename
                if hasattr(signal, 'metadata') and hasattr(signal.metadata, 'General'):
                    if hasattr(signal.metadata.General, 'title') and signal.metadata.General.title:
                        title = signal.metadata.General.title

                item_text = f"{title} ({signal.data.shape})"
                tree_item = self.data_tree.AppendItem(self.tree_root, item_text)
                self.data_tree.SetItemData(tree_item, item)

                # Add metadata children
                self.add_metadata_to_tree(tree_item, signal)

        self.data_tree.Expand(self.tree_root)

    def add_metadata_to_tree(self, parent_item, signal):
        """Add signal metadata as tree children"""
        self.data_tree.AppendItem(parent_item, f"Shape: {signal.data.shape}")
        self.data_tree.AppendItem(parent_item, f"Dtype: {signal.data.dtype}")

        if hasattr(signal, 'axes_manager'):
            axes_item = self.data_tree.AppendItem(parent_item, "Axes")
            for i, axis in enumerate(signal.axes_manager.signal_axes):
                axis_text = f"Signal {i}: {axis.name} [{axis.units}] - {axis.size} pts"
                self.data_tree.AppendItem(axes_item, axis_text)
            for i, axis in enumerate(signal.axes_manager.navigation_axes):
                axis_text = f"Nav {i}: {axis.name} [{axis.units}] - {axis.size} pts"
                self.data_tree.AppendItem(axes_item, axis_text)

    def on_tree_select(self, event):
        """Handle tree selection"""
        item = event.GetItem()
        data = self.data_tree.GetItemData(item)

        if data is not None and 'data' in data:
            self.parent.current_data = data['data']
            self.parent.plot_current_map()
            self.parent.extract_elements(data['data'])

    def on_close(self, event):
        """Handle close"""
        self.Hide()


class InfoWindow(wx.Frame):
    """Separate window for data info"""

    def __init__(self, parent):
        super().__init__(parent, title="Data Info", size=(350, 400))
        self.parent = parent

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.info_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        sizer.Add(self.info_text, 1, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(sizer)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def update_info(self, signal):
        """Update info display"""
        info = []
        info.append(f"Shape: {signal.data.shape}")
        info.append(f"Data type: {signal.data.dtype}")
        info.append(f"Min: {np.min(signal.data):.2f}")
        info.append(f"Max: {np.max(signal.data):.2f}")
        info.append(f"Mean: {np.mean(signal.data):.2f}")

        if hasattr(signal, 'axes_manager'):
            info.append("")
            info.append("Axes:")
            for axis in signal.axes_manager.signal_axes:
                info.append(f"  {axis.name}: {axis.size} pts, {axis.scale:.4f} {axis.units}/pt")
            for axis in signal.axes_manager.navigation_axes:
                info.append(f"  {axis.name}: {axis.size} pts, {axis.scale:.4f} {axis.units}/pt")

        self.info_text.SetValue('\n'.join(info))

    def on_close(self, event):
        self.Hide()


class PeriodicTableDialog(wx.Dialog):
    """Periodic table element selector"""

    def __init__(self, parent, selected_elements=None):
        super().__init__(parent, title="Select Elements", size=(470, 230))

        self.selected_elements = selected_elements or []
        self.element_buttons = {}

        panel = wx.Panel(self, style=wx.BORDER_RAISED)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Periodic table layout
        table_panel = wx.Panel(panel)
        grid_sizer = wx.GridBagSizer(0, 0)

        # Periodic table structure (row, col, symbol)
        elements = [
            (0, 0, 'H'), (0, 17, 'He'),
            (1, 0, 'Li'), (1, 1, 'Be'), (1, 12, 'B'), (1, 13, 'C'), (1, 14, 'N'), (1, 15, 'O'), (1, 16, 'F'), (1, 17, 'Ne'),
            (2, 0, 'Na'), (2, 1, 'Mg'), (2, 12, 'Al'), (2, 13, 'Si'), (2, 14, 'P'), (2, 15, 'S'), (2, 16, 'Cl'), (2, 17, 'Ar'),
            (3, 0, 'K'), (3, 1, 'Ca'), (3, 2, 'Sc'), (3, 3, 'Ti'), (3, 4, 'V'), (3, 5, 'Cr'), (3, 6, 'Mn'), (3, 7, 'Fe'),
            (3, 8, 'Co'), (3, 9, 'Ni'), (3, 10, 'Cu'), (3, 11, 'Zn'), (3, 12, 'Ga'), (3, 13, 'Ge'), (3, 14, 'As'), (3, 15, 'Se'), (3, 16, 'Br'), (3, 17, 'Kr'),
            (4, 0, 'Rb'), (4, 1, 'Sr'), (4, 2, 'Y'), (4, 3, 'Zr'), (4, 4, 'Nb'), (4, 5, 'Mo'), (4, 6, 'Tc'), (4, 7, 'Ru'),
            (4, 8, 'Rh'), (4, 9, 'Pd'), (4, 10, 'Ag'), (4, 11, 'Cd'), (4, 12, 'In'), (4, 13, 'Sn'), (4, 14, 'Sb'), (4, 15, 'Te'), (4, 16, 'I'), (4, 17, 'Xe'),
            (5, 0, 'Cs'), (5, 1, 'Ba'), (5, 2, 'La'), (5, 3, 'Hf'), (5, 4, 'Ta'), (5, 5, 'W'), (5, 6, 'Re'), (5, 7, 'Os'),
            (5, 8, 'Ir'), (5, 9, 'Pt'), (5, 10, 'Au'), (5, 11, 'Hg'), (5, 12, 'Tl'), (5, 13, 'Pb'), (5, 14, 'Bi'), (5, 15, 'Po'), (5, 16, 'At'), (5, 17, 'Rn'),
        ]

        for row, col, symbol in elements:
            btn = wx.ToggleButton(table_panel, label=symbol, size=(25, 25))
            btn.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))

            if symbol in self.selected_elements:
                btn.SetValue(True)
                btn.SetBackgroundColour(wx.Colour(100, 200, 100))

            btn.Bind(wx.EVT_TOGGLEBUTTON, lambda evt, s=symbol, b=btn: self.on_element_toggle(evt, s, b))
            grid_sizer.Add(btn, pos=(row, col), flag=wx.ALL, border=0)
            self.element_buttons[symbol] = btn

        table_panel.SetSizer(grid_sizer)
        main_sizer.Add(table_panel, 1, wx.EXPAND | wx.ALL, 0)

        # Buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        clear_btn = wx.Button(panel, label="Clear All")
        clear_btn.Bind(wx.EVT_BUTTON, self.on_clear_all)
        ok_btn = wx.Button(panel, wx.ID_OK, "OK")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "Cancel")

        btn_sizer.Add(clear_btn, 0, wx.ALL, 2)
        btn_sizer.AddStretchSpacer()
        btn_sizer.Add(ok_btn, 0, wx.ALL, 2)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 2)

        main_sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 2)

        panel.SetSizer(main_sizer)
        self.Centre()

    def on_element_toggle(self, event, symbol, btn):
        """Handle element toggle"""
        if btn.GetValue():
            if symbol not in self.selected_elements:
                self.selected_elements.append(symbol)
            btn.SetBackgroundColour(wx.Colour(100, 200, 100))
        else:
            if symbol in self.selected_elements:
                self.selected_elements.remove(symbol)
            btn.SetBackgroundColour(wx.NullColour)
        btn.Refresh()

    def on_clear_all(self, event):
        """Clear all selections"""
        self.selected_elements = []
        for btn in self.element_buttons.values():
            btn.SetValue(False)
            btn.SetBackgroundColour(wx.NullColour)
            btn.Refresh()

    def get_selected_elements(self):
        return self.selected_elements


class EDXSensitivityWindow(wx.Frame):
    """Popup window for EDX sensitivity and display controls"""

    def __init__(self, parent):
        super().__init__(parent, title="EDX Display Controls",
                         size=(300, 250),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)

        self.parent = parent
        self.init_ui()
        self.Centre()

    def init_ui(self):
        """Initialize the control interface"""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Add some padding
        main_sizer.AddSpacer(10)

        # X Max control
        x_max_box = wx.BoxSizer(wx.HORIZONTAL)
        x_max_label = wx.StaticText(panel, label="X Max (keV):")
        x_max_box.Add(x_max_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.x_max_spin = wx.SpinCtrl(panel, value="20", min=1, max=50, size=(80, -1))
        self.x_max_spin.SetToolTip("Set maximum X-axis energy (0 to X)")
        self.x_max_spin.Bind(wx.EVT_SPINCTRL, self.on_x_max_change)
        x_max_box.Add(self.x_max_spin, 0, wx.ALIGN_CENTER_VERTICAL)

        main_sizer.Add(x_max_box, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        # Y Max control
        y_max_box = wx.BoxSizer(wx.HORIZONTAL)
        y_max_label = wx.StaticText(panel, label="Y Max (counts):")
        y_max_box.Add(y_max_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.y_max_spin = wx.SpinCtrlDouble(panel, value="10000", min=100, max=1000000,
                                            inc=1000, size=(100, -1))
        self.y_max_spin.SetToolTip("Set maximum Y-axis intensity")
        self.y_max_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_y_max_change)
        y_max_box.Add(self.y_max_spin, 0, wx.ALIGN_CENTER_VERTICAL)

        main_sizer.Add(y_max_box, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        # Sensitivity slider
        sensitivity_box = wx.BoxSizer(wx.VERTICAL)
        sensitivity_label = wx.StaticText(panel, label="Peak Detection Sensitivity:")
        sensitivity_box.Add(sensitivity_label, 0, wx.BOTTOM, 5)

        slider_box = wx.BoxSizer(wx.HORIZONTAL)
        low_label = wx.StaticText(panel, label="Low")
        slider_box.Add(low_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        self.sensitivity_slider = wx.Slider(panel, value=50, minValue=1, maxValue=100,
                                            style=wx.SL_HORIZONTAL, size=(150, -1))
        self.sensitivity_slider.SetToolTip("Adjust peak detection sensitivity\n(Higher = more peaks detected)")
        self.sensitivity_slider.Bind(wx.EVT_SLIDER, self.on_sensitivity_change)
        slider_box.Add(self.sensitivity_slider, 1, wx.ALIGN_CENTER_VERTICAL)

        high_label = wx.StaticText(panel, label="High")
        slider_box.Add(high_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 5)

        sensitivity_box.Add(slider_box, 0, wx.EXPAND)

        # Show current sensitivity value
        self.sensitivity_value_label = wx.StaticText(panel, label="Value: 50")
        sensitivity_box.Add(self.sensitivity_value_label, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.TOP, 5)

        main_sizer.Add(sensitivity_box, 0, wx.ALL | wx.EXPAND, 10)

        # Buttons
        button_box = wx.BoxSizer(wx.HORIZONTAL)

        reset_btn = wx.Button(panel, label="Reset to Default")
        reset_btn.Bind(wx.EVT_BUTTON, self.on_reset)
        button_box.Add(reset_btn, 0, wx.RIGHT, 5)

        apply_btn = wx.Button(panel, label="Apply")
        apply_btn.Bind(wx.EVT_BUTTON, self.on_apply)
        button_box.Add(apply_btn, 0)

        main_sizer.Add(button_box, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        panel.SetSizer(main_sizer)

        # Initialize values from current plot
        self.initialize_from_plot()

    def initialize_from_plot(self):
        """Set initial values from current plot state"""
        # Set default to 20 keV
        self.x_max_spin.SetValue(20)
        self.y_max_spin.SetValue(10000)

        if self.parent.parent is not None and hasattr(self.parent.parent, 'ax'):
            # Get current X max if plot exists
            try:
                xlim = self.parent.parent.ax.get_xlim()
                if xlim[1] > 0:
                    self.x_max_spin.SetValue(int(xlim[1]))

                # Get current Y max
                ylim = self.parent.parent.ax.get_ylim()
                if ylim[1] > 0:
                    self.y_max_spin.SetValue(ylim[1])
            except:
                pass

    def on_x_max_change(self, event):
        """Handle X max spin control change"""
        if self.parent.parent is not None and hasattr(self.parent.parent, 'ax'):
            x_max = self.x_max_spin.GetValue()
            current_xlim = self.parent.parent.ax.get_xlim()
            self.parent.parent.ax.set_xlim(0, x_max)  # Always start from 0

            # Store the X max preference
            current_sheet = self.parent.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot' and 'Core levels' in self.parent.parent.Data:
                if current_sheet in self.parent.parent.Data['Core levels']:
                    self.parent.parent.Data['Core levels'][current_sheet]['_EDX_display_max'] = x_max

            self.parent.parent.canvas.draw()

    def on_y_max_change(self, event):
        """Handle Y max spin control change"""
        if self.parent.parent is not None and hasattr(self.parent.parent, 'ax'):
            y_max = self.y_max_spin.GetValue()
            self.parent.parent.ax.set_ylim(0, y_max)
            self.parent.parent.canvas.draw()

    def on_sensitivity_change(self, event):
        """Handle sensitivity slider change"""
        sensitivity = self.sensitivity_slider.GetValue()
        self.sensitivity_value_label.SetLabel(f"Value: {sensitivity}")

    def on_apply(self, event):
        """Apply sensitivity and re-plot"""
        if self.parent.parent is not None and hasattr(self.parent.parent, 'sheet_combobox'):
            sensitivity = self.sensitivity_slider.GetValue()
            # Sensitivity: 1-100, where 100 = most sensitive (1% threshold)
            # Convert to threshold: 100 -> 0.01, 1 -> 0.10
            threshold = 0.11 - (sensitivity / 1000.0)

            print(f"Applying sensitivity: {sensitivity}, threshold: {threshold:.4f}")

            # Store threshold in parent EDX window
            self.parent.peak_threshold = threshold

            # Clear current plot
            self.parent.parent.ax.clear()

            # Re-plot with new sensitivity
            current_sheet = self.parent.parent.sheet_combobox.GetValue()
            if current_sheet == 'EDX~Plot':
                # Make sure the threshold is used
                if 'Core levels' in self.parent.parent.Data and current_sheet in self.parent.parent.Data['Core levels']:
                    sheet_data = self.parent.parent.Data['Core levels'][current_sheet]
                    sheet_data['_EDX_sensitivity'] = threshold

                window.plot_manager.plot_edx_data(self.parent.parent, current_sheet)

    def on_reset(self, event):
        """Reset all values to defaults"""
        self.x_max_spin.SetValue(20)
        self.y_max_spin.SetValue(10000)
        self.sensitivity_slider.SetValue(50)
        self.sensitivity_value_label.SetLabel("Value: 50")

        # Apply defaults
        if self.parent.parent is not None and hasattr(self.parent.parent, 'ax'):
            self.parent.parent.ax.set_xlim(0, 20)
            self.parent.parent.ax.set_ylim(0, 10000)
            self.parent.parent.canvas.draw()


class RotatableRectangle:
    """Custom rotatable rectangle selector for area selection"""

    def __init__(self, ax, callback):
        self.ax = ax
        self.callback = callback
        self.active = False
        self.angle = 0

        self.center = None
        self.width = None
        self.height = None

        self.rectangle = None
        self.rotation_handle = None
        self.resize_handles = []

        self.dragging = False
        self.rotating = False
        self.resizing = False
        self.drag_start = None
        self.resize_corner = None
        self.angle_text = None

        self._stored_xlim = None
        self._stored_ylim = None

        self.cid_press = ax.figure.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = ax.figure.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = ax.figure.canvas.mpl_connect('motion_notify_event', self.on_motion)
        self.cid_scroll = ax.figure.canvas.mpl_connect('scroll_event', self.on_scroll)

    def start_selection(self, event):
        """Start creating a new rectangle"""
        if event.inaxes != self.ax or event.button != 1:
            return

        # If not active, create new rectangle at click position
        if not self.active:
            # Store axis limits BEFORE any drawing
            self._stored_xlim = self.ax.get_xlim()
            self._stored_ylim = self.ax.get_ylim()

            self.center = (event.xdata, event.ydata)
            self.width = 10
            self.height = 10
            self.angle = 0
            self.active = True
            self.draw_rectangle()

    def draw_rectangle(self):
        """Draw the rotatable rectangle and handles"""
        # Clear existing elements safely
        if self.rectangle:
            try:
                self.rectangle.remove()
            except (ValueError, AttributeError):
                pass
            self.rectangle = None

        for handle in self.resize_handles:
            try:
                handle.remove()
            except (ValueError, AttributeError):
                pass
        self.resize_handles = []

        if self.rotation_handle:
            try:
                self.rotation_handle.remove()
            except (ValueError, AttributeError):
                pass
            self.rotation_handle = None

        if self.angle_text:
            try:
                self.angle_text.remove()
            except (ValueError, AttributeError):
                pass
            self.angle_text = None

        if self.center is None:
            self.ax.figure.canvas.draw_idle()
            return

        from matplotlib.patches import Rectangle
        from matplotlib.transforms import Affine2D

        rect = Rectangle((-self.width / 2, -self.height / 2), self.width, self.height,
                         linewidth=2, edgecolor='lime', facecolor='green', alpha=0.3)

        t = Affine2D().rotate_deg(self.angle).translate(*self.center) + self.ax.transData
        rect.set_transform(t)
        rect._is_selection_marker = True

        self.rectangle = self.ax.add_patch(rect)

        # Rotation handle
        handle_dist = max(self.width, self.height) / 2 + 10
        angle_rad = np.radians(self.angle)
        handle_x = self.center[0] + handle_dist * np.sin(angle_rad)
        handle_y = self.center[1] + handle_dist * np.cos(angle_rad)

        self.rotation_handle = self.ax.plot(handle_x, handle_y, 'go', markersize=10,
                                            markeredgecolor='darkgreen', markeredgewidth=2)[0]
        self.rotation_handle._is_selection_marker = True

        # Show angle text when rotating
        if self.rotating:
            self.angle_text = self.ax.text(handle_x, handle_y + 5, f'{self.angle:.1f}°',
                                           fontsize=9, color='white', fontweight='bold',
                                           ha='center', va='bottom',
                                           bbox=dict(boxstyle='round,pad=0.2', facecolor='green', alpha=0.7))
            self.angle_text._is_selection_marker = True

        # Resize handles at corners
        corners = self.get_corners()
        for corner in corners:
            handle = self.ax.plot(corner[0], corner[1], 'gs', markersize=8,
                                  markeredgecolor='darkgreen', markeredgewidth=1)[0]
            handle._is_selection_marker = True
            self.resize_handles.append(handle)

        # ALWAYS restore stored axis limits to prevent expansion
        if self._stored_xlim is not None and self._stored_ylim is not None:
            self.ax.set_xlim(self._stored_xlim)
            self.ax.set_ylim(self._stored_ylim)

        self.ax.figure.canvas.draw_idle()

    def get_corners(self):
        """Get the four corners of the rotated rectangle"""
        angle_rad = np.radians(self.angle)
        cos_a = np.cos(angle_rad)
        sin_a = np.sin(angle_rad)

        hw = self.width / 2
        hh = self.height / 2
        corners_local = [(-hw, -hh), (hw, -hh), (hw, hh), (-hw, hh)]

        corners = []
        for x, y in corners_local:
            rx = x * cos_a - y * sin_a + self.center[0]
            ry = x * sin_a + y * cos_a + self.center[1]
            corners.append((rx, ry))

        return corners

    def on_scroll(self, event):
        """Handle mouse scroll to rotate rectangle"""
        if not self.active or event.inaxes != self.ax:
            return

        # Check if mouse is near the rectangle
        if self.center is None:
            return

        # Rotate by 5 degrees per scroll step
        if event.button == 'up':
            self.angle += 5
        elif event.button == 'down':
            self.angle -= 5

        # Normalize angle to -180 to 180
        self.angle = ((self.angle + 180) % 360) - 180

        self.draw_rectangle()

        # Trigger callback
        if self.callback:
            self.callback(self.center, self.width, self.height, self.angle)

    def on_press(self, event):
        """Handle mouse press"""
        if not self.active or event.inaxes != self.ax or event.button != 1:
            return

        # Store current limits
        if self._stored_xlim is None:
            self._stored_xlim = self.ax.get_xlim()
            self._stored_ylim = self.ax.get_ylim()

        # Check rotation handle first (highest priority)
        if self.rotation_handle:
            handle_data = self.rotation_handle.get_data()
            dist = np.sqrt((event.xdata - handle_data[0][0]) ** 2 + (event.ydata - handle_data[1][0]) ** 2)
            if dist < 8:
                self.rotating = True
                self.drag_start = (event.xdata, event.ydata)
                return

        # Check resize handles - but only if very close (within 5 pixels)
        closest_handle_dist = float('inf')
        closest_handle_idx = None
        for i, handle in enumerate(self.resize_handles):
            handle_data = handle.get_data()
            dist = np.sqrt((event.xdata - handle_data[0][0]) ** 2 + (event.ydata - handle_data[1][0]) ** 2)
            if dist < closest_handle_dist:
                closest_handle_dist = dist
                closest_handle_idx = i

        # Only resize if very close to handle (< 5 pixels)
        if closest_handle_dist < 5:
            self.resizing = True
            self.resize_corner = closest_handle_idx
            self.drag_start = (event.xdata, event.ydata)
            return

        # Check if inside rectangle for dragging (most common operation)
        if self.point_in_rectangle(event.xdata, event.ydata):
            self.dragging = True
            self.drag_start = (event.xdata, event.ydata)
            return

        # If not inside but close to resize handle (< 10 pixels), allow resize
        if closest_handle_dist < 10:
            self.resizing = True
            self.resize_corner = closest_handle_idx
            self.drag_start = (event.xdata, event.ydata)

    def on_release(self, event):
        """Handle mouse release"""
        was_interacting = self.dragging or self.resizing or self.rotating

        self.dragging = False
        self.rotating = False
        self.resizing = False
        self.drag_start = None

        # Restore limits after any operation
        if self._stored_xlim is not None and self._stored_ylim is not None:
            self.ax.set_xlim(self._stored_xlim)
            self.ax.set_ylim(self._stored_ylim)

        # Trigger callback after interaction
        if was_interacting and self.callback and self.center:
            self.callback(self.center, self.width, self.height, self.angle)

    def on_motion(self, event):
        """Handle mouse motion"""
        if event.inaxes != self.ax or self.drag_start is None:
            return

        if event.xdata is None or event.ydata is None:
            return

        dx = event.xdata - self.drag_start[0]
        dy = event.ydata - self.drag_start[1]

        if self.dragging:
            new_x = self.center[0] + dx
            new_y = self.center[1] + dy

            # Clamp to stored bounds
            if self._stored_xlim is not None:
                new_x = max(self._stored_xlim[0], min(self._stored_xlim[1], new_x))
            if self._stored_ylim is not None:
                y_min = min(self._stored_ylim[0], self._stored_ylim[1])
                y_max = max(self._stored_ylim[0], self._stored_ylim[1])
                new_y = max(y_min, min(y_max, new_y))

            self.center = (new_x, new_y)
            self.drag_start = (event.xdata, event.ydata)
            self.draw_rectangle()

        elif self.rotating:
            angle_to_center = np.arctan2(event.xdata - self.center[0], event.ydata - self.center[1])
            self.angle = np.degrees(angle_to_center)
            self.draw_rectangle()

        elif self.resizing:
            angle_rad = np.radians(-self.angle)
            cos_a = np.cos(angle_rad)
            sin_a = np.sin(angle_rad)

            local_x = (event.xdata - self.center[0]) * cos_a - (event.ydata - self.center[1]) * sin_a
            local_y = (event.xdata - self.center[0]) * sin_a + (event.ydata - self.center[1]) * cos_a

            # Allow minimum size of 1 pixel
            self.width = max(1, abs(local_x) * 2)
            self.height = max(1, abs(local_y) * 2)
            self.draw_rectangle()

    def point_in_rectangle(self, x, y):
        """Check if point is inside the rotated rectangle"""
        if self.center is None:
            return False

        angle_rad = np.radians(-self.angle)
        cos_a = np.cos(angle_rad)
        sin_a = np.sin(angle_rad)

        local_x = (x - self.center[0]) * cos_a - (y - self.center[1]) * sin_a
        local_y = (x - self.center[0]) * sin_a + (y - self.center[1]) * cos_a

        return abs(local_x) <= self.width / 2 and abs(local_y) <= self.height / 2

    def clear(self):
        """Clear the rectangle and handles, allow new selection"""
        if self.rectangle:
            try:
                self.rectangle.remove()
            except (ValueError, AttributeError):
                pass
            self.rectangle = None

        for handle in self.resize_handles:
            try:
                handle.remove()
            except (ValueError, AttributeError):
                pass
        self.resize_handles = []

        if self.rotation_handle:
            try:
                self.rotation_handle.remove()
            except (ValueError, AttributeError):
                pass
            self.rotation_handle = None

        if self.angle_text:
            try:
                self.angle_text.remove()
            except (ValueError, AttributeError):
                pass
            self.angle_text = None

        # Reset state to allow new selection
        self.center = None
        self.width = None
        self.height = None
        self.angle = 0
        self.active = False
        self.dragging = False
        self.rotating = False
        self.resizing = False
        self.drag_start = None

        self.ax.figure.canvas.draw_idle()

    def disconnect(self):
        """Disconnect event handlers"""
        try:
            self.ax.figure.canvas.mpl_disconnect(self.cid_press)
            self.ax.figure.canvas.mpl_disconnect(self.cid_release)
            self.ax.figure.canvas.mpl_disconnect(self.cid_motion)
            self.ax.figure.canvas.mpl_disconnect(self.cid_scroll)
        except (ValueError, AttributeError):
            pass

def open_edx_sem_window(parent):
    """Open EDX/SEM analysis window"""
    window = EDXSEMWindow(parent)
    window.Show()
    return window