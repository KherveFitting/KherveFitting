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

    def __init__(self, parent, title="EDX/SEM Analysis"):
        super().__init__(parent, title=title, size=(500, 600), style = wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)



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

        self.loaded_signals = []
        self.data_browser_window = None
        self.info_window = None
        self.current_colorbar = None

        self.zoom_selector = None

        self.init_ui()
        self.Centre()


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

        # Set Elements button
        self.set_elements_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        id_path = os.path.join(icon_path, "ID-3.png")
        if os.path.exists(id_path):
            self.set_elements_btn.SetBitmap(wx.Bitmap(id_path, wx.BITMAP_TYPE_PNG))
        else:
            self.set_elements_btn.SetBitmap(self.create_icon_bitmap('point'))
        self.set_elements_btn.SetToolTip("Set Elements")
        toolbar_sizer.Add(self.set_elements_btn, 0, wx.ALL, 2)

        # Plot Maps button
        self.plot_maps_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        heatmap_path = os.path.join(icon_path, "heatmap-3.png")
        if os.path.exists(heatmap_path):
            self.plot_maps_btn.SetBitmap(wx.Bitmap(heatmap_path, wx.BITMAP_TYPE_PNG))
        else:
            self.plot_maps_btn.SetBitmap(self.create_icon_bitmap('area'))

        # Intensity Map button
        self.intensity_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        self.intensity_btn.SetBitmap(self.create_icon_bitmap('intensity'))
        self.intensity_btn.SetToolTip("Show Intensity Map")
        toolbar_sizer.Add(self.intensity_btn, 0, wx.ALL, 2)

        self.plot_maps_btn.SetToolTip("Plot Element Maps")
        toolbar_sizer.Add(self.plot_maps_btn, 0, wx.ALL, 2)

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
            dc.SetBrush(wx.Brush(wx.Colour(255, 0, 0)))
            dc.DrawCircle(10, 10, 4)
        elif icon_type == 'area':
            dc.SetBrush(wx.TRANSPARENT_BRUSH)
            dc.DrawRectangle(3, 3, 14, 14)
        elif icon_type == 'line':
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

        dc.SelectObject(wx.NullBitmap)
        return bmp



    # ==================== TOOLBAR HANDLERS ====================

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

    def on_zoom_out(self, event):
        """Zoom out / reset view"""
        self.map_ax.autoscale()
        self.map_canvas.draw()

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
        """Toggle area selection mode"""
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

            # Create rectangle selector
            if self.rect_selector is None:
                self.rect_selector = RectangleSelector(
                    self.map_ax,
                    self.on_area_select,
                    useblit=True,
                    props=dict(facecolor='red', edgecolor='white', alpha=0.3, fill=True),
                    button=[1],
                    minspanx=5,
                    minspany=5,
                    spancoords='pixels',
                    interactive=True
                )
            else:
                self.rect_selector.set_active(True)
        else:
            self.selection_mode = None
            if self.rect_selector:
                self.rect_selector.set_active(False)

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

        # Clear selection markers
        self.clear_selection_markers()

        self.map_canvas.draw()

    def on_colormap_change(self, event):
        """Change colormap"""
        if self.current_data is not None:
            self.plot_current_map()

    # ==================== SELECTION HANDLERS ====================

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
            # Clear previous point markers
            self.clear_selection_markers()
            self.selected_points = [(x, y)]

            # Draw marker
            marker, = self.map_ax.plot(x, y, 'r+', markersize=15, markeredgewidth=2)
            marker._is_selection_marker = True
            self.map_canvas.draw()

            # Extract and plot spectrum at point
            self.plot_point_spectrum(x, y)

        elif self.selection_mode == 'line':
            if self.line_start is None:
                # First click - start of line
                self.line_start = (x, y)
                # Draw start marker
                marker, = self.map_ax.plot(x, y, 'wo', markersize=8, markeredgecolor='black', markeredgewidth=2)
                marker._is_selection_marker = True
                self.map_canvas.draw()
            else:
                # Second click - end of line
                self.line_end = (x, y)

                # Clear previous markers
                self.clear_selection_markers()

                # Draw line
                line, = self.map_ax.plot([self.line_start[0], self.line_end[0]],
                                         [self.line_start[1], self.line_end[1]],
                                         'w-', linewidth=2, markeredgecolor='black')
                line._is_selection_marker = True

                # Draw end points
                marker1, = self.map_ax.plot(self.line_start[0], self.line_start[1], 'wo',
                                            markersize=8, markeredgecolor='black', markeredgewidth=2)
                marker1._is_selection_marker = True
                marker2, = self.map_ax.plot(self.line_end[0], self.line_end[1], 'wo',
                                            markersize=8, markeredgecolor='black', markeredgewidth=2)
                marker2._is_selection_marker = True

                self.map_canvas.draw()

                # Plot line spectrum
                self.plot_line_spectrum(self.line_start[0], self.line_start[1],
                                        self.line_end[0], self.line_end[1])

                # Reset for next line
                self.line_start = None
                self.line_end = None

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
        """Plot spectrum at a single point in KherveFitting main window"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D (x, y, energy) for point spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        # Extract spectrum at point (y, x order for numpy)
        spectrum = data[y, x, :]
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self.parent.ax.clear()
            self.parent.ax.plot(energy, spectrum, 'b-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'EDX Spectrum at Point ({x}, {y})')
            self.parent.ax.grid(True, alpha=0.3)
            self.parent.canvas.draw()

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
            self.parent.ax.plot(energy, spectrum, 'b-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'Summed EDX Spectrum - Area: {(x2 - x1 + 1) * (y2 - y1 + 1)} pixels')
            self.parent.ax.grid(True, alpha=0.3)
            self.parent.canvas.draw()

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
            self.parent.ax.plot(energy, spectrum, 'b-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title(f'EDX Line Spectrum ({x1},{y1}) to ({x2},{y2}) - {num_points} pts')
            self.parent.ax.grid(True, alpha=0.3)
            self.parent.canvas.draw()

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
        """Import EDX map/spectrum image"""
        with wx.FileDialog(self, "Open EDX Map/Spectrum Image",
                          wildcard="All files (*.*)|*.*|BCF files (*.bcf)|*.bcf|"
                                  "HDF5 files (*.hdf5;*.h5)|*.hdf5;*.h5|"
                                  "EMSA files (*.emsa)|*.emsa|Raw files (*.raw)|*.raw",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            file_path = fileDialog.GetPath()

        self.load_file(file_path, 'edx_map')

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
        """Generic file loader with reader selection"""
        try:
            # Handle HDF5 files with multiple readers
            if file_path.lower().endswith(('.hdf5', '.h5')):
                reader_dialog = wx.SingleChoiceDialog(
                    self,
                    "Multiple readers available for HDF5 file.\nSelect the appropriate reader:",
                    "Select Reader",
                    ['HSPY', 'USID', 'Delmic']
                )
                if reader_dialog.ShowModal() == wx.ID_OK:
                    reader = reader_dialog.GetStringSelection()
                    loaded_data = hs.load(file_path, reader=reader)
                else:
                    reader_dialog.Destroy()
                    return
                reader_dialog.Destroy()
            else:
                # Try to load normally
                try:
                    loaded_data = hs.load(file_path)
                except ValueError as e:
                    if "multiple file readers" in str(e).lower():
                        # Extract available readers from error message
                        reader_dialog = wx.SingleChoiceDialog(
                            self,
                            "Multiple readers available.\nSelect the appropriate reader:",
                            "Select Reader",
                            ['HSPY', 'USID', 'Delmic', 'EMSA', 'Bruker']
                        )
                        if reader_dialog.ShowModal() == wx.ID_OK:
                            reader = reader_dialog.GetStringSelection()
                            loaded_data = hs.load(file_path, reader=reader)
                        else:
                            reader_dialog.Destroy()
                            return
                        reader_dialog.Destroy()
                    else:
                        raise

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
        self.map_ax.set_xlabel('X (pixels)')
        self.map_ax.set_ylabel('Y (pixels)')
        self.map_ax.set_title(title)

        # Add colorbar for sum image
        self.current_colorbar = self.map_figure.colorbar(im, ax=self.map_ax, fraction=0.046, pad=0.04)

        self.map_figure.tight_layout(pad=0.5)
        self.map_canvas.draw()

        # Plot sum spectrum in parent KherveFitting window
        self.plot_sum_spectrum_to_parent()

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
        self.map_ax.set_xlabel('X (pixels)')
        self.map_ax.set_ylabel('Y (pixels)')

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
        """Plot sum spectrum in KherveFitting main window"""
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
            self.parent.ax.plot(energy, spectrum, 'b-', linewidth=0.8)
            self.parent.ax.set_xlabel('Energy (keV)')
            self.parent.ax.set_ylabel('Counts')
            self.parent.ax.set_title('EDX Sum Spectrum')
            self.parent.ax.grid(True, alpha=0.3)
            self.parent.canvas.draw()

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

def open_edx_sem_window(parent):
    """Open EDX/SEM analysis window"""
    window = EDXSEMWindow(parent)
    window.Show()
    return window