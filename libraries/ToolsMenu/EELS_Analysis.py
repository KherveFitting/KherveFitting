"""
EELS_Analysis.py
Module for Electron Energy Loss Spectroscopy (EELS) data analysis
Uses HyperSpy library for .dm3 file import and analysis
Follows same structure as EDX_SEM_Analysis.py but without peak labelling/quantification
"""

import wx
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.widgets import RectangleSelector
import matplotlib.patches as patches
import matplotlib.patheffects as path_effects


class EELSWindow(wx.Frame):
    """Main window for EELS data analysis"""

    def __init__(self, parent, title="EELS HeatMap"):
        super().__init__(parent, title=title, size=(590, 700),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)

        self.parent = parent
        self.eels_data = None
        self.current_data = None

        # Selection modes
        self.selection_mode = None  # 'point', 'area', 'line'
        self.rect_selector = None
        self.selected_points = []
        self.selected_areas = []
        self.line_start = None
        self.line_end = None

        self.point_size = 1  # Size in pixels (1 = single pixel)

        self.loaded_signals = []
        self.current_colorbar = None
        self.zoom_selector = None

        self.init_ui()
        self.Centre()

        # Bind close event
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
        import_dm3 = import_menu.Append(wx.ID_ANY, "EELS Map (.dm3)...",
                                        "Import EELS map from DM3 file")
        import_dm4 = import_menu.Append(wx.ID_ANY, "EELS Map (.dm4)...",
                                        "Import EELS map from DM4 file")
        import_menu.AppendSeparator()
        import_hdf5 = import_menu.Append(wx.ID_ANY, "EELS Map (HDF5)...",
                                         "Import EELS map from HDF5 file")

        file_menu.AppendSubMenu(import_menu, "Import")

        # Export submenu
        export_menu = wx.Menu()
        export_excel = export_menu.Append(wx.ID_ANY, "Export to Excel...",
                                          "Export data to Excel")
        export_csv = export_menu.Append(wx.ID_ANY, "Export to CSV...",
                                        "Export spectrum to CSV")
        export_menu.AppendSeparator()
        export_map_image = export_menu.Append(wx.ID_ANY, "Export Map as Image...",
                                              "Save map as PNG/TIFF")

        file_menu.AppendSubMenu(export_menu, "Export")

        file_menu.AppendSeparator()
        close_item = file_menu.Append(wx.ID_CLOSE, "Close\tCtrl+W", "Close window")

        menubar.Append(file_menu, "&File")

        self.SetMenuBar(menubar)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_import_dm3, import_dm3)
        self.Bind(wx.EVT_MENU, self.on_import_dm4, import_dm4)
        self.Bind(wx.EVT_MENU, self.on_import_hdf5, import_hdf5)
        self.Bind(wx.EVT_MENU, self.on_export_excel, export_excel)
        self.Bind(wx.EVT_MENU, self.on_export_csv, export_csv)
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

        toolbar_sizer.Add(wx.StaticLine(toolbar_panel, style=wx.LI_VERTICAL),
                          0, wx.EXPAND | wx.ALL, 3)

        # Selection tools
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

        toolbar_sizer.Add(wx.StaticLine(toolbar_panel, style=wx.LI_VERTICAL),
                          0, wx.EXPAND | wx.ALL, 3)

        # Intensity Map button
        self.intensity_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        heatmap_path = os.path.join(icon_path, "heatmap-3.png")
        if os.path.exists(heatmap_path):
            self.intensity_btn.SetBitmap(wx.Bitmap(heatmap_path, wx.BITMAP_TYPE_PNG))
        else:
            self.intensity_btn.SetBitmap(self.create_icon_bitmap('heatmap'))
        self.intensity_btn.SetToolTip("Show Intensity Map")
        toolbar_sizer.Add(self.intensity_btn, 0, wx.ALL, 2)

        # Add spacer
        toolbar_sizer.AddStretchSpacer()

        # Sensitivity/Display controls button
        self.sensitivity_btn = wx.BitmapButton(toolbar_panel, size=btn_size)
        self.sensitivity_btn.SetBitmap(self.create_icon_bitmap('sensitivity'))
        self.sensitivity_btn.SetToolTip("Display Controls")
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
        elif icon_type == 'heatmap':
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
            dc.SetPen(wx.Pen(wx.Colour(79, 190, 159), 2))
            dc.DrawLine(5, 10, 15, 10)
            dc.SetBrush(wx.Brush(wx.Colour(79, 190, 159)))
            dc.DrawCircle(10, 10, 3)
            dc.DrawLine(10, 4, 10, 7)
            dc.DrawLine(10, 13, 10, 16)

        dc.SelectObject(wx.NullBitmap)
        return bmp

    # ==================== TOOLBAR HANDLERS ====================

    def on_sensitivity_controls(self, event):
        """Open sensitivity and display controls window"""
        if hasattr(self, 'sensitivity_window') and self.sensitivity_window:
            if not self.sensitivity_window.IsBeingDeleted():
                self.sensitivity_window.Raise()
                return
        self.sensitivity_window = EELSSensitivityWindow(self)
        self.sensitivity_window.Show()

    def on_zoom_in(self, event):
        """Toggle zoom in mode using rectangle selector"""
        if self.zoom_in_btn.GetValue():
            self.pan_btn.SetValue(False)
            self.deactivate_selection_modes()

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

        self.zoom_in_btn.SetValue(False)
        if hasattr(self, 'zoom_selector') and self.zoom_selector:
            self.zoom_selector.set_active(False)

    def on_zoom_out(self, event):
        """Zoom out / reset view"""
        self.map_ax.autoscale()
        self.map_canvas.draw()

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

            self._pan_start = None
            self._pan_cid_press = self.map_canvas.mpl_connect(
                'button_press_event', self._on_pan_press)
            self._pan_cid_release = self.map_canvas.mpl_connect(
                'button_release_event', self._on_pan_release)
            self._pan_cid_motion = self.map_canvas.mpl_connect(
                'motion_notify_event', self._on_pan_motion)
        else:
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

    def deactivate_selection_modes(self):
        """Deactivate all selection modes"""
        self.point_btn.SetValue(False)
        self.area_btn.SetValue(False)
        self.line_btn.SetValue(False)
        self.selection_mode = None

        if self.rect_selector:
            self.rect_selector.set_active(False)

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

            self.clear_selection_markers()

            if self.rect_selector is None:
                self.rect_selector = RectangleSelector(
                    self.map_ax,
                    self.on_area_select,
                    useblit=True,
                    props=dict(facecolor='green', edgecolor='lime',
                               alpha=0.3, fill=True),
                    button=[1],
                    minspanx=2,
                    minspany=2,
                    spancoords='pixels',
                    interactive=False
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

            self.clear_selection_markers()
            self.map_canvas.draw()
        else:
            self.selection_mode = None

    def on_clear_selections(self, event):
        """Clear all selections"""
        self.selected_points = []
        self.selected_areas = []
        self.line_start = None
        self.line_end = None

        self.clear_selection_markers()
        self._clear_line_preview()
        self.map_canvas.draw()

    def on_intensity_map(self, event):
        """Show intensity map"""
        if self.current_data is None:
            return
        self.plot_current_map()

    def on_right_click(self, event):
        """Handle right-click context menu"""
        menu = wx.Menu()

        # Colormap submenu
        cmap_menu = wx.Menu()
        cmaps = ['plasma', 'viridis', 'inferno', 'magma', 'hot', 'jet', 'gray']
        for cmap in cmaps:
            item = cmap_menu.AppendRadioItem(wx.ID_ANY, cmap)
            if cmap == self.current_cmap:
                item.Check(True)
            self.Bind(wx.EVT_MENU,
                      lambda e, c=cmap: self.change_colormap(c), item)
        menu.AppendSubMenu(cmap_menu, "Colormap")

        menu.AppendSeparator()

        # Reset view
        reset_item = menu.Append(wx.ID_ANY, "Reset View")
        self.Bind(wx.EVT_MENU, lambda e: self.on_zoom_out(None), reset_item)

        self.PopupMenu(menu)
        menu.Destroy()

    def change_colormap(self, cmap):
        """Change the colormap"""
        self.current_cmap = cmap
        self.plot_current_map()

    # ==================== SELECTION HANDLERS ====================

    def on_map_scroll(self, event):
        """Handle mouse scroll on map"""
        if event.inaxes != self.map_ax:
            return

        if self.selection_mode == 'point':
            if event.button == 'up':
                self.point_size = min(50, self.point_size + 1)
            elif event.button == 'down':
                self.point_size = max(1, self.point_size - 1)

            self._update_point_preview(event.xdata, event.ydata)

    def _update_point_preview(self, x, y):
        """Update point size preview rectangle"""
        self._clear_point_preview()

        if x is None or y is None:
            return

        if self.point_size == 1:
            marker, = self.map_ax.plot(x, y, 'g+', markersize=15,
                                       markeredgewidth=2, alpha=0.5)
            marker._is_point_preview = True
        else:
            half_size = self.point_size / 2
            rect = patches.Rectangle((x - half_size, y - half_size),
                                      self.point_size, self.point_size,
                                      linewidth=2, edgecolor='lime',
                                      facecolor='green', alpha=0.3)
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
            self.clear_selection_markers()
            self._clear_point_preview()

            self.selected_points = [(x, y, self.point_size)]

            if self.point_size == 1:
                marker, = self.map_ax.plot(x, y, 'g+', markersize=15,
                                           markeredgewidth=2)
                marker._is_selection_marker = True
            else:
                half_size = self.point_size / 2
                rect = patches.Rectangle((x - half_size, y - half_size),
                                          self.point_size, self.point_size,
                                          linewidth=2, edgecolor='lime',
                                          facecolor='green', alpha=0.3)
                rect._is_selection_marker = True
                self.map_ax.add_patch(rect)

            self.map_canvas.draw()
            self.plot_point_spectrum(x, y)

        elif self.selection_mode == 'line':
            if self.line_start is None:
                self.clear_selection_markers()
                self._clear_line_preview()

                self.line_start = (x, y)
                marker, = self.map_ax.plot(x, y, 'go', markersize=8,
                                           markeredgecolor='darkgreen',
                                           markeredgewidth=2)
                marker._is_selection_marker = True
                self.map_canvas.draw()
            else:
                self.line_end = (x, y)
                self._clear_line_preview()
                self.clear_selection_markers()

                line, = self.map_ax.plot([self.line_start[0], self.line_end[0]],
                                         [self.line_start[1], self.line_end[1]],
                                         'g-', linewidth=2)
                line._is_selection_marker = True

                marker1, = self.map_ax.plot(self.line_start[0], self.line_start[1],
                                            'go', markersize=8,
                                            markeredgecolor='darkgreen',
                                            markeredgewidth=2)
                marker1._is_selection_marker = True
                marker2, = self.map_ax.plot(self.line_end[0], self.line_end[1],
                                            'go', markersize=8,
                                            markeredgecolor='darkgreen',
                                            markeredgewidth=2)
                marker2._is_selection_marker = True

                self.map_canvas.draw()

                self.plot_line_spectrum(self.line_start[0], self.line_start[1],
                                        self.line_end[0], self.line_end[1])

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

        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        self.selected_areas.append((x_min, y_min, x_max, y_max))
        self.plot_area_spectrum(x_min, y_min, x_max, y_max)

    # ==================== SPECTRUM PLOTTING ====================

    def plot_point_spectrum(self, x, y):
        """Plot spectrum from point or area around point"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D (x, y, energy) for spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        point_size = getattr(self, 'point_size', 1)

        if point_size == 1:
            spectrum = data[y, x, :]
            title = f'EELS Spectrum at Point ({x}, {y})'
        else:
            half_size = point_size // 2
            y1 = max(0, y - half_size)
            y2 = min(data.shape[0], y + half_size + 1)
            x1 = max(0, x - half_size)
            x2 = min(data.shape[1], x + half_size + 1)

            spectrum = np.mean(data[y1:y2, x1:x2, :], axis=(0, 1))
            title = f'EELS Spectrum at Point ({x}, {y}) [{point_size}×{point_size} px]'

        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        # Plot in parent KherveFitting window
        if self.parent is not None and hasattr(self.parent, 'ax'):
            self._plot_to_parent(energy, spectrum, title, 'point', x, y, point_size)

    def plot_area_spectrum(self, x1, y1, x2, y2):
        """Plot summed spectrum from selected area"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            wx.MessageBox("Data must be 3D for area spectrum extraction",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        # Sum spectra from area
        spectrum = np.sum(data[y1:y2+1, x1:x2+1, :], axis=(0, 1))
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        num_pixels = (y2 - y1 + 1) * (x2 - x1 + 1)
        title = f'EELS Area Spectrum ({num_pixels} pixels)'

        if self.parent is not None and hasattr(self.parent, 'ax'):
            self._plot_to_parent(energy, spectrum, title, 'area',
                                 selection_info={'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2})

    def plot_line_spectrum(self, x1, y1, x2, y2):
        """Plot summed spectrum from line profile"""
        if self.current_data is None:
            return

        data = self.current_data.data

        if len(data.shape) != 3:
            return

        # Get line coordinates using Bresenham's algorithm
        num_points = int(max(abs(x2 - x1), abs(y2 - y1))) + 1
        x_coords = np.linspace(x1, x2, num_points).astype(int)
        y_coords = np.linspace(y1, y2, num_points).astype(int)

        # Clamp to valid range
        x_coords = np.clip(x_coords, 0, data.shape[1] - 1)
        y_coords = np.clip(y_coords, 0, data.shape[0] - 1)

        # Sum spectra along line
        spectrum = np.sum(data[y_coords, x_coords, :], axis=0)
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        title = f'EELS Line Spectrum ({num_points} pixels)'

        if self.parent is not None and hasattr(self.parent, 'ax'):
            self._plot_to_parent(energy, spectrum, title, 'line',
                                 selection_info={'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2})

    def _plot_to_parent(self, energy, spectrum, title, sel_type,
                        x=None, y=None, size=None, selection_info=None):
        """Plot spectrum to parent KherveFitting window and store data"""
        if self.parent is None:
            return

        # Create EELS~Plot sheet if it doesn't exist
        if 'EELS~Plot' not in self.parent.Data['Core levels']:
            self.parent.Data['Core levels']['EELS~Plot'] = {}

        # Build selection info
        if selection_info is None:
            selection_info = {}
        if sel_type == 'point':
            selection_info = {'type': 'point', 'x': x, 'y': y, 'size': size}
        elif sel_type:
            selection_info['type'] = sel_type

        # Store data in B.E./Raw Data format for compatibility
        self.parent.Data['Core levels']['EELS~Plot'] = {
            'Name': 'EELS~Plot',
            'B.E.': [float(f"{v:.2f}") for v in energy],
            'Raw Data': [float(f"{v:.2f}") for v in spectrum],
            '_EELS_selection': selection_info,
            '_EELS_type': 'plot'
        }

        # Update sheet combobox
        if 'EELS~Plot' not in [self.parent.sheet_combobox.GetString(i)
                               for i in range(self.parent.sheet_combobox.GetCount())]:
            self.parent.sheet_combobox.Append('EELS~Plot')
        self.parent.sheet_combobox.SetValue('EELS~Plot')

        # Plot
        self.parent.ax.clear()
        self.parent.ax.plot(energy, spectrum, 'k-', linewidth=0.8)
        self.parent.ax.set_xlabel('Energy Loss (eV)')
        self.parent.ax.set_ylabel('Counts')
        self.parent.ax.set_title(title)
        self.parent.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))
        self.parent.ax.grid(False)

        self.parent.canvas.draw()

    def plot_sum_spectrum_to_parent(self):
        """Plot sum spectrum from all pixels to parent window"""
        if self.current_data is None or self.parent is None:
            return

        data = self.current_data.data
        if len(data.shape) != 3:
            return

        # Sum over spatial dimensions
        spectrum = np.sum(data, axis=(0, 1))
        energy = self.get_energy_axis()

        if energy is None:
            energy = np.arange(len(spectrum))

        title = f'EELS Sum Spectrum ({data.shape[0]}×{data.shape[1]} px)'

        self._plot_to_parent(energy, spectrum, title, 'sum')

    def get_energy_axis(self):
        """Get energy axis from current data"""
        if self.current_data is None:
            return None

        if hasattr(self.current_data, 'axes_manager'):
            if len(self.current_data.axes_manager.signal_axes) > 0:
                return self.current_data.axes_manager.signal_axes[0].axis

        # Fallback: create channel numbers
        if len(self.current_data.data.shape) == 3:
            return np.arange(self.current_data.data.shape[2])
        elif len(self.current_data.data.shape) == 1:
            return np.arange(len(self.current_data.data))

        return None

    # ==================== FILE IMPORT ====================

    def on_import_dm3(self, event):
        """Import DM3 file"""
        self._import_file("DM3 files (*.dm3)|*.dm3|All files (*.*)|*.*")

    def on_import_dm4(self, event):
        """Import DM4 file"""
        self._import_file("DM4 files (*.dm4)|*.dm4|All files (*.*)|*.*")

    def on_import_hdf5(self, event):
        """Import HDF5 file"""
        self._import_file("HDF5 files (*.hdf5;*.h5)|*.hdf5;*.h5|All files (*.*)|*.*")

    def _import_file(self, wildcard):
        """Generic file import"""
        with wx.FileDialog(self, "Open EELS Map",
                           wildcard=wildcard,
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                file_path = dlg.GetPath()
                self.load_file(file_path)

                if self.current_data is not None:
                    self.create_eels_map_output(file_path)

    def load_file(self, file_path):
        """Load EELS file using HyperSpy"""
        try:
            import hyperspy.api as hs

            loaded_data = None

            # Try different readers
            readers_to_try = ['HSPY', 'Delmic', None]

            for reader in readers_to_try:
                try:
                    if reader:
                        loaded_data = hs.load(file_path, reader=reader)
                    else:
                        loaded_data = hs.load(file_path)
                    print(f"Successfully loaded with {reader or 'default'} reader")
                    break
                except Exception as e:
                    print(f"Failed with {reader or 'default'} reader: {e}")
                    continue

            if loaded_data is None:
                wx.MessageBox(f"Could not load file with any available reader.",
                              "Load Error", wx.OK | wx.ICON_ERROR)
                return

            # Handle list of signals
            if isinstance(loaded_data, list):
                if len(loaded_data) == 1:
                    loaded_data = loaded_data[0]
                else:
                    # Use first 3D signal found
                    for sig in loaded_data:
                        if len(sig.data.shape) == 3:
                            loaded_data = sig
                            break
                    else:
                        loaded_data = loaded_data[0]

            self.loaded_signals.append({
                'data': loaded_data,
                'filename': os.path.basename(file_path),
                'path': file_path
            })

            self.current_data = loaded_data
            self.plot_current_map()

        except Exception as e:
            wx.MessageBox(f"Error loading file:\n{str(e)}",
                          "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def create_eels_map_output(self, file_path):
        """Create output files and add EELS data to parent window.Data"""
        import json
        import shutil

        try:
            # Initialize parent Data structure if needed
            if self.parent is not None:
                if not hasattr(self.parent, 'Data'):
                    from libraries.ConfigFile import Init_Measurement_Data
                    self.parent.Data = Init_Measurement_Data(self.parent)

                if 'Core levels' not in self.parent.Data:
                    self.parent.Data['Core levels'] = {}

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            excel_path = os.path.join(os.path.dirname(file_path),
                                      f"{base_name}_EELS.xlsx")
            json_path = os.path.join(os.path.dirname(file_path),
                                     f"{base_name}_EELS.json")

            # Set FilePath
            if self.parent is not None and hasattr(self.parent, 'Data'):
                self.parent.Data['FilePath'] = excel_path

                if hasattr(self.parent, 'current_file_path'):
                    self.parent.current_file_path = excel_path

                if hasattr(self.parent, 'Working_directory'):
                    self.parent.Working_directory = os.path.dirname(excel_path)

            # Get energy range
            energy_axis = self.get_energy_axis()
            if energy_axis is not None:
                energy_min = f"{np.min(energy_axis):.2f}"
                energy_max = f"{np.max(energy_axis):.2f}"
                energy_range = f"{energy_min} - {energy_max} eV"
            else:
                energy_range = "N/A"

            # Create sum spectrum
            sum_signal = self.current_data.sum()
            spectrum_data = sum_signal.data

            # Create Excel file
            import openpyxl
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            # EELS~Plot sheet
            ws_plot = wb.create_sheet("EELS~Plot")
            ws_plot.append(['Energy Loss (eV)', 'Intensity', f'Range: {energy_range}'])

            for i, intensity in enumerate(spectrum_data):
                if energy_axis is not None:
                    ws_plot.append([f"{energy_axis[i]:.2f}", f"{intensity:.2f}"])
                else:
                    ws_plot.append([f"{i:.2f}", f"{intensity:.2f}"])

            # EELS~Map sheet
            ws_map = wb.create_sheet("EELS~Map")
            map_data = np.sum(self.current_data.data, axis=2)

            ws_map.append([f'EELS Intensity Map - Range: {energy_range}'])
            ws_map.append([''] * (map_data.shape[1] + 1))

            for row in map_data:
                ws_map.append([f"{val:.2f}" for val in row])

            wb.save(excel_path)
            print(f"EELS data exported to: {excel_path}")

            # Add to parent window.Data
            if self.parent is not None and hasattr(self.parent, 'Data'):
                energy_values = energy_axis if energy_axis is not None else np.arange(len(spectrum_data))

                self.parent.Data['Core levels']['EELS~Plot'] = {
                    'Name': 'EELS~Plot',
                    'B.E.': [float(f"{v:.2f}") for v in energy_values],
                    'Raw Data': [float(f"{v:.2f}") for v in spectrum_data],
                    '_EELS_type': 'plot'
                }

                self.parent.Data['Core levels']['EELS~Map'] = {
                    'Name': 'EELS~Map',
                    'Map_Intensity': map_data.tolist(),
                    'Map_Shape': list(map_data.shape),
                    'Energy_Range': energy_range,
                    '_EELS_type': 'map',
                    '_Source_Path': file_path
                }

                # Update UI
                if hasattr(self.parent, 'SetStatusText'):
                    self.parent.SetStatusText(
                        f"Working Directory: {os.path.dirname(excel_path)}", 0)

                if hasattr(self.parent, 'SetTitle'):
                    self.parent.SetTitle(
                        f"KherveFitting - {os.path.basename(excel_path)}")

                # Add sheets to combobox
                for sheet_name in ['EELS~Plot', 'EELS~Map']:
                    if sheet_name not in [self.parent.sheet_combobox.GetString(i)
                                          for i in range(self.parent.sheet_combobox.GetCount())]:
                        self.parent.sheet_combobox.Append(sheet_name)

            # Create JSON file
            json_data = {
                'FilePath': excel_path,
                'Core levels': {
                    'EELS~Plot': {
                        'Name': 'EELS~Plot',
                        'B.E.': [float(f"{v:.2f}") for v in energy_values],
                        'Raw Data': [float(f"{v:.2f}") for v in spectrum_data],
                        '_EELS_type': 'plot'
                    },
                    'EELS~Map': {
                        'Name': 'EELS~Map',
                        'Map_Intensity': [[float(f"{val:.2f}") for val in row]
                                          for row in map_data],
                        'Map_Shape': list(map_data.shape),
                        'Energy_Range': energy_range,
                        '_EELS_type': 'map',
                        '_Source_Path': file_path
                    }
                }
            }

            with open(json_path, 'w') as jf:
                json.dump(json_data, jf, indent=2)
            print(f"JSON data saved to: {json_path}")

            wx.MessageBox(f"EELS data exported to:\n{excel_path}\n\nJSON: {json_path}",
                          "Export Complete", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error creating EELS output:\n{str(e)}",
                          "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

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
            title = 'EELS Image'
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

        # Remove axis labels - use scale bar instead
        self.map_ax.set_xlabel('')
        self.map_ax.set_ylabel('')
        self.map_ax.set_title('')
        self.map_ax.set_xticks([])
        self.map_ax.set_yticks([])

        # Add scale bar
        self._add_scale_bar()

        # Add colorbar
        from mpl_toolkits.axes_grid1 import make_axes_locatable
        divider = make_axes_locatable(self.map_ax)
        cax = divider.append_axes("right", size="5%", pad=0.05)
        self.current_colorbar = self.map_figure.colorbar(im, cax=cax)

        self.map_figure.tight_layout(pad=0.5)
        self._add_scale_bar()

        self.map_canvas.draw()

        # Plot sum spectrum in parent window
        self.plot_sum_spectrum_to_parent()

        self.reinitialize_selectors()

    def _add_scale_bar(self):
        """Add scale bar to map if scale information is available"""
        self._remove_scale_bar()

        if self.current_data is None:
            return

        # Try to get scale from axes_manager
        scale_value = None
        units = 'px'

        if hasattr(self.current_data, 'axes_manager'):
            nav_axes = self.current_data.axes_manager.navigation_axes
            if nav_axes:
                axis = nav_axes[0]
                scale_value = axis.scale
                units = axis.units if axis.units else 'px'

        if scale_value is None:
            return

        # Get current view limits
        xlim = self.map_ax.get_xlim()
        ylim = self.map_ax.get_ylim()

        view_width = abs(xlim[1] - xlim[0])
        view_height = abs(ylim[1] - ylim[0])

        # Calculate appropriate scale bar length
        target_fraction = 0.2
        target_pixels = view_width * target_fraction
        target_physical = target_pixels * scale_value

        # Round to nice number
        nice_values = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000]
        scale_bar_physical = min(nice_values, key=lambda x: abs(x - target_physical))
        scale_bar_pixels = scale_bar_physical / scale_value

        # Position in lower right
        margin_x = view_width * 0.05
        margin_y = view_height * 0.05

        x_start = xlim[1] - margin_x - scale_bar_pixels
        y_pos = max(ylim[0], ylim[1]) - margin_y

        bar_height = view_height * 0.02

        # Draw scale bar
        rect = patches.Rectangle((x_start, y_pos - bar_height), scale_bar_pixels, bar_height,
                                  linewidth=1, edgecolor='white', facecolor='white',
                                  zorder=101)
        rect._is_scale_bar = True
        self.map_ax.add_patch(rect)

        # Add label
        if units == 'nm':
            if scale_bar_physical >= 1000:
                label = f'{scale_bar_physical/1000:.0f} µm'
            else:
                label = f'{scale_bar_physical:.0f} nm'
        else:
            label = f'{scale_bar_physical:.0f} {units}'

        text = self.map_ax.text(x_start + scale_bar_pixels / 2,
                                y_pos - bar_height - margin_y * 0.5,
                                label, ha='center', va='top', fontsize=9,
                                fontweight='bold', color='white', zorder=102,
                                path_effects=[path_effects.withStroke(
                                    linewidth=2, foreground='black')])
        text._is_scale_bar = True

    def _remove_scale_bar(self):
        """Remove existing scale bar elements"""
        for artist in self.map_ax.patches[:]:
            if hasattr(artist, '_is_scale_bar') and artist._is_scale_bar:
                artist.remove()
        for artist in self.map_ax.texts[:]:
            if hasattr(artist, '_is_scale_bar') and artist._is_scale_bar:
                artist.remove()

    def draw_saved_selection(self, selection_info):
        """Draw a saved selection on the map"""
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
                half_size = size / 2
                rect = patches.Rectangle((x - half_size, y - half_size), size, size,
                                          linewidth=2, edgecolor='black', facecolor='none')
                rect._is_saved_selection = True
                self.map_ax.add_patch(rect)

        elif sel_type == 'line':
            line, = self.map_ax.plot([selection_info['x1'], selection_info['x2']],
                                     [selection_info['y1'], selection_info['y2']],
                                     'k-', linewidth=2)
            line._is_saved_selection = True

        elif sel_type == 'area':
            x1, y1 = selection_info['x1'], selection_info['y1']
            x2, y2 = selection_info['x2'], selection_info['y2']
            rect = patches.Rectangle((x1, y1), x2 - x1, y2 - y1,
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

    def reinitialize_selectors(self):
        """Reinitialize rectangle selector after redraw"""
        if self.rect_selector is not None:
            try:
                # Recreate selector
                self.rect_selector = RectangleSelector(
                    self.map_ax,
                    self.on_area_select,
                    useblit=True,
                    props=dict(facecolor='green', edgecolor='lime',
                               alpha=0.3, fill=True),
                    button=[1],
                    minspanx=2,
                    minspany=2,
                    spancoords='pixels',
                    interactive=False
                )
                if self.selection_mode != 'area':
                    self.rect_selector.set_active(False)
            except:
                pass

    # ==================== EXPORT ====================

    def on_export_excel(self, event):
        """Export data to Excel"""
        if self.current_data is None:
            wx.MessageBox("No data loaded", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Save Excel File",
                           wildcard="Excel files (*.xlsx)|*.xlsx",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self._export_to_excel(dlg.GetPath())

    def _export_to_excel(self, path):
        """Export current data to Excel"""
        import openpyxl

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "EELS Spectrum"

            energy = self.get_energy_axis()
            spectrum = np.sum(self.current_data.data, axis=(0, 1))

            ws.append(['Energy Loss (eV)', 'Intensity'])
            for i, val in enumerate(spectrum):
                e = energy[i] if energy is not None else i
                ws.append([f"{e:.2f}", f"{val:.2f}"])

            wb.save(path)
            wx.MessageBox(f"Exported to: {path}", "Export Complete",
                          wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Export error: {e}", "Error", wx.OK | wx.ICON_ERROR)

    def on_export_csv(self, event):
        """Export spectrum to CSV"""
        if self.current_data is None:
            wx.MessageBox("No data loaded", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Save CSV File",
                           wildcard="CSV files (*.csv)|*.csv",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self._export_to_csv(dlg.GetPath())

    def _export_to_csv(self, path):
        """Export current spectrum to CSV"""
        try:
            energy = self.get_energy_axis()
            spectrum = np.sum(self.current_data.data, axis=(0, 1))

            with open(path, 'w') as f:
                f.write("Energy Loss (eV),Intensity\n")
                for i, val in enumerate(spectrum):
                    e = energy[i] if energy is not None else i
                    f.write(f"{e:.2f},{val:.2f}\n")

            wx.MessageBox(f"Exported to: {path}", "Export Complete",
                          wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Export error: {e}", "Error", wx.OK | wx.ICON_ERROR)

    def on_export_map_image(self, event):
        """Export map as image"""
        if self.current_data is None:
            wx.MessageBox("No data loaded", "Error", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(self, "Save Map Image",
                           wildcard="PNG files (*.png)|*.png|TIFF files (*.tiff)|*.tiff",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.map_figure.savefig(dlg.GetPath(), dpi=300, bbox_inches='tight')
                wx.MessageBox(f"Map saved to: {dlg.GetPath()}",
                              "Export Complete", wx.OK | wx.ICON_INFORMATION)

    # ==================== WINDOW MANAGEMENT ====================

    def on_close(self, event):
        """Handle window close event"""
        # Clear parent reference
        if self.parent is not None and hasattr(self.parent, 'eels_window'):
            self.parent.eels_window = None

        # Destroy sensitivity window if open
        if hasattr(self, 'sensitivity_window') and self.sensitivity_window:
            try:
                self.sensitivity_window.Destroy()
            except:
                pass

        self.Destroy()


class EELSSensitivityWindow(wx.Frame):
    """Window for EELS display controls"""

    def __init__(self, parent):
        super().__init__(parent, title="Display Controls", size=(300, 200),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)

        self.parent = parent
        self.init_ui()
        self.Centre()

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Colormap selection
        cmap_sizer = wx.BoxSizer(wx.HORIZONTAL)
        cmap_label = wx.StaticText(panel, label="Colormap:")
        self.cmap_combo = wx.ComboBox(panel,
                                      choices=['plasma', 'viridis', 'inferno',
                                               'magma', 'hot', 'jet', 'gray'],
                                      style=wx.CB_READONLY)
        self.cmap_combo.SetValue(self.parent.current_cmap)
        self.cmap_combo.Bind(wx.EVT_COMBOBOX, self.on_cmap_change)

        cmap_sizer.Add(cmap_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        cmap_sizer.Add(self.cmap_combo, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(cmap_sizer, 0, wx.EXPAND)

        # Brightness/Contrast (placeholder for future)
        sizer.Add(wx.StaticLine(panel), 0, wx.EXPAND | wx.ALL, 5)

        info_text = wx.StaticText(panel, label="Scroll wheel adjusts point size\n"
                                               "when in point selection mode.")
        sizer.Add(info_text, 0, wx.ALL, 10)

        panel.SetSizer(sizer)

    def on_cmap_change(self, event):
        self.parent.current_cmap = self.cmap_combo.GetValue()
        self.parent.plot_current_map()


def open_eels_window(parent):
    """Open EELS analysis window"""
    window = EELSWindow(parent)
    window.Show()
    return window