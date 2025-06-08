# KherveFitting------------------------------------------------------------------


# LIBRARIES----------------------------------------------------------------------
import multiprocessing
import os
import psutil

os.environ['OMP_NUM_THREADS'] = str(multiprocessing.cpu_count())
os.environ['MKL_NUM_THREADS'] = str(multiprocessing.cpu_count())
os.environ['OPENBLAS_NUM_THREADS'] = str(multiprocessing.cpu_count())
os.environ['VECLIB_MAXIMUM_THREADS'] = str(multiprocessing.cpu_count())
os.environ['NUMEXPR_NUM_THREADS'] = str(multiprocessing.cpu_count())


import matplotlib
from matplotlib.backends.backend_wxagg import NavigationToolbar2WxAgg as NavigationToolbar

import wx.grid
import wx.adv
import sys
import platform
import re

from matplotlib.ticker import ScalarFormatter, AutoMinorLocator
matplotlib.use('WXAgg')  # Use WXAgg backend for wxPython compatibility
from libraries.Fitting_Screen import *
from libraries.AreaFit_Screen import *
from libraries.Save import *
from libraries.NoiseAnalysis import NoiseAnalysisWindow
from libraries.ConfigFile import *
from libraries.Export import export_results
from libraries.PlotConfig import PlotConfig
from libraries.Utilities import check_first_time_use, DraggableText
from libraries.Plot_Operations import PlotManager
from libraries.Peak_Functions import PeakFunctions

# from libraries.Peak_Functions import AtomicConcentrations
# from libraries.Peak_Functions import gauss_lorentz, S_gauss_lorentz

from Functions import toggle_Col_1, update_sheet_names, rename_sheet, on_sheet_selected_wrapper
from libraries.PreferenceWindow import PreferenceWindow

# from libraries.Sheet_Operations import on_sheet_selected

from libraries.Sheet_Operations import CheckboxRenderer, on_sheet_selected
from libraries.SplashScreen import show_splash
from libraries.Save import save_state, undo, redo

# from libraries.Open import ExcelDropTarget
# from libraries.Utilities import copy_cell, paste_cell

from Functions import on_save, on_exit
from libraries.Open import load_recent_files_from_config, open_avg_file
from libraries.survey import PeriodicTableWindow
from libraries.Widgets_Toolbars import create_widgets, create_menu
from libraries.Widgets_Toolbars import create_statusbar, update_statusbar
from libraries.Open import open_xlsx_file

# from libraries.Export import export_word_report

from libraries.Peak_Functions import OtherCalc, AtomicConcentrations
from libraries.Dpara_Screen import DParameterWindow
from libraries.Update import UpdateChecker

from libraries.PeakFittingGrid import PeakFittingGrid
from libraries.PeakManipulation import PeakManipulation
from libraries.On_Key_Defs import setup_key_handlers
from libraries.On_Mouse_Defs import setup_mouse_handlers

from libraries.QuickSettings import QuickSettings




class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(1470, 680))

        # Get the directory of the current script
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # Construct the path to the icon file
        icon_path = os.path.join(current_dir, "Icons", "Icon.ico")

        FIRST_TIME_USE = True

        # Set the icon
        icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        # Initialise Peak Fitting Grid
        self.peak_fitting_grid = PeakFittingGrid(self)
        self.peak_manipulation = PeakManipulation(self)


        self.SetMinSize((800, 600))
        self.panel = wx.Panel(self)

        # Will hold reference to FileManagerWindow when opened
        self.file_manager = None

        # Center the window on the screen
        self.Centre()

        self.Data = Init_Measurement_Data(self)

        #
        self.previous_size = None
        self.file_manager_position = None

        self.last_popup_time = 0

        self.shift_key_pressed = False
        self.actual_fwhms = {}  # Store calculated FWHM values

        # Initial folder path
        self.Working_directory =  os.path.join(os.path.dirname(os.path.abspath(__file__)), "Data")

        self.selected_files = []     # Initialize selected_files variable
        self.selected_indices = []  # Initialize selected_indices variable

        # BKG selected
        self.background_tab_selected = False
        self.peak_fitting_tab_selected = False
        self.fitting_window = None

        self.noise_analysis_window = None
        self.noise_tab_selected = False

        # Variables for Undo & Redo
        self.history = []
        self.redo_stack = []
        self.max_history = 50

        # New attribute to track plot state for showing fit or not
        self.show_fit = True

        # For FWHM calculation
        self.selected_peak_index = 0
        self.initial_fwhm = None
        self.initial_x = None

        self.peak_params = []  # Initialize
        self.peak_count = 0  # To keep track of the number of peaks

        # Dictionary for the X and Y limits
        self.plot_config = PlotConfig()

        self.is_right_panel_hidden = False

        self.peak_letter = None
        self.peak_info = None
        self.peak_letter_t = None
        self.peak_info_t = None

        # X axis correction from KE to BE
        self.photons = 1486.67
        self.workfunction = 0
        self.ref_peak_name = "C1s C-C"
        self.ref_peak_be = 284.8

        # Add a state variable for energy mode (default to Binding Energy)
        self.energy_scale = 'BE'  # 'BE' for binding energy, 'KE' for kinetic energy
        self.toggle_energy_item = None  # Will be set in create_menu

        # Initialize variables for vertical lines and background energy
        self.vline1 = None
        self.vline2 = None
        self.active_vline = None
        self.vline_drag_threshold = 5  # pixels


        self.vline3 = None  # For noise range
        self.vline4 = None  # For noise range
        self.some_threshold = 0.1
        self.bg_min_energy = None
        self.bg_max_energy = None
        self.noise_min_energy = None
        self.noise_max_energy = None

        # Initial max iteration value
        self.max_iterations = 50
        # Initial fitting method
        self.selected_fitting_method = "GL (Area)"

        # self.fitting_results_visible = False

        self.residuals_state = 0  # Add this line
        self.residuals_subplot = None  # Add this line
        self.legend_visible = 1  # Add this line
        self.y_axis_state = 0  # Add this line
        self.rsd_text = None  # Add this line

        self.background_method = "Multi-Regions Smart"
        self.offset_h = 0
        self.offset_l = 0

        # Initial background method
        self.selected_bkg_method  = "Multi-Regions Smart"

        # initial state for residual
        self.Data['residuals_state'] = 2  # 0: off, 1: on main plot, 2: separate subplot


        # Zoom initialisation variables
        self.zoom_mode = False
        self.zoom_rect = None

        # Drag button initialisation
        self.drag_mode = False
        self.drag_tool = None

        # Initialize attributes for background and noise data
        self.background = None
        self.noise_data = None

        self.x_values = None  # To store x values for noise plotting
        self.y_values = None  # To store y values for noise plotting

        # initialise peak fill or not
        self.peak_fill_enabled = True

        # Add new attributes for text settings
        self.plot_font = 'DejaVu Sans'
        self.axis_title_size = 12
        self.axis_number_size = 10
        self.x_sublines = 5
        self.y_sublines = 5
        self.legend_font_size = 8
        self.label_font_size = 8
        self.core_level_text_size = 15
        self.x_axis_label = "Binding Energy (eV)"
        self.y_axis_label = "Intensity (CPS)"

        # Initialize right_frame to None before creating widgets
        self.right_frame = None
        self.figure = plt.figure()
        self.ax = self.figure.add_subplot(111)  # Create the Matplotlib axes
        self.ax.set_xlabel("Binding Energy (eV)")  # Set x-axis label
        self.ax.set_ylabel("Intensity (CPS)")  # Set y-axis label
        self.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        # Apply text settings
        plt.rcParams['font.family'] = self.plot_font
        plt.rcParams['figure.dpi'] = 100
        plt.rcParams['savefig.dpi'] = 100
        self.ax.tick_params(axis='both', labelsize=self.axis_number_size)
        if self.x_sublines > 0:
            self.ax.xaxis.set_minor_locator(AutoMinorLocator(self.x_sublines + 1))
        if self.y_sublines > 0:
            self.ax.yaxis.set_minor_locator(AutoMinorLocator(self.y_sublines + 1))

        self.moving_vline = None  # Initialize moving_vline attribute
        self.selected_peak_index = None  # Add this attribute to track selected peak index

        self.be_correction = 0.00

        self.current_instrument = 'A-ALTHERMO01'  # Default instrument
        self.library_data = load_library_data()
        # data, instruments = load_library_data()

        self.averaging_points = 5

        # Number of column to remove from the excel file
        self.num_fitted_columns = 15

        # Initialize plot preference attributes with default values
        self.plot_style = "scatter"
        self.scatter_size = 4
        self.line_width = 1
        self.line_alpha = 0.7
        self.scatter_color = "#000000"
        self.line_color = "#000000"
        self.scatter_marker = "o"
        self.background_color = "#808080"
        self.background_alpha = 0.5
        self.background_linestyle = "--"
        self.background_thickness = 1
        self.envelope_color = "#0000FF"
        self.envelope_alpha = 0.6
        self.envelope_linestyle = "-"
        self.envelope_thickness = 1
        self.residual_color = "#00FF00"
        self.residual_alpha = 0.4
        self.residual_linestyle = "-"
        self.raw_data_linestyle = "-"
        self.residual_thickness = 1
        self.hatch_density = 2

        self.peak_colors = ["#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
                            "#00FFFF", "#800000", "#008000", "#000080", "#808000",
                            "#800080", "#008080", "#C0C0C0", "#808080", "#9B30FF"]
        self.peak_alpha = 0.3

        self.peak_line_style = "Same Color"  # Options: "no_line", "black", "same_color", "grey"
        self.peak_line_alpha = 0.7
        self.peak_line_thickness = 1
        self.peak_line_pattern = '-'  # Options: '-', '--', '-.', ':'

        # Hatch init
        self.peak_fill_types = ["Solid Fill" for _ in range(15)]  # List of types for each peak
        self.peak_hatch_patterns = ["/", "\\", "|", "-", "+", "x", "o", "O", ".", "*"] * 2

        # Add to plot preferences section
        self.excel_width = 5.2
        self.excel_height = 5.2
        self.excel_dpi = 100
        self.survey_excel_width = 10
        self.survey_excel_height = 5
        self.survey_excel_dpi = 100

        # Most recent File Initialisation
        self.recent_files = []
        self.max_recent_files = 20  # Maximum number of recent files to keep

        self.library_type = "TPP-2M"  # Default value

        # Add Backup initializations
        self.backup_timer = None
        self.enable_auto_backup = False
        self.backup_interval = 30  # Default to 30 minutes

        # Initialize registered flag
        self.registered = False

        self.load_config()
        # If you stored the config as an attribute:
        if not hasattr(self, 'registered') or not self.registered:
            self.times_opened = 0

        # Initialize quick settings
        self.quick_settings = QuickSettings(self)

        create_widgets(self)
        create_menu(self)
        load_recent_files_from_config(self)

        # Set initial splitter position to make plot ~900 pixels wide
        if hasattr(self, 'splitter'):
            self.splitter.SetSashPosition(650)
            self.initial_sash_position = 650

        # Add specific event handling for checkbox clicks
        self.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.on_grid_cell_click)

        # Start Backup timer
        self.setup_backup_timer()



        create_statusbar(self)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)


        # self.canvas.mpl_connect("button_press_event", self.on_click)
        # self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        # self.canvas.mpl_connect('scroll_event', self.on_mouse_wheel)
        # self.canvas.mpl_connect('button_press_event', self.on_right_click)
        self.mouse_handler = setup_mouse_handlers(self)

        # self.canvas.mpl_connect('key_press_event', self.on_key_press)
        # self.Bind(wx.EVT_CHAR_HOOK, self.on_key_press_global)
        setup_key_handlers(self)

        self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.peak_fitting_grid.on_peak_params_cell_changed)
        self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGING, self.peak_fitting_grid.on_peak_params_cell_changed)
        self.Bind(wx.EVT_CLOSE, self.on_close)

        # self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_peak_params_right_click)

        self.Bind(wx.EVT_SIZE, self.on_window_resize)

        self.plot_manager = PlotManager(self.ax, self.canvas, self)
        self.plot_manager.peak_colors = self.peak_colors.copy()
        self.plot_manager.residuals_state = self.residuals_state
        self.plot_manager.legend_visible = self.legend_visible
        self.plot_manager.y_axis_state = self.y_axis_state


    def on_peak_params_context_menu(self, event):
        # Alternative event handler for context menu
        position = event.GetPosition()
        # Convert screen position to grid cell
        x, y = self.peak_params_grid.ScreenToClient(position)
        row, col = self.peak_params_grid.XYToCell(x, y)
        if row != -1 and col != -1:  # Valid cell
            peak_index = row // 2

            menu = wx.Menu()
            copy_item = menu.Append(wx.ID_ANY, "Copy Peak Parameters")
            paste_item = menu.Append(wx.ID_ANY, "Paste Peak Parameters")

            import os
            import tempfile
            clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')
            paste_item.Enable(os.path.exists(clipboard_file))

            from libraries.Save import copy_peak_parameters, paste_peak_parameters
            self.Bind(wx.EVT_MENU, lambda evt: copy_peak_parameters(self, peak_index), copy_item)
            self.Bind(wx.EVT_MENU, lambda evt: paste_peak_parameters(self, peak_index), paste_item)

            self.peak_params_grid.PopupMenu(menu)
            menu.Destroy()
        else:
            event.Skip()


    def add_toggle_tool(self, toolbar, label, bmp):
        tool = toolbar.AddTool(wx.ID_ANY, label, bmp, kind=wx.ITEM_NORMAL)
        self.Bind(wx.EVT_TOOL, self.on_toggle_right_panel, tool)
        return tool


    def on_toggle_right_panel(self, event):
        splitter_width = self.splitter.GetSize().GetWidth()
        current_size = self.GetSize()

        if self.is_right_panel_hidden:
            # The right panel is currently hidden, so show it
            new_sash_position = self.initial_sash_position
            new_bmp = wx.ArtProvider.GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR)
            self.is_right_panel_hidden = False
            self.SetMinSize((800, 600))  # Reset min size to allow resizing
            self.SetMaxSize((-1, -1))
            # Restore previous size
            if hasattr(self, 'previous_size'):
                self.SetSize(self.previous_size)

            # Save current sheet selection before destroying toolbar
            old_sheet_value = ""
            if hasattr(self, 'sheet_combobox') and self.sheet_combobox:
                old_sheet_value = self.sheet_combobox.GetValue()

            # Restore the original toolbar
            toolbar_panel = self.panel.GetChildren()[0]  # First child is toolbar panel
            toolbar_sizer = toolbar_panel.GetSizer()

            # Remove current toolbar
            if self.toolbar:
                toolbar_sizer.Detach(self.toolbar)
                self.toolbar.Destroy()

            # Import here to avoid circular imports
            from libraries.Widgets_Toolbars import create_horizontal_toolbar
            self.toolbar = create_horizontal_toolbar(toolbar_panel, self)
            toolbar_sizer.Add(self.toolbar, 0, wx.EXPAND)

            # Repopulate sheet combobox and restore selection
            if 'Core levels' in self.Data:
                sheets = list(self.Data['Core levels'].keys())
                self.sheet_combobox.Clear()
                self.sheet_combobox.AppendItems(sheets)
                if old_sheet_value in sheets:
                    self.sheet_combobox.SetValue(old_sheet_value)
                elif sheets:
                    self.sheet_combobox.SetValue(sheets[0])

            toolbar_panel.Layout()
        else:
            # The right panel is currently visible, so hide it
            new_sash_position = splitter_width
            new_bmp = wx.ArtProvider.GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR)
            self.is_right_panel_hidden = True

            # Store current size and set to fixed width of 865
            self.previous_size = current_size
            fixed_width = 865
            fixed_height = current_size.height

            # Fix both width and height
            self.SetSize((fixed_width, fixed_height))
            self.SetMinSize((fixed_width, fixed_height))
            self.SetMaxSize((fixed_width, fixed_height))  # Fix both dimensions

            # Create minimal toolbar
            toolbar_panel = self.panel.GetChildren()[0]
            toolbar_sizer = toolbar_panel.GetSizer()

            # Current toolbar value to save sheet selection
            old_sheet_value = ""
            old_sheets = []
            if self.sheet_combobox:
                old_sheet_value = self.sheet_combobox.GetValue()
                old_sheets = [self.sheet_combobox.GetString(i) for i in range(self.sheet_combobox.GetCount())]

            # Remove current toolbar
            if self.toolbar:
                toolbar_sizer.Detach(self.toolbar)
                self.toolbar.Destroy()

            # Create new minimal toolbar
            current_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(current_dir, "libraries", "Icons")
            if not os.path.exists(icon_path):
                icon_path = os.path.join(current_dir, "Icons")

            self.toolbar = wx.ToolBar(toolbar_panel, style=wx.TB_FLAT)
            self.toolbar.SetToolBitmapSize(wx.Size(25, 25))

            # Define the rename dialog function here
            def _show_rename_dialog(window):
                dlg = wx.TextEntryDialog(window, 'Enter new sheet name:', 'Rename Sheet')
                if dlg.ShowModal() == wx.ID_OK:
                    from libraries.Utilities import rename_sheet
                    rename_sheet(window, dlg.GetValue())
                dlg.Destroy()

            # Add all requested tools
            open_tool = self.toolbar.AddTool(wx.ID_ANY, 'Open File',
                                             wx.Bitmap(os.path.join(icon_path, "open-folder-25-green.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Open File\tCtrl+O")

            save_tool = self.toolbar.AddTool(wx.ID_ANY, 'Save',
                                             wx.Bitmap(os.path.join(icon_path, "save-Excel-25.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Save")

            save_all_tool = self.toolbar.AddTool(wx.ID_ANY, 'Save All Sheets',
                                                 wx.Bitmap(os.path.join(icon_path, "save-Multi-25.png"),
                                                           wx.BITMAP_TYPE_PNG),
                                                 shortHelp="Save all sheets with plots")

            # Undo/Redo
            self.undo_tool = self.toolbar.AddTool(wx.ID_ANY, 'Undo',
                                                  wx.Bitmap(os.path.join(icon_path, "undo-25.png"),
                                                            wx.BITMAP_TYPE_PNG),
                                                  shortHelp="Undo")
            self.redo_tool = self.toolbar.AddTool(wx.ID_ANY, 'Redo',
                                                  wx.Bitmap(os.path.join(icon_path, "redo-25.png"),
                                                            wx.BITMAP_TYPE_PNG),
                                                  shortHelp="Redo")

            # File manager
            file_manager_tool = self.toolbar.AddTool(wx.ID_ANY, "Sample Manager",
                                                     wx.Bitmap(os.path.join(icon_path, "list-view-25.png"),
                                                               wx.BITMAP_TYPE_PNG),
                                                     shortHelp="Sample Manager")

            # Sheet combobox
            self.sheet_combobox = wx.ComboBox(self.toolbar, style=wx.CB_READONLY)
            if 'Core levels' in self.Data:
                sheets = list(self.Data['Core levels'].keys())
                self.sheet_combobox.AppendItems(sheets)
                if old_sheet_value in sheets:
                    self.sheet_combobox.SetValue(old_sheet_value)
                elif sheets:
                    self.sheet_combobox.SetValue(sheets[0])
            self.toolbar.AddControl(self.sheet_combobox)

            # Refresh, Delete, Copy, Join, Rename
            refresh_tool = self.toolbar.AddTool(wx.ID_ANY, 'Refresh',
                                                wx.Bitmap(os.path.join(icon_path, "Refresh-25.png"),
                                                          wx.BITMAP_TYPE_PNG),
                                                shortHelp="Refresh")
            delete_tool = self.toolbar.AddTool(wx.ID_ANY, 'Delete',
                                               wx.Bitmap(os.path.join(icon_path, "delete-25.png"),
                                                         wx.BITMAP_TYPE_PNG),
                                               shortHelp="Delete")
            copy_tool = self.toolbar.AddTool(wx.ID_ANY, 'Copy',
                                             wx.Bitmap(os.path.join(icon_path, "copy-25.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Copy")
            join_tool = self.toolbar.AddTool(wx.ID_ANY, 'Join',
                                             wx.Bitmap(os.path.join(icon_path, "join2-25.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Join")
            rename_tool = self.toolbar.AddTool(wx.ID_ANY, 'Rename',
                                               wx.Bitmap(os.path.join(icon_path, "rename-25.png"),
                                                         wx.BITMAP_TYPE_PNG),
                                               shortHelp="Rename")

            # BE correction
            self.be_correction_spinbox = wx.SpinCtrlDouble(self.toolbar, value='0.00', min=-10000.00, max=10000.00,
                                                           inc=0.01, size=(70, -1))
            self.be_correction_spinbox.SetDigits(2)
            self.be_correction_spinbox.SetValue(self.be_correction)
            self.toolbar.AddControl(self.be_correction_spinbox)

            auto_be_tool = self.toolbar.AddTool(wx.ID_ANY, 'Auto BE',
                                                wx.Bitmap(os.path.join(icon_path, "BEcorrect-25.png"),
                                                          wx.BITMAP_TYPE_PNG),
                                                shortHelp="Auto BE")

            # Add the requested tools
            bkg_tool = self.toolbar.AddTool(wx.ID_ANY, 'Background/Area',
                                            wx.Bitmap(os.path.join(icon_path, "BKG-25.png"),
                                                      wx.BITMAP_TYPE_PNG),
                                            shortHelp="Calculate Area Under Curve\tCtrl+A")

            peak_fit_tool = self.toolbar.AddTool(wx.ID_ANY, 'Peak Fit',
                                                 wx.Bitmap(os.path.join(icon_path, "C1s-25.png"),
                                                           wx.BITMAP_TYPE_PNG),
                                                 shortHelp="Peak Fit")

            diff_tool = self.toolbar.AddTool(wx.ID_ANY, 'D-parameter',
                                             wx.Bitmap(os.path.join(icon_path, "Dpara-25.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="D-parameter Calculation")

            id_tool = self.toolbar.AddTool(wx.ID_ANY, 'Element ID',
                                           wx.Bitmap(os.path.join(icon_path, "ID-25.png"),
                                                     wx.BITMAP_TYPE_PNG),
                                           shortHelp="Element identifications (ID)")

            # Toggle right panel tool
            self.toggle_right_panel_tool = self.toolbar.AddTool(wx.ID_ANY, "Toggle Right Panel",
                                                                wx.ArtProvider.GetBitmap(wx.ART_GO_FORWARD,
                                                                                         wx.ART_TOOLBAR),
                                                                shortHelp="Toggle Right Panel")

            self.toolbar.Realize()
            toolbar_sizer.Add(self.toolbar, 0, wx.EXPAND)
            toolbar_panel.Layout()

            # Rebind events
            from libraries.Open import open_xlsx_file
            from Functions import on_save, refresh_sheets, save_all_sheets_with_plots
            from libraries.Sheet_Operations import on_sheet_selected
            from libraries.Utilities import on_delete_sheet, copy_sheet, JoinSheetsWindow

            self.Bind(wx.EVT_TOOL, lambda e: open_xlsx_file(self), open_tool)
            self.Bind(wx.EVT_TOOL, lambda e: on_save(self), save_tool)
            self.Bind(wx.EVT_TOOL, lambda e: save_all_sheets_with_plots(self), save_all_tool)
            self.Bind(wx.EVT_TOOL, lambda e: undo(self), self.undo_tool)
            self.Bind(wx.EVT_TOOL, lambda e: redo(self), self.redo_tool)
            self.Bind(wx.EVT_TOOL, self.on_open_file_manager, file_manager_tool)
            self.Bind(wx.EVT_TOOL, lambda e: refresh_sheets(self, on_sheet_selected_wrapper), refresh_tool)
            self.Bind(wx.EVT_TOOL, lambda e: on_delete_sheet(self, e), delete_tool)
            self.Bind(wx.EVT_TOOL, lambda e: copy_sheet(self), copy_tool)
            self.Bind(wx.EVT_TOOL, lambda e: JoinSheetsWindow(self).Show(), join_tool)
            self.Bind(wx.EVT_TOOL, lambda evt: _show_rename_dialog(self), rename_tool)
            self.be_correction_spinbox.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_be_correction_change)
            self.Bind(wx.EVT_TOOL, self.on_auto_be, auto_be_tool)
            self.Bind(wx.EVT_TOOL, lambda e: self.on_open_background_window(), bkg_tool)
            self.Bind(wx.EVT_TOOL, lambda e: self.on_open_fitting_window(), peak_fit_tool)
            self.Bind(wx.EVT_TOOL, self.on_differentiate, diff_tool)
            self.Bind(wx.EVT_TOOL, self.open_periodic_table, id_tool)
            self.Bind(wx.EVT_TOOL, self.on_toggle_right_panel, self.toggle_right_panel_tool)
            self.sheet_combobox.Bind(wx.EVT_COMBOBOX, lambda e: on_sheet_selected(self, e))

        self.splitter.SetSashPosition(new_sash_position)
        self.toolbar.SetToolNormalBitmap(self.toggle_right_panel_tool.GetId(), new_bmp)

        # Ensure the splitter and its children are properly updated
        self.splitter.UpdateSize()
        self.right_frame.Layout()
        self.splitter.Refresh()
        self.canvas.draw()

    def on_splitter_changed(self, event):
        self.right_frame.Layout()
        self.canvas.draw()

    def on_window_resize(self, event):
        event.Skip()
        if hasattr(self, 'splitter') and self.splitter:
            self.splitter.UpdateSize()
            for panel in self.splitter.GetChildren():
                panel.Layout()
            self.splitter.Refresh()

    # I DON'T THINK IT IS USED
    def on_listbox_selection(self, event):
        self.selected_indices = self.file_listbox.GetSelections()
        update_sheet_names(self)


    def on_change_sheet_name(self, event):
        new_sheet_name = self.change_to_textctrl.GetValue()
        if new_sheet_name:
            rename_sheet(self, new_sheet_name)
            wx.MessageBox(f'Sheet renamed to "{new_sheet_name}"', "Info", wx.OK | wx.ICON_INFORMATION)









    def number_to_letter(n):
        return chr(65 + n)  # 65 is the ASCII value for 'A'


    def on_open_background_window(self):
        if not hasattr(self, 'background_window') or not self.background_window:
            self.background_window = BackgroundWindow(self)

            # Set position relative to main window
            main_pos = self.GetPosition()
            main_size = self.GetSize()
            bg_size = self.background_window.GetSize()

            x = main_pos.x + (main_size.width - bg_size.width) // 2
            y = main_pos.y + (main_size.height - bg_size.height) // 2

            self.background_window.SetPosition((x, y))

            self.background_tab_selected = True
            self.background_window.Bind(wx.EVT_CLOSE, self.on_background_window_close)
        self.background_window.Show()
        self.background_window.Raise()

    def on_background_window_close(self, event):
        save_state(self)
        self.background_tab_selected = False
        self.show_hide_vlines()
        self.background_window = None
        event.Skip()

    def enable_background_interaction(self):
        self.background_tab_selected = True
        self.show_hide_vlines()

    def disable_background_interaction(self):
        self.vline1 = None
        self.vline2 = None
        self.vline3 = None
        self.vline4 = None
        self.background_tab_selected = False
        self.show_hide_vlines()

    def on_open_fitting_window(self):
        save_state(self)
        if self.fitting_window is None or not self.fitting_window:
            self.fitting_window = FittingWindow(self)
            self.background_tab_selected = True
            self.peak_fitting_tab_selected = False

            # Set the position of the fitting window relative to the main window
            main_pos = self.GetPosition()
            main_size = self.GetSize()
            fitting_size = self.fitting_window.GetSize()

            # Calculate the position to center the fitting window on the main window
            x = main_pos.x + (main_size.width - fitting_size.width) // 2
            y = main_pos.y + (main_size.height - fitting_size.height) // 2

            self.fitting_window.SetPosition((x, y))

            self.show_hide_vlines()
            self.peak_manipulation.deselect_all_peaks()

        self.fitting_window.Show()
        self.fitting_window.Raise()  # Bring the window to the front

    def on_open_noise_analysis_window(self, event):
        if self.noise_analysis_window is None or not self.noise_analysis_window:
            self.noise_analysis_window = NoiseAnalysisWindow(self)
            self.noise_tab_selected = True
            self.show_hide_vlines()

        # Get the position and size of the main window
        main_pos = self.GetPosition()
        main_size = self.GetSize()

        # Get the size of the noise analysis windo
        noise_size = self.noise_analysis_window.GetSize()

        # Calculate the position to center the noise analysis window on the main windo
        x = main_pos.x + (main_size.width - noise_size.width) // 2
        y = main_pos.y + (main_size.height - noise_size.height) // 2

        # Set the position of the noise analysis window
        self.noise_analysis_window.SetPosition((x, y))

        self.noise_analysis_window.Show()
        self.noise_analysis_window.Raise()

        # Ensure the noise window stays on top
        self.noise_analysis_window.SetWindowStyle(self.noise_analysis_window.GetWindowStyle() | wx.STAY_ON_TOP)

    def noise_window_closed(self):
        self.noise_tab_selected = False
        self.show_hide_vlines()

    def clear_and_replot(self):
        self.plot_manager.clear_and_replot(self)

    def plot_data(self):
        self.plot_manager.plot_data(self)

    def update_overall_fit_and_residuals(self):
        self.plot_manager.update_overall_fit_and_residuals(self)

    def update_peak_plot(self, x, y, remove_old_peaks=True):
        self.plot_manager.update_peak_plot(self, x, y, remove_old_peaks)

    def update_peak_fwhm_SECOND(self, x):
        self.plot_manager.update_peak_fwhm(self, x)


    def adjust_plot_limits(self, axis, direction):
        self.plot_config.adjust_plot_limits(self, axis, direction)

    def update_constraint(self, event):
        row = event.GetRow()
        col = event.GetCol()
        if row % 2 == 1:  # If it's a constraint row
            value = self.peak_params_grid.GetCellValue(row, col).lower()
            if value == 'f':
                self.peak_params_grid.SetCellValue(row, col, 'fixed')
        event.Skip()

    def on_grid_select(self, event):
        if self.peak_fitting_tab_selected:
            row = event.GetRow()
            if row % 2 == 0:  # Assuming peak parameters are in even rows and constraints in odd rows
                # Removing any cross that were there
                self.selected_peak_index = None
                self.clear_and_replot()

                peak_index = row // 2
                self.selected_peak_index = peak_index

                self.peak_manipulation.remove_cross_from_peak()
                self.peak_manipulation.highlight_selected_peak()  # Highlight the selected peak
            else:
                self.selected_peak_index = None
                self.peak_manipulation.deselect_all_peaks()
        else:
            # If peak fitting tab is not selected, don't allow peak selection
            self.peak_manipulation.selected_peak_index = None
            self.peak_manipulation.deselect_all_peaks()

        self.update_checkboxes_from_data()
        event.Skip()




    def get_linked_peaks(self, peak_index):
        linked_peaks = []
        row = peak_index * 2
        for i in range(self.peak_params_grid.GetNumberRows() // 2):
            constraint_row = i * 2 + 1
            position_constraint = self.peak_params_grid.GetCellValue(constraint_row, 2)
            if position_constraint.startswith(chr(65 + peak_index)):
                linked_peaks.append(i)
        return linked_peaks

    def update_linked_peak(self, peak_index, new_x, new_height, area=None, original_peak_index=None):
        row = peak_index * 2
        constraint_row = row + 1
        position_constraint = self.peak_params_grid.GetCellValue(constraint_row, 2)
        height_constraint = self.peak_params_grid.GetCellValue(constraint_row, 3)
        area_constraint = self.peak_params_grid.GetCellValue(constraint_row, 6)

        sheet_name = self.sheet_combobox.GetValue()
        peak_label = self.peak_params_grid.GetCellValue(row, 1)
        peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
        fitting_model = self.peak_params_grid.GetCellValue(row, 13)

        original_peak_letter = chr(65 + original_peak_index)

        # Update position if constrained (same as original)
        if position_constraint.startswith(original_peak_letter):
            if '+' in position_constraint:
                offset = float(position_constraint.split('+')[1].split('#')[0])
                new_position = new_x + offset
            elif '-' in position_constraint:
                offset = float(position_constraint.split('-')[1].split('#')[0])
                new_position = new_x - offset
            elif '*' in position_constraint:
                factor = float(position_constraint.split('*')[1].split('#')[0])
                new_position = new_x * factor
            elif '/' in position_constraint:
                factor = float(position_constraint.split('/')[1].split('#')[0])
                new_position = new_x / factor
            else:
                new_position = new_x

            self.peak_params_grid.SetCellValue(row, 2, f"{new_position:.2f}")
            if peak_label in peaks:
                peaks[peak_label]['Position'] = new_position

        if (("LA" in fitting_model or "GL (Area)" in fitting_model or "Voigt" in fitting_model or "ExpGauss" in
            fitting_model) or "DS" in fitting_model) and area_constraint.startswith(
                original_peak_letter):
            current_area = float(self.peak_params_grid.GetCellValue(original_peak_index * 2, 6))
            if '*' in area_constraint:
                factor = float(area_constraint.split('*')[1].split('#')[0])
                new_linked_area = current_area * factor
            elif '/' in area_constraint:
                factor = float(area_constraint.split('/')[1].split('#')[0])
                new_linked_area = current_area / factor
            elif '+' in area_constraint:
                offset = float(area_constraint.split('+')[1].split('#')[0])
                new_linked_area = current_area + offset
            elif '-' in area_constraint:
                offset = float(area_constraint.split('-')[1].split('#')[0])
                new_linked_area = current_area - offset
            else:
                new_linked_area = current_area

            self.peak_params_grid.SetCellValue(row, 6, f"{new_linked_area:.2f}")

            # Recalculate height from area
            fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
            new_linked_height = self.calculate_height_from_area(new_linked_area, fwhm, fitting_model, row)
            self.peak_params_grid.SetCellValue(row, 3, f"{new_linked_height:.2f}")

            if peak_label in peaks:
                peaks[peak_label]['Area'] = new_linked_area
                peaks[peak_label]['Height'] = new_linked_height
        elif height_constraint.startswith(original_peak_letter):
            # Check if model uses height as primary parameter
            height_based_models = ["GL (Height)", "SGL (Height)", "D-parameter"]

            if fitting_model in height_based_models:
                # Only update height for height-based models
                if '*' in height_constraint:
                    factor = float(height_constraint.split('*')[1].split('#')[0])
                    new_linked_height = new_height * factor
                elif '/' in height_constraint:
                    factor = float(height_constraint.split('/')[1].split('#')[0])
                    new_linked_height = new_height / factor
                elif '+' in height_constraint:
                    offset = float(height_constraint.split('+')[1].split('#')[0])
                    new_linked_height = new_height + offset
                elif '-' in height_constraint:
                    offset = float(height_constraint.split('-')[1].split('#')[0])
                    new_linked_height = new_height - offset
                else:
                    new_linked_height = new_height

                self.peak_params_grid.SetCellValue(row, 3, f"{new_linked_height:.2f}")
                if peak_label in peaks:
                    peaks[peak_label]['Height'] = new_linked_height


        # Recalculate area if not LA model
        # if not "LA" in fitting_model:
        if not ("LA" in fitting_model or "GL (Area)" in fitting_model or "Voigt" in fitting_model or "ExpGauss" in fitting_model) and area_constraint.startswith(original_peak_letter):
            self.recalculate_peak_area(peak_index)




    def calculate_height_from_area(self, area, fwhm, model, row=None):
        if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
            # For Voigt, this is an approximation
            return area / (fwhm * np.sqrt(np.pi / (4 * np.log(2))))
        elif model in ["Voigt (Area, L/G, \u03c3, S)"]:
            if row is None:
                raise ValueError("Row must be provided for Skewed Voigt model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7)) / 2.355
            gamma = float(self.peak_params_grid.GetCellValue(row, 8)) / 2
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            height = PeakFunctions.get_skewedvoigt_height(area, sigma, gamma, skew)
            return height
        elif model == "DS*G (A, \u03c3, \u03b3, S)":
            if row is None:
                raise ValueError("Row must be provided for DS*G model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            # Calculate height numerically for DS*G model
            x_range = np.linspace(center - 5 * fwhm, center + 5 * fwhm, 1000)
            y_values = PeakFunctions.DS_G(x_range, center, area, gamma, skew, sigma)
            height = np.max(y_values)
            return height
        elif model == "DS (A, \u03c3, \u03b3)":
            if row is None:
                raise ValueError("Row must be provided for DS model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            # Create DS model instance
            model = lmfit.models.DoniachModel()

            # Calculate amplitude from area for DS model
            amplitude = PeakFunctions.doniach_sunjic_area_to_amplitude(area, sigma, gamma, skew)

            # Calculate height numerically for DS model
            x_range = np.linspace(center - 5 * fwhm, center + 5 * fwhm, 1000)
            y_values = model.eval(x=x_range, amplitude=amplitude, center=center,
                                  sigma=sigma, gamma=gamma, asymmetry=skew)
            height = np.max(y_values)
            return height
        elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
            if row is None:
                raise ValueError("Row must be provided for ExpGauss model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))

            # Create model instance first
            exp_gauss_model = lmfit.models.ExponentialGaussianModel()

            # Then evaluate with parameters
            x_range = np.linspace(center - 10 * sigma, center + 10 * sigma, 1000)
            y_values = exp_gauss_model.eval(x=x_range, amplitude=area, center=center, sigma=sigma, gamma=gamma)
            height = np.max(y_values)
            return height

        elif model == "Pseudo-Voigt (Area)":
            # For Pseudo-Voigt, use the linked peak's parameters
            if row is None:
                return area / (fwhm * np.pi / 2)  # Default approximation
            sigma = fwhm / 2
            fraction = float(self.peak_params_grid.GetCellValue(row, 5)) / 100  # Get L/G ratio of linked peak

            # Calculate proper height using pseudo-voigt formula with correct parameters
            return PeakFunctions.get_pseudo_voigt_height(area, sigma, fraction)

        elif model in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
            if row is None:
                raise ValueError("Row must be provided for LA model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))

            # Calculate height numerically
            x_range = np.linspace(center - 5 * fwhm, center + 5 * fwhm, 1000)
            y_values = PeakFunctions.LA(x_range, center, area, fwhm, sigma, gamma)
            height = np.max(y_values)
            return height
        elif model in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            if row is None:
                raise ValueError("Row must be provided for LA model")
            center = float(self.peak_params_grid.GetCellValue(row, 2))
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            # Calculate height numerically
            x_range = np.linspace(center - 5 * fwhm, center + 5 * fwhm, 1000)
            y_values = PeakFunctions.LAxG(x_range, center, area, fwhm, sigma, gamma, skew)
            height = np.max(y_values)
            return height

        elif model in ["GL (Area)", "GL (Height)", "SGL (Height)"]:
            return area / (fwhm * np.sqrt(np.pi / (4 * np.log(2))))
        elif model in ["SGL (Area)"]:
            if row is None:
                raise ValueError("Row must be provided for SGL model")
            fraction = float(self.peak_params_grid.GetCellValue(row, 5))
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = fwhm / 2
            return area / ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (fraction / 100) * np.pi * gamma)
        elif model == "D-parameter":
            # D-parameter doesn't have an area
            return 0.0

        else:  # GL, SGL, or other models
            return area / (fwhm * np.sqrt(np.pi / (4 * np.log(2))))




    def update_linked_peak_fwhm(self, peak_index, new_fwhm):
        row = peak_index * 2
        constraint_row = row + 1
        model = self.peak_params_grid.GetCellValue(row, 13)

        new_sigma = None
        new_gamma = None
        new_linked_fwhm = None

        if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
            sigma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 7)
            current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            lg_ratio = float(self.peak_params_grid.GetCellValue(row, 5))

            if '*' in sigma_constraint:
                factor = float(sigma_constraint.split('*')[1].split('#')[0])
                new_sigma = current_sigma * factor
                new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
        elif model in ["Voigt (Area, L/G, \u03c3, S)"]:
            sigma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 7)
            current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            lg_ratio = float(self.peak_params_grid.GetCellValue(row, 5))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            if '*' in sigma_constraint:
                factor = float(sigma_constraint.split('*')[1].split('#')[0])
                new_sigma = current_sigma * factor
                new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
        elif model in ["DS (A, \u03c3, \u03b3)"]:
            sigma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 7)
            current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))

            if '*' in sigma_constraint:
                factor = float(sigma_constraint.split('*')[1].split('#')[0])
                new_sigma = current_sigma * factor
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")
                # No need to update gamma as it's independent in DS model
        elif model == "DS*G (A, \u03c3, \u03b3, S)":
            # For DS*G model, get constraints from the grid
            sigma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 7)
            gamma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 8)
            skew_constraint = self.peak_params_grid.GetCellValue(constraint_row, 9)

            current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            current_skew = float(self.peak_params_grid.GetCellValue(row, 9))

            # Handle sigma constraint
            if '*' in sigma_constraint:
                factor = float(sigma_constraint.split('*')[1].split('#')[0])
                new_sigma = current_sigma * factor
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")

            # Handle gamma constraint
            if '*' in gamma_constraint:
                factor = float(gamma_constraint.split('*')[1].split('#')[0])
                new_gamma = current_gamma * factor
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")

            # Handle skew constraint
            if '*' in skew_constraint:
                factor = float(skew_constraint.split('*')[1].split('#')[0])
                new_skew = current_skew * factor
                self.peak_params_grid.SetCellValue(row, 9, f"{new_skew:.2f}")
        elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
            gamma_constraint = self.peak_params_grid.GetCellValue(constraint_row, 8)
            current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))

            if '*' in gamma_constraint:
                factor = float(gamma_constraint.split('*')[1].split('#')[0])
                new_gamma = current_gamma * factor
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")

        else:
            fwhm_constraint = self.peak_params_grid.GetCellValue(constraint_row, 4)
            if '*' in fwhm_constraint:
                factor = float(fwhm_constraint.split('*')[1].split('#')[0])
                new_linked_fwhm = new_fwhm * factor
                self.peak_params_grid.SetCellValue(row, 4, f"{new_linked_fwhm:.2f}")

        self.recalculate_peak_area(peak_index)

        sheet_name = self.sheet_combobox.GetValue()
        peak_label = self.peak_params_grid.GetCellValue(row, 1)
        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            if peak_label in peaks:
                if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"] and new_sigma is not None:
                    peaks[peak_label]['Sigma'] = new_sigma
                    peaks[peak_label]['Gamma'] = new_gamma
                elif model in ["Voigt (Area, L/G, \u03c3, S)", "DS (A, \u03c3, \u03b3)", "DS*G (A, \u03c3, "
                                                                                         "\u03b3, S)"] and new_sigma is not None:
                    peaks[peak_label]['Sigma'] = new_sigma
                    peaks[peak_label]['Gamma'] = new_gamma
                elif model == "ExpGauss.(Area, \u03c3, \u03b3)" and new_gamma is not None:
                    peaks[peak_label]['Gamma'] = new_gamma
                elif new_linked_fwhm is not None:
                    peaks[peak_label]['FWHM'] = new_linked_fwhm

    def update_linked_peaks_recursive(self, original_peak_index, new_x, new_height, area=None, visited=None):
        if visited is None:
            visited = set()

        if original_peak_index in visited:
            return

        visited.add(original_peak_index)

        linked_peaks = self.get_linked_peaks(original_peak_index)
        for linked_peak in linked_peaks:
            if linked_peak not in visited:
                if area is not None:
                    self.update_linked_peak(linked_peak, new_x, new_height, area, original_peak_index)
                else:
                    self.update_linked_peak(linked_peak, new_x, new_height, None, original_peak_index)

                linked_x = float(self.peak_params_grid.GetCellValue(linked_peak * 2, 2))
                linked_height = float(self.peak_params_grid.GetCellValue(linked_peak * 2, 3))
                linked_area = float(
                    self.peak_params_grid.GetCellValue(linked_peak * 2, 6)) if area is not None else None

                self.update_linked_peaks_recursive(linked_peak, linked_x, linked_height, linked_area, visited)

    def update_linked_fwhm_recursive(self, peak_index, new_fwhm, visited=None):
        if visited is None:
            visited = set()

        if peak_index in visited:
            return

        visited.add(peak_index)

        linked_peaks = self.get_linked_peaks(peak_index)
        for linked_peak in linked_peaks:
            if linked_peak not in visited:
                row = peak_index * 2
                model = self.peak_params_grid.GetCellValue(row, 13)

                if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, L/G, \u03c3, S)", "DS (A, \u03c3, \u03b3)", "DS*G (A, \u03c3, \u03b3, S)"]:
                    original_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                    original_lg = float(self.peak_params_grid.GetCellValue(row, 5))
                    linked_sigma = original_sigma
                    linked_gamma = (original_lg / 100 * linked_sigma) / (1 - original_lg / 100)
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 7, f"{linked_sigma:.2f}")
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 8, f"{linked_gamma:.2f}")
                    self.update_linked_fwhm_recursive(linked_peak, new_fwhm, visited)
                elif model in ["Voigt (Area, \u03c3, \u03b3)", "DS*G (A, \u03c3, \u03b3, S)"]:
                    original_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                    original_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                    linked_gamma = original_gamma
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 7, f"{original_sigma:.3f}")
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 8, f"{linked_gamma:.3f}")
                    self.update_linked_fwhm_recursive(linked_peak, new_fwhm, visited)
                elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
                    original_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                    original_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                    linked_gamma = original_gamma
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 7, f"{original_sigma:.2f}")
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 8, f"{linked_gamma:.2f}")
                    self.update_linked_fwhm_recursive(linked_peak, new_fwhm, visited)
                else:
                    linked_fwhm = float(self.peak_params_grid.GetCellValue(linked_peak * 2, 4))
                    self.update_linked_peak_fwhm(linked_peak, new_fwhm)
                    self.update_linked_fwhm_recursive(linked_peak, linked_fwhm, visited)

    def update_peak_fwhm(self, x):
        if self.initial_fwhm is not None and self.initial_x is not None:
            row = self.selected_peak_index * 2
            peak_label = self.peak_params_grid.GetCellValue(row, 1)
            model = self.peak_params_grid.GetCellValue(row, 13)
            delta_x = x - self.initial_x

            if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
                current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                lg_ratio = float(self.peak_params_grid.GetCellValue(row, 5))
                new_sigma = max(current_sigma + delta_x * 1, 0.4)
                new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
                new_fwhm = self.initial_fwhm
            elif model in ["DS (A, \u03c3, \u03b3)"]:
                current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                skew = float(self.peak_params_grid.GetCellValue(row, 9))

                new_sigma = max(current_sigma + delta_x * 1, 0.4)
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")
                # No change to gamma as it's independent in DS model
                new_fwhm = self.initial_fwhm
            elif model == "DS*G (A, \u03c3, \u03b3, S)":
                # Get current parameter values
                current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                current_skew = float(self.peak_params_grid.GetCellValue(row, 9))

                # For DS*G, adjust the Gaussian width with drag
                new_sigma = max(current_sigma + delta_x * 0.5, 0.2)
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.3f}")

                # We can also slightly adjust gamma if desired, but keep skew unchanged
                new_gamma = max(current_gamma + delta_x * 0.3, 0.1)
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")

                # Keep original FWHM reference
                new_fwhm = self.initial_fwhm
            elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
                current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                # new_sigma = max(current_sigma + delta_x * 0.5, 0.2)
                new_gamma = max(current_gamma + delta_x * 0.5, 0.2)
                self.peak_params_grid.SetCellValue(row, 7, f"{current_sigma:.3f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.3f}")
                new_fwhm = self.initial_fwhm
            else:
                new_fwhm = max(self.initial_fwhm + delta_x * 1, 0.3)
                self.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.3f}")

            # Update FWHM in window.Data
            sheet_name = self.sheet_combobox.GetValue()
            if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][
                sheet_name] and 'Peaks' in self.Data['Core levels'][sheet_name]['Fitting']:
                peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                if peak_label in peaks:
                    peaks[peak_label]['FWHM'] = new_fwhm
                    if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                 "ExpGauss.(Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)", "DS (A, \u03c3, "
                                "\u03b3)", "DS*G (A, \u03c3, \u03b3, S)"]:
                        peaks[peak_label]['Sigma'] = new_sigma
                        peaks[peak_label]['Gamma'] = new_gamma

            # Recalculate area
            self.recalculate_peak_area(self.selected_peak_index)

            # Redraw everything with updated peak info
            self.clear_and_replot()

            # Add this line to update the displayed FWHM
            self.peak_manipulation.highlight_selected_peak()

            return new_fwhm

        return None

    def recalculate_peak_area(self, peak_index):
        row = peak_index * 2
        sheet_name = self.sheet_combobox.GetValue()
        peak_label = self.peak_params_grid.GetCellValue(row, 1)

        height = float(self.peak_params_grid.GetCellValue(row, 3))
        fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
        fraction = float(self.peak_params_grid.GetCellValue(row, 5))
        model = self.peak_params_grid.GetCellValue(row, 13)

        if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)", "ExpGauss.(Area, \u03c3, \u03b3)",
                     "LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)", "LA*G (Area, \u03c3/\u03b3, "
                        "\u03b3)", "Voigt (Area, L/G, \u03c3, S)", "DS (A, \u03c3, \u03b3)", "DS*G (A, \u03c3, "
                                                                                             "\u03b3, S)"]:
            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
            gamma = float(self.peak_params_grid.GetCellValue(row, 8))
            skew = float(self.peak_params_grid.GetCellValue(row, 9))
            area = self.calculate_peak_area(model, height, fwhm, fraction, sigma, gamma, skew)
        elif model == "D-parameter":
            return
        else:
            area = self.calculate_peak_area(model, height, fwhm, fraction)

        self.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")

        # Update area in self.Data
        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            if peak_label in self.Data['Core levels'][sheet_name]['Fitting']['Peaks']:
                self.Data['Core levels'][sheet_name]['Fitting']['Peaks'][peak_label]['Area'] = area

        return area

    def show_hide_vlines(self):
        # Hide vlines if zooming or dragging
        if self.zoom_mode or self.drag_mode:
            if self.vline1 is not None:
                self.vline1.set_visible(False)
            if self.vline2 is not None:
                self.vline2.set_visible(False)
            if self.vline3 is not None:
                self.vline3.set_visible(False)
            if self.vline4 is not None:
                self.vline4.set_visible(False)
            return

        # Existing visibility logic
        background_lines_visible = hasattr(self, 'fitting_window') and self.background_tab_selected
        noise_lines_visible = self.noise_analysis_window is not None and self.noise_tab_selected

        if self.vline1 is not None:
            self.vline1.set_visible(background_lines_visible)
        if self.vline2 is not None:
            self.vline2.set_visible(background_lines_visible)
        if self.vline3 is not None:
            self.vline3.set_visible(noise_lines_visible)
        if self.vline4 is not None:
            self.vline4.set_visible(noise_lines_visible)

        self.canvas.draw_idle()


    def is_mouse_on_peak(self, event):
        if self.selected_peak_index is not None:

            return True
        return False













    def show_popup_message(self, message):
        import platform
        import time

        try:
            # Check if platform is macOS
            if platform.system() == 'Darwin':  # 'Darwin' is the system name for macOS
                current_time = time.time()
                # Only show popup if 5+ seconds have passed since last one
                if not hasattr(self, 'last_popup_time') or (current_time - self.last_popup_time > 5):
                    popup = wx.adv.RichToolTip("Are you trying to select a peak?", message)
                    popup.ShowFor(self)
                    self.last_popup_time = current_time
            else:
                # For non-Mac platforms, just show the popup
                popup = wx.adv.RichToolTip("Are you trying to select a peak?", message)
                popup.ShowFor(self)
        except wx.PyDeadObjectError:
            # Object already destroyed, ignore
            pass
        except Exception as e:
            print(f"Error showing popup message: {e}")

    def show_popup_message2(self, message1, message2):
        import platform
        import time

        try:
            # Check if platform is macOS
            if platform.system() == 'Darwin':
                current_time = time.time()
                if not hasattr(self, 'last_popup_time') or (current_time - self.last_popup_time > 5):
                    popup = wx.adv.RichToolTip(message1, message2)
                    popup.ShowFor(self)
                    self.last_popup_time = current_time
            else:
                popup = wx.adv.RichToolTip(message1, message2)
                popup.ShowFor(self)
        except wx.PyDeadObjectError:
            # Object already destroyed, ignore
            pass
        except Exception as e:
            print(f"Error showing popup message: {e}")





    def on_vline_drag(self, event):
        if event.inaxes and self.active_vline is not None:
            self.active_vline.set_xdata([event.xdata, event.xdata])
            self.canvas.draw_idle()

    def on_vline_release(self, event):
        if self.active_vline is not None:
            self.active_vline = None
            self.canvas.mpl_disconnect(self.motion_cid)
            self.canvas.mpl_disconnect(self.release_cid)

            # Update background if in Adaptive Smart mode
            if self.background_method == "Multi-Regions Smart":
                self.plot_manager.plot_background(self)


    def on_zoom_in_tool(self, event):
        self.plot_config.on_zoom_in_tool(self)

    def on_zoom_select(self, eclick, erelease):
        self.plot_config.on_zoom_select(self, eclick, erelease)

        ymin, ymax = self.ax.get_ylim()
        if hasattr(self, 'rsd_text') and self.rsd_text:
            residual_height = 1.07 * max(self.y_values)
            if residual_height > ymax:
                self.rsd_text.remove()
                self.rsd_text = None

    def on_zoom_out(self, event):
        self.plot_config.on_zoom_out(self)

    def on_drag_tool(self, event):
        self.plot_config.on_drag_tool(self)

    def enable_drag(self):
        self.plot_config.enable_drag(self)

    def disable_drag(self):
        self.plot_config.disable_drag(self)

    def on_drag_release(self, event):
        self.plot_config.on_drag_release(self, event)



    def update_checkboxes_from_data(self):
        sheet_name = self.sheet_combobox.GetValue()
        match = re.search(r'(\d+)$', sheet_name)
        row_number = match.group(1) if match else "0"
        results_table_key = f'Results Table{row_number}'

        if results_table_key in self.Data and 'Peak' in self.Data[results_table_key]:
            from libraries.Grid_Operations import CheckboxRenderer

            for row in range(self.results_grid.GetNumberRows()):
                peak_label = f"Peak_{row}"
                if peak_label in self.Data[results_table_key]['Peak']:
                    checkbox_state = self.Data[results_table_key]['Peak'][peak_label].get('Checkbox', '0')
                    current_grid_state = self.results_grid.GetCellValue(row, 7)

                    # Only update if the value has changed
                    if checkbox_state != current_grid_state:
                        self.results_grid.SetCellValue(row, 7, checkbox_state)

                    # Always ensure the renderer is correct
                    self.results_grid.SetCellRenderer(row, 7, CheckboxRenderer())
                    self.results_grid.SetReadOnly(row, 7)
                    self.results_grid.RefreshAttr(row, 7)

        # Final refresh
        self.results_grid.ForceRefresh()

    def on_plot_mouse_release(self, event):
        self.update_checkboxes_from_data()
        # No need to call event.Skip() for Matplotlib events

    def on_peak_params_mouse_release(self, event):
        self.update_checkboxes_from_data()
        event.Skip()

    def on_peak_params_cell_select(self, event):
        self.update_checkboxes_from_data()
        event.Skip()


    def export_results(self):
        save_state(self)
        export_results(self)


    def on_cell_changed(self, event):
        return


    # In MyFrame class
    def update_atomic_percentages(self):
        # Get current sheet's row number
        sheet_name = self.sheet_combobox.GetValue()
        row_number = 0

        import re
        match = re.search(r'(\d+)$', sheet_name)
        if match:
            row_number = int(match.group(1))

        results_table_key = f'Results Table{row_number}'

        # Ensure the results table exists
        if results_table_key not in self.Data:
            self.Data[results_table_key] = {'Peak': {}}

        current_rows = self.results_grid.GetNumberRows()
        total_normalized_area = 0
        checked_indices = []

        # Calculate ECF for each peak
        for i in range(current_rows):
            if self.results_grid.GetCellValue(i, 7) == '1':  # If checkbox ticked
                # Get values from grid
                peak_name = self.results_grid.GetCellValue(i, 0)
                binding_energy = float(self.results_grid.GetCellValue(i, 1))
                area = float(self.results_grid.GetCellValue(i, 5))
                rsf = float(self.results_grid.GetCellValue(i, 8))

                # Calculate kinetic energy
                kinetic_energy = self.photons - binding_energy

                # Calculate ECF based on method selected
                if self.library_type == "Scofield":
                    ecf = kinetic_energy ** 0.6
                    self.results_grid.SetCellValue(i, 10, "KE^0.6")
                elif self.library_type == "Wagner":
                    ecf = kinetic_energy ** 1.0
                    self.results_grid.SetCellValue(i, 10, "KE^1.0")
                elif self.library_type == "TPP-2M":
                    # Calculate IMFP using TPP-2M
                    imfp = AtomicConcentrations.calculate_imfp_tpp2m(kinetic_energy)
                    ecf = imfp * 26.2
                    self.results_grid.SetCellValue(i, 10, "TPP-2M")
                elif self.library_type == "EAL":
                    z_avg = 50  # Default value
                    eal = (0.65 + 0.007 * kinetic_energy ** 0.93) / (z_avg ** 0.38)
                    ecf = eal
                    self.results_grid.SetCellValue(i, 10, "EAL")
                elif self.library_type == "None":
                    ecf = 1.0
                    self.results_grid.SetCellValue(i, 10, "None: 1.0")
                else:
                    ecf = 1.0  # Default no correction
                    self.results_grid.SetCellValue(i, 10, "None: 1.0")


                # Calculate Transmission function
                txfn = 1.0  # Transmission function

                # Angular correction
                angular_correction = 1.0
                if self.use_angular_correction:
                    angular_correction = AtomicConcentrations.calculate_angular_correction(
                        self, peak_name, self.analysis_angle
                    )

                # Calculate normalized area with ECF correction
                normalized_area = area / (rsf * txfn * ecf * angular_correction)

                total_normalized_area += normalized_area
                checked_indices.append((i, normalized_area))

                # Update the data in the correct Results Table
                peak_key = f"Peak_{i}"
                if peak_key not in self.Data[results_table_key]['Peak']:
                    self.Data[results_table_key]['Peak'][peak_key] = {}

                # Update ECF, TXFN values in the data structure
                self.Data[results_table_key]['Peak'][peak_key].update({
                    # 'ECF': ecf,
                    'TXFN': txfn,
                    'RSF': rsf,
                    'Name': peak_name,
                    'Position': binding_energy,
                    'Area': area,
                    'Checkbox': '1'
                })
            else:
                # Set the atomic percentage to 0 for unticked rows
                self.results_grid.SetCellValue(i, 6, "0.00")

                # Update in data structure if it exists
                peak_key = f"Peak_{i}"
                if results_table_key in self.Data and 'Peak' in self.Data[results_table_key] and peak_key in \
                        self.Data[results_table_key]['Peak']:
                    self.Data[results_table_key]['Peak'][peak_key]['at. %'] = 0.00
                    self.Data[results_table_key]['Peak'][peak_key]['Checkbox'] = '0'

        # Calculate atomic percentages
        for i, norm_area in checked_indices:
            atomic_percent = (norm_area / total_normalized_area) * 100 if total_normalized_area > 0 else 0
            self.results_grid.SetCellValue(i, 6, f"{atomic_percent:.2f}")

            # Update in data structure
            peak_key = f"Peak_{i}"
            if peak_key in self.Data[results_table_key]['Peak']:
                self.Data[results_table_key]['Peak'][peak_key]['at. %'] = atomic_percent

        # Force a refresh to update the display
        self.results_grid.ForceRefresh()

    def on_height_changed(self, event):
        row = event.GetRow()
        col = event.GetCol()

        if col == 2:  # Check if the height column was edited
            try:
                # Get the new height value
                new_height = float(self.results_grid.GetCellValue(row, col))
                # Get the corresponding FWHM value
                fwhm = float(self.results_grid.GetCellValue(row, 3))
                # Get the corresponding RSF value
                rsf = float(self.results_grid.GetCellValue(row, 8))

                # Recalculate the area
                new_area = new_height * fwhm * (np.sqrt(2 * np.pi) / 2.355)
                self.results_grid.SetCellValue(row, 5, f"{new_area:.2f}")

                # Recalculate the relative area
                new_rel_area = new_area / rsf
                self.results_grid.SetCellValue(row, 10, f"{new_rel_area:.2f}")

                # Update the atomic percentages if necessary
                self.update_atomic_percentages()

            except ValueError:
                wx.MessageBox("Invalid height value", "Error", wx.OK | wx.ICON_ERROR)

    def calculate_peak_area(self, model, height, fwhm, fraction, sigma=None, gamma=None, skew=None):
        if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:#, "Voigt (Area, L/G, \u03c3, S)"]:
            if sigma is None or gamma is None:
                raise ValueError("Sigma and gamma are required for Voigt models")
            area = PeakFunctions.voigt_height_to_area(height, sigma / 2.355, gamma / 2)
        elif model == "Voigt (Area, L/G, \u03c3, S)":
            # Set default values if parameters are missing
            sigma = sigma or 0.5
            gamma = gamma or 0.5
            skew = skew or 0.1
            if sigma is None or gamma is None:
                raise ValueError("Sigma and gamma are required for Voigt models")
            area = PeakFunctions.skewedvoigt_height_to_area(height, sigma / 2.355, gamma / 2, skew)
        elif model == "DS (A, \u03c3, \u03b3)":
            # Set default values if parameters are missing
            sigma = sigma or 1.0
            gamma = gamma or 0.0
            skew = skew or 0.0
            if sigma is None or gamma is None:
                raise ValueError("Sigma and gamma are required for DS models")
            # height_test = PeakFunctions.get_doniach_sunjic_height(area, sigma,gamma,skew)
            area = PeakFunctions.doniach_sunjic_height_to_area(height, sigma, gamma, skew)
        elif model == "DS*G (A, \u03c3, \u03b3, S)":
            # Create x_range centered around 0 for area calculation
            x_range = np.linspace(-10 * fwhm, 10 * fwhm, 1000)
            # Use position=0 since we only need the shape
            y_values = PeakFunctions.DS_G(x_range, 0, 1.0, gamma, skew, sigma)
            max_height = np.max(y_values)
            area = height / max_height if max_height > 0 else 0
            return area
        elif model in ["Pseudo-Voigt (Area)"]:#, "SGL (Area)"]:
            sigma = fwhm / 2
            amplitude = height / PeakFunctions.get_pseudo_voigt_height(1, sigma, fraction)
            area = amplitude
        elif model in ["GL (Height)", "SGL (Height)", "Unfitted"]:
            area = height * fwhm * np.sqrt(np.pi / (4 * np.log(2)))
        elif model in ["GL (Area)"]: #, "SGL (Area)"]:
            area = height * fwhm * np.sqrt(np.pi / (4 * np.log(2)))
        elif model in ["SGL (Area)"]:
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = fwhm / 2
            area = height * ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (fraction / 100) * np.pi * gamma)
        elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
            area = height * sigma * np.sqrt(2 * np.pi) * np.exp(gamma ** 2 * sigma ** 2 / 4)
        elif model in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)"]:
            if sigma is None or gamma is None:
                raise ValueError("Sigma and gamma are required for LA model")

            x_range = np.linspace(-10 * fwhm, 10 * fwhm, 1000)
            y_temp = PeakFunctions.LA(x_range, 0, 1.0, fwhm, sigma, gamma)  # Use unit amplitude
            max_height = np.max(y_temp)
            y_values = PeakFunctions.LA(x_range, 0, height / max_height, fwhm, sigma, gamma)
            area = np.trapz(y_values, x_range)
            return round(area, 2)
        elif model in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            if sigma is None or gamma is None or skew is None:
                raise ValueError("Sigma, gamma and skew are required for LA*G model")

            x_range = np.linspace(-10 * fwhm, 10 * fwhm, 1000)
            y_temp = PeakFunctions.LAxG(x_range, 0, 1.0, fwhm, sigma, gamma, skew)  # Use unit amplitude
            max_height = np.max(y_temp)
            y_values = PeakFunctions.LAxG(x_range, 0, height / max_height, fwhm, sigma, gamma, skew)
            area = np.trapz(y_values, x_range)
            return round(area, 2)
        elif model =="D-parameter":
            return
        else:
            raise ValueError(f"Unknown fitting model: {model}")
        return round(area, 2)





    def update_ratios(self):
        num_peaks = self.peak_params_grid.GetNumberRows() // 2
        if num_peaks < 1:
            return

            # Calculate total area for concentrations
        total_area = 0
        for i in range(num_peaks):
            row = i * 2
            try:
                area = float(self.peak_params_grid.GetCellValue(row, 6))
                total_area += area
            except ValueError:
                continue

        # Get first peak area for A/Aa ratio calculation
        try:
            first_position = float(self.peak_params_grid.GetCellValue(0, 2))
            first_area = float(self.peak_params_grid.GetCellValue(0, 6))
        except ValueError:
            return

        for i in range(num_peaks):
            row = i * 2
            try:
                position = float(self.peak_params_grid.GetCellValue(row, 2))
                area = float(self.peak_params_grid.GetCellValue(row, 6))

                # Calculate concentration from area
                concentration = (area / total_area * 100) if total_area > 0 else 0

                # Calculate A/Aa ratio
                a_ratio = area / first_area if first_area != 0 else 0

                split = position - first_position

                # Update grid
                self.peak_params_grid.SetCellValue(row, 10, f"{concentration:.1f}")
                self.peak_params_grid.SetCellValue(row, 11, f"{a_ratio * 100:.1f}")
                self.peak_params_grid.SetCellValue(row, 12, f"{split:.2f}")
            except ValueError:
                continue

        self.peak_params_grid.ForceRefresh()



    # In MyFrame class
    def on_checkbox_update(self, event):
        row = event.GetRow()
        col = event.GetCol()

        if col == 7:
            # Toggle checkbox value
            current_value = self.results_grid.GetCellValue(row, col)
            new_value = '0' if current_value == '1' else '1'
            self.results_grid.SetCellValue(row, col, new_value)

            # Get current sheet's row number
            sheet_name = self.sheet_combobox.GetValue()
            row_number = 0

            import re
            match = re.search(r'(\d+)$', sheet_name)
            if match:
                row_number = int(match.group(1))

            results_table_key = f'Results Table{row_number}'

            # Update data structure
            peak_label = f"Peak_{row}"
            if results_table_key in self.Data and 'Peak' in self.Data[results_table_key] and peak_label in \
                    self.Data[results_table_key]['Peak']:
                self.Data[results_table_key]['Peak'][peak_label]['Checkbox'] = new_value

            # Update calculations and display
            self.update_atomic_percentages()
            self.clear_and_replot()
            self.canvas.draw_idle()
            save_state(self)

        event.Skip()

    def on_key_down(self, event):
        keycode = event.GetKeyCode()
        if keycode == wx.WXK_DELETE:
            selected_rows = self.get_selected_rows()
            if selected_rows:
                # Get the current sheet and extract its row number
                sheet_name = self.sheet_combobox.GetValue()
                row_number = "0"

                # Extract row number from sheet name
                import re
                match = re.search(r'(\d+)$', sheet_name)
                if match:
                    row_number = match.group(1)

                results_table_key = f'Results Table{row_number}'

                # Check if the results table exists
                if results_table_key not in self.Data:
                    print(f"Results table {results_table_key} not found")
                    event.Skip()
                    return

                # Sort rows in descending order to avoid index shifting
                selected_rows.sort(reverse=True)

                for row in selected_rows:
                    peak_name = self.results_grid.GetCellValue(row, 0)  # Get peak name
                    peak_label = f"Peak_{row}"

                    # Delete the row from the grid
                    self.results_grid.DeleteRows(row)

                    # Remove the peak from self.Data
                    if peak_label in self.Data[results_table_key]['Peak']:
                        del self.Data[results_table_key]['Peak'][peak_label]

                # Renumber the remaining peaks in self.Data
                new_data = {}
                for i, (key, value) in enumerate(self.Data[results_table_key]['Peak'].items()):
                    new_key = f"Peak_{i}"
                    new_data[new_key] = value

                self.Data[results_table_key]['Peak'] = new_data

                # Refresh and update
                self.results_grid.ForceRefresh()
                self.update_atomic_percentages()
                from libraries.Save import save_state
                save_state(self)
            else:
                event.Skip()
        else:
            event.Skip()

    def refresh_results_grid(self):
        self.results_grid.ClearGrid()
        for i, (peak_label, peak_data) in enumerate(self.Data['Results']['Peak'].items()):
            for j, value in enumerate(peak_data.values()):
                self.results_grid.SetCellValue(i, j, str(value))

    def get_selected_rows(self):
        """
        Get a list of selected rows in the grid.
        """
        selected_rows = []
        for row in range(self.results_grid.GetNumberRows()):
            if self.results_grid.IsInSelection(row, 0):  # Check if the row is selected
                selected_rows.append(row)
        return selected_rows




    def set_max_iterations(self, value):
        self.max_iterations = value

    def set_fitting_method(self, method):
        self.selected_fitting_method = method


    def set_background_method(self, method):
        self.background_method = method

    # Method to update Offset (H)
    def set_offset_h(self, value):
        try:
            self.offset_h = float(value)
        except ValueError:
            self.offset_h = 0

    def set_offset_l(self, value):
        try:
            self.offset_l = float(value)
        except ValueError:
            self.offset_l = 0


    def get_data_for_save(self):
        data = {
            'x_values': self.x_values,
            'y_values': self.y_values,
            'background': self.background if hasattr(self, 'background') else None,
            'calculated_fit': None,
            'residuals': None,
            'individual_peak_fits': [],
            'peak_params_grid': self.peak_params_grid if hasattr(self, 'peak_params_grid') else None,
            'results_grid': self.results_grid if hasattr(self, 'results_grid') else None
        }

        if self.peak_params_grid.GetNumberRows() == 0:
            return data

        row = 0  # First peak row
        grid_fitting_method = self.peak_params_grid.GetCellValue(row, 13)  # Column 13 contains fitting method
        sheet_name = self.sheet_combobox.GetValue()

        # Skip fit_peaks for D-parameter model
        if grid_fitting_method not in ["D-parameter", "Unfitted"]:
            if hasattr(self, 'peak_params_grid') and self.peak_params_grid.GetNumberRows() > 0:
                from Functions import fit_peaks
                fit_result = fit_peaks(self, self.peak_params_grid, evaluate=True)

                if fit_result:
                    if hasattr(self, 'ax'):
                        for line in self.ax.lines:
                            if line.get_label() == 'Overall Fit':
                                data['calculated_fit'] = line.get_ydata()
                            elif line.get_label() == 'Residuals':
                                data['residuals'] = line.get_ydata()

                        # for collection in self.ax.collections:
                        #     if collection.get_label().startswith(sheet_name):
                        #         data['individual_peak_fits'].append(collection.get_paths()[0].vertices[:, 1])
                        peak_labels = []
                        for i in range(self.peak_params_grid.GetNumberRows() // 2):
                            peak_labels.append(self.peak_params_grid.GetCellValue(i * 2, 1))

                        for collection in self.ax.collections:
                            # Check if the collection label matches any peak label
                            for peak_label in peak_labels:
                                if collection.get_label() == peak_label:
                                    print(f"Found collection for peak: {peak_label} "
                                          f"max height: {max(collection.get_paths()[0].vertices[:, 1])}")
                                    data['individual_peak_fits'].append(collection.get_paths()[0].vertices[:, 1])
                                    break

        return data

    def load_config(self):
        if os.path.exists('config.json'):
            with open('config.json', 'r') as f:
                config = json.load(f)
                self.plot_style = config.get('plot_style', self.plot_style)
                self.scatter_size = config.get('scatter_size', self.scatter_size)
                self.line_width = config.get('line_width', self.line_width)
                self.line_alpha = config.get('line_alpha', self.line_alpha)
                self.scatter_color = config.get('scatter_color', self.scatter_color)
                self.line_color = config.get('line_color', self.line_color)
                self.scatter_marker = config.get('scatter_marker', self.scatter_marker)
                self.background_color = config.get('background_color', self.background_color)
                self.background_alpha = config.get('background_alpha', self.background_alpha)
                self.background_linestyle = config.get('background_linestyle', self.background_linestyle)
                self.envelope_color = config.get('envelope_color', self.envelope_color)
                self.envelope_alpha = config.get('envelope_alpha', self.envelope_alpha)
                self.envelope_linestyle = config.get('envelope_linestyle', self.envelope_linestyle)
                self.residual_color = config.get('residual_color', self.residual_color)
                self.residual_alpha = config.get('residual_alpha', self.residual_alpha)
                self.residual_linestyle = config.get('residual_linestyle', self.residual_linestyle)
                self.background_thickness = config.get('background_thickness', 1)
                self.envelope_thickness = config.get('envelope_thickness', 1)
                self.residual_thickness = config.get('residual_thickness', 1)
                self.raw_data_linestyle = config.get('raw_data_linestyle', self.raw_data_linestyle)
                self.peak_colors = config.get('peak_colors', self.peak_colors)
                self.peak_alpha = config.get('peak_alpha', self.peak_alpha)
                self.peak_line_style = config.get('peak_line_style', 'Same Color')
                self.peak_line_alpha = config.get('peak_line_alpha', 0.7)
                self.peak_line_thickness = config.get('peak_line_thickness', 1)
                self.peak_line_pattern = config.get('peak_line_pattern', '-')
                self.recent_files = config.get('recent_files', [])
                self.peak_fill_types = config.get('peak_fill_types', ["Solid Fill" for _ in range(15)])
                self.peak_hatch_patterns = config.get('peak_hatch_patterns', ["/", "\\", "|", "-", "+", "x", "o", "O", ".", "*"] * 2)
                self.hatch_density = config.get('hatch_density', 2)
                self.current_instrument = config.get('current_instrument', 'Al1486')
                self.plot_font = config.get('plot_font', 'Arial')
                self.axis_title_size = config.get('axis_title_size', 12)
                self.axis_number_size = config.get('axis_number_size', 10)
                self.x_sublines = config.get('x_sublines', 5)
                self.y_sublines = config.get('y_sublines', 5)
                self.legend_font_size = config.get('legend_font_size', 8)
                self.core_level_text_size = config.get('core_level_text_size', 15)
                self.label_font_size = config.get('label_font_size', 8)
                self.excel_width = config.get('excel_width', 5.2)
                self.excel_height = config.get('excel_height', 5.2)
                self.excel_dpi = config.get('excel_dpi', 100)
                self.survey_excel_width = config.get('survey_excel_width', 10)
                self.survey_excel_height = config.get('survey_excel_height', 7)
                self.survey_excel_dpi = config.get('survey_excel_dpi', 100)
                self.word_width = config.get('word_width', 5)
                self.word_height = config.get('word_height', 5)
                self.survey_word_width = config.get('survey_word_width', 10)
                self.survey_word_height = config.get('survey_word_height', 5)
                self.survey_word_dpi = config.get('survey_word_dpi', 200)
                self.word_dpi = config.get('word_dpi', 300)
                self.export_width = config.get('export_width', 8)
                self.export_height = config.get('export_height', 6)
                self.export_dpi = config.get('export_dpi', 300)
                self.library_type = config.get('library_type', 'TPP-2M')
                self.use_angular_correction = config.get('use_angular_correction', False)
                self.analysis_angle = config.get('analysis_angle', 54.7)
                self.ref_peak_name = config.get('ref_peak_name', 'C1s C-C')
                self.ref_peak_be = config.get('ref_peak_be', 284.8)
                self.photons = config.get('photons', 1486.67)
                self.times_opened = config.get('times_opened', 0)
                self.enable_auto_backup = config.get('enable_auto_backup', False)
                self.backup_interval = config.get('backup_interval', 30)
                self.legend_visible = config.get('legend_visible', 1)
                self.y_axis_state = config.get('y_axis_state', 0)
                self.residuals_state = config.get('residuals_state', 2)
                self.enable_quick_settings = config.get('enable_quick_settings', False)

                # Set registered flag
                self.registered = config.get('registered', False)

                # Only increment times_opened if registered
                if self.registered:
                    self.times_opened = config.get('times_opened', 0) + 1
                else:
                    self.times_opened = 0

        else:
            config = {}
            print("No config file found, using default values.")
            self.recent_files = []

            self.registered = False
            self.times_opened = 0

        return config

    def save_config(self):
        config = {
            'plot_style': self.plot_style,
            'scatter_size': self.scatter_size,
            'line_width': self.line_width,
            'line_alpha': self.line_alpha,
            'scatter_color': self.scatter_color,
            'line_color': self.line_color,
            'scatter_marker': self.scatter_marker,
            'background_color': self.background_color,
            'background_alpha': self.background_alpha,
            'background_linestyle': self.background_linestyle,
            'envelope_color': self.envelope_color,
            'envelope_alpha': self.envelope_alpha,
            'envelope_linestyle': self.envelope_linestyle,
            'residual_color': self.residual_color,
            'residual_alpha': self.residual_alpha,
            'residual_linestyle': self.residual_linestyle,
            'background_thickness': self.background_thickness,
            'envelope_thickness': self.envelope_thickness,
            'residual_thickness': self.residual_thickness,
            'raw_data_linestyle': self.raw_data_linestyle,
            'peak_colors': self.peak_colors,
            'peak_alpha': self.peak_alpha,
            'peak_line_style': self.peak_line_style,
            'peak_line_alpha': self.peak_line_alpha,
            'peak_line_thickness': self.peak_line_thickness,
            'peak_line_pattern': self.peak_line_pattern,
            'recent_files': self.recent_files,
            'peak_fill_types': self.peak_fill_types,
            'peak_hatch_patterns': self.peak_hatch_patterns,
            'hatch_density': self.hatch_density,
            'legend_visible': self.legend_visible,
            'y_axis_state': self.y_axis_state,
            'residuals_state': self.residuals_state,
            'current_instrument': self.current_instrument,
            'plot_font': self.plot_font,
            'axis_title_size': self.axis_title_size,
            'axis_number_size': self.axis_number_size,
            'x_sublines': self.x_sublines,
            'y_sublines': self.y_sublines,
            'legend_font_size': self.legend_font_size,
            'core_level_text_size': self.core_level_text_size,
            'label_font_size': self.label_font_size,

            # Instruments settings
            'library_type': self.library_type,
            'use_angular_correction': self.use_angular_correction,
            'analysis_angle': self.analysis_angle,
            'ref_peak_name': self.ref_peak_name,
            'ref_peak_be': self.ref_peak_be,
            'photons': self.photons,


            # Excel file settings
            'excel_width': self.excel_width,
            'excel_height': self.excel_height,
            'excel_dpi': self.excel_dpi,
            'survey_excel_width': self.survey_excel_width,
            'survey_excel_height': self.survey_excel_height,
            'survey_excel_dpi': self.survey_excel_dpi,

            # Word report settings
            'word_width': self.word_width,
            'word_height': self.word_height,
            'word_dpi': self.word_dpi,
            'survey_word_width': self.survey_word_width,
            'survey_word_height': self.survey_word_height,
            'survey_word_dpi': self.survey_word_dpi,

            # Export settings
            'export_width': self.export_width,
            'export_height': self.export_height,
            'export_dpi': self.export_dpi,

            # Quick Settings
            'enable_quick_settings': self.enable_quick_settings,

            # Auto backup settings
            'enable_auto_backup': self.enable_auto_backup,
            'backup_interval': self.backup_interval,

            #Tines opened
            'times_opened': getattr(self, 'times_opened', 0),
            'registered': self.registered,
        }

        with open('config.json', 'w') as f:
            json.dump(config, f, indent=2)



    def on_preferences(self, event):
        pref_window = PreferenceWindow(self)
        pref_window.Show()

    def update_plot_preferences(self):
        self.plot_manager.update_plot_style(
            self.plot_style,
            self.scatter_size,
            self.line_width,
            self.line_alpha,
            self.scatter_color,
            self.line_color,
            self.scatter_marker,
            self.background_color,
            self.background_alpha,
            self.background_linestyle,
            self.envelope_color,
            self.envelope_alpha,
            self.envelope_linestyle,
            self.residual_color,
            self.residual_alpha,
            self.residual_linestyle,
            self.raw_data_linestyle,
            self.peak_colors,
            self.peak_alpha,
            self.background_thickness,
            self.envelope_thickness,
            self.residual_thickness
        )

        if hasattr(self, 'sheet_combobox'):
            selected_sheet = self.sheet_combobox.GetValue()
            if selected_sheet and 'FilePath' in self.Data and self.Data['FilePath']:
                self.clear_and_replot()
            else:
                print("No sheet selected or no file open. Skipping replot.")
        else:
            print("sheet_combobox not created yet. Skipping replot.")

        if hasattr(self, 'canvas'):
            self.canvas.draw_idle()

    def on_differentiate(self, event):
        if not hasattr(self, 'd_param_window') or not self.d_param_window:
            self.d_param_window = DParameterWindow(self)

            # Get the position and size of the main window
            main_pos = self.GetPosition()
            main_size = self.GetSize()

            # Get the size of the d-param window
            d_param_size = self.d_param_window.GetSize()

            # Calculate the position to center the d-param window on the main window
            x = main_pos.x + (main_size.width - d_param_size.width) // 2
            y = main_pos.y + (main_size.height - d_param_size.height) // 2

            # Set the position of the d-param window
            self.d_param_window.SetPosition((x, y))

        self.d_param_window.Show()
        self.d_param_window.Raise()

    # Add to imports section
    from libraries.On_BE_Corrections_Defs import on_be_correction_change, on_auto_be, apply_be_correction, \
        calculate_c1s_correction

    # Remove the existing function definitions and replace with:
    def on_be_correction_change(self, event):
        from libraries.On_BE_Corrections_Defs import on_be_correction_change
        on_be_correction_change(self, event)

    def on_auto_be(self, event):
        from libraries.On_BE_Corrections_Defs import on_auto_be
        on_auto_be(self, event)

    def apply_be_correction(self, correction):
        from libraries.On_BE_Corrections_Defs import apply_be_correction
        apply_be_correction(self, correction)

    def calculate_c1s_correction(self):
        from libraries.On_BE_Corrections_Defs import calculate_c1s_correction
        return calculate_c1s_correction(self)

    def load_be_correction(self):
        from libraries.On_BE_Corrections_Defs import load_be_correction
        load_be_correction(self)


    def on_toggle_peak_fill(self, event):
        new_state = self.plot_manager.toggle_peak_fill()
        self.clear_and_replot()
        print(f"Peak fill toggled in main window. New state: {new_state}")  # Debugging line

    def on_mini_help(self, event):
        from libraries.Help import show_quick_help
        show_quick_help(self)

    def on_undo(self, event):
        undo(self)

    def on_redo(self, event):
        redo(self)

    def on_close(self, event):
        # Close all matplotlib figures first
        import matplotlib.pyplot as plt
        plt.close('all')

        if self.backup_timer:
            self.backup_timer.Stop()

        self.Destroy()
        wx.GetApp().ExitMainLoop()

    def open_periodic_table(self, event):
        periodic_table = PeriodicTableWindow(self)

        # Set position relative to main window
        main_pos = self.GetPosition()
        main_size = self.GetSize()
        pt_size = periodic_table.GetSize()

        x = main_pos.x + (main_size.width - pt_size.width) // 2
        y = main_pos.y + (main_size.height - pt_size.height) // 2

        periodic_table.SetPosition((x, y))

        periodic_table.Show()

    def toggle_energy_scale(self):
        self.energy_scale = 'KE' if self.energy_scale == 'BE' else 'BE'
        self.plot_manager.clear_and_replot(self)

        # Update the menu item check state
        menu_item = self.GetMenuBar().FindItemById(self.toggle_energy_item.GetId())
        menu_item.Check(self.energy_scale == 'KE')

    def update_instrument(self, instrument):
        self.current_instrument = instrument
        # Recalculate peaks and update display
        self.recalculate_peaks()

    def recalculate_peaks(self):
        # Implement logic to update RSF and doublet splitting for all peaks
        # Then update peak fitting and display
        pass

    def on_text_size_increase(self, event):
        self.axis_title_size += 1
        self.axis_number_size += 1
        self.legend_font_size += 1
        self.core_level_text_size += 1
        self.label_font_size += 1

        # Update the preference window if it's open
        for window in wx.GetTopLevelWindows():
            if isinstance(window, PreferenceWindow):
                window.axis_title_spin.SetValue(self.axis_title_size)
                window.axis_num_spin.SetValue(self.axis_number_size)
                window.legend_size_spin.SetValue(self.legend_font_size)
                window.core_level_spin.SetValue(self.core_level_text_size)
                window.label_size_spin.SetValue(self.label_font_size)

        # Update plot
        self.update_plot_preferences()
        self.clear_and_replot()
        self.save_config()

    def on_text_size_decrease(self, event):
        # Make sure sizes don't go below minimum values
        if self.axis_title_size > 8:
            self.axis_title_size -= 1
        if self.axis_number_size > 8:
            self.axis_number_size -= 1
        if self.legend_font_size > 6:
            self.legend_font_size -= 1
        if self.core_level_text_size > 8:
            self.core_level_text_size -= 1
        if self.core_level_text_size > 6:
            self.label_font_size -= 1

        # Update the preference window if it's open
        for window in wx.GetTopLevelWindows():
            if isinstance(window, PreferenceWindow):
                window.axis_title_spin.SetValue(self.axis_title_size)
                window.axis_num_spin.SetValue(self.axis_number_size)
                window.legend_size_spin.SetValue(self.legend_font_size)
                window.core_level_spin.SetValue(self.core_level_text_size)
                window.label_size_spin.SetValue(self.label_font_size)

        # Update plot
        self.update_plot_preferences()
        self.clear_and_replot()
        self.save_config()

    def add_text_annotation(self, event):
        if event.inaxes:
            text = self.ax.text(event.xdata, event.ydata, "Text")
            draggable_text = DraggableText(text)
            self.canvas.draw()

    def open_labels_window(self, event):
        from libraries.Labels_Screen import LabelWindow
        if not hasattr(self, 'labels_window') or not self.labels_window:
            self.labels_window = LabelWindow(self)
        self.labels_window.Show()
        self.labels_window.Raise()

    def create_navigation_toolbar(self):
        if hasattr(self, 'navigation_toolbar') and self.navigation_toolbar:
            self.navigation_toolbar.Destroy()
        self.navigation_toolbar = NavigationToolbar(self.canvas)
        self.navigation_toolbar.Hide()

    def on_open_file_manager(self, event):
        """Open the file manager window"""
        from libraries.FileManager import FileManagerWindow
        if not hasattr(self, 'file_manager') or not self.file_manager:
            self.file_manager = FileManagerWindow(self)
        self.file_manager.Show()
        self.file_manager.Raise()

    def setup_backup_timer(self):
        """Set up the automatic backup timer based on preferences"""
        # Stop existing timer if it exists
        if self.backup_timer is not None:
            self.backup_timer.Stop()
            self.backup_timer = None

        if self.enable_auto_backup and hasattr(self, 'Data') and self.Data.get('FilePath'):
            # Get interval in milliseconds
            interval_ms = self.backup_interval * 60 * 1000

            # Create and start the timer
            self.backup_timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self.on_backup_timer, self.backup_timer)
            self.backup_timer.Start(interval_ms)
            print(f"Auto backup set up: Every {self.backup_interval} minutes")

    def on_backup_timer(self, event):
        """Handler for backup timer event"""
        # Only backup if we have a file open and changes made
        if hasattr(self, 'Data') and self.Data.get('FilePath') and self.history:
            from libraries.Utilities import perform_auto_backup
            perform_auto_backup(self)
        else:
            print("No changes or no file open, skipping auto backup")

    def on_grid_cell_click(self, event):
        row = event.GetRow()
        col = event.GetCol()

        if col == 7:  # Checkbox column
            # Get current state and toggle
            current_value = self.results_grid.GetCellValue(row, col)
            new_value = '0' if current_value == '1' else '1'

            # Get data structure key
            sheet_name = self.sheet_combobox.GetValue()
            match = re.search(r'(\d+)$', sheet_name)
            row_number = match.group(1) if match else "0"
            results_table_key = f'Results Table{row_number}'
            peak_label = f"Peak_{row}"

            # Update data structure with new checkbox state
            if results_table_key in self.Data and 'Peak' in self.Data[results_table_key]:
                if peak_label in self.Data[results_table_key]['Peak']:
                    self.Data[results_table_key]['Peak'][peak_label]['Checkbox'] = new_value

                    # Set cell value
                    self.results_grid.SetCellValue(row, col, new_value)

                    # Important: Update calculations immediately
                    self.update_atomic_percentages()

                    # Force complete grid refresh
                    self.results_grid.ForceRefresh()

                    # Redraw the plot to show changes
                    self.clear_and_replot()

                    # Save state for undo functionality
                    from libraries.Save import save_state
                    save_state(self)

                    self.after_checkbox_update()

            # Stop event propagation to prevent other handlers from overriding
            return

        event.Skip()

    def after_checkbox_update(self):
        # Update percentages
        self.update_atomic_percentages()

        # Make sure all checkboxes reflect the current data state
        self.update_checkboxes_from_data()

        # Update the plot to show changes
        self.clear_and_replot()

        # Make sure GUI is refreshed
        self.results_grid.Update()
        self.canvas.draw_idle()

        # Save state
        from libraries.Save import save_state
        save_state(self)

    def try_float(self, value, default=0.0):
        try:
            return float(value)
        except ValueError:
            return default

    def open_dream_nist_OLD(self):
        """Open the Dream NIST periodic table window"""
        from libraries.LibraryID import PeriodicTableXPS
        nist_window = PeriodicTableXPS()
        nist_window.mainloop()

    def open_kherve_db(self, kherveDB_wxpython=None):
        from libraries.kherveDB_wxpython import PeriodicTableXPS
        kherve_frame = PeriodicTableXPS()
        kherve_frame.Show()



def set_high_priority():
    try:
        proc = psutil.Process(os.getpid())
        if os.name == 'nt':  # Windows
            proc.nice(psutil.HIGH_PRIORITY_CLASS)
        else:  # Unix
            proc.nice(-20)
    except:
        pass

if __name__ == '__main__':

    multiprocessing.freeze_support()
    set_high_priority()

    app = wx.App(False)
    # app = wx.App(redirect=False, clearSigInt=False)

    # Create Splash Screen
    splash = show_splash(duration=2000, delay=0)

    # Detect OS
    os_name = platform.system()

    if os_name == "Darwin":  # Mac OS
        frame = MyFrame(None, "KherveFitting-v1.53 25f07")
    elif os_name == "Windows":
        frame = MyFrame(None, "KherveFitting-v1.53 25f07")
    else:
        frame = MyFrame(None, "KherveFitting-v1.53 25f07")

    # Apply preferences before showing the frame
    if hasattr(frame, 'times_opened') and frame.times_opened > 1:
        pref_window = PreferenceWindow(frame)
        pref_window.OnSave(None)  # Save settings without user interaction
        print("Preferences applied")

    frame.Show(True)

    if splash:
        splash.Destroy()

    # Open and close preference window to ensure consistent styling
    def apply_preferences(window):
        pref_window = PreferenceWindow(window)
        pref_window.OnSave(None)  # Save settings without user interaction
        print("Preferences applied")



    check_first_time_use(frame)


    updater = UpdateChecker()
    updater.check_update_delayed(frame)



    app.MainLoop()
    sys.exit(0)

