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

from libraries.Peak_Functions import PeakFunctions

# from libraries.Peak_Functions import AtomicConcentrations
# from libraries.Peak_Functions import gauss_lorentz, S_gauss_lorentz

from Functions import toggle_Col_1, update_sheet_names, rename_sheet
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


class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(1640, 740))

        # Get the directory of the current script
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # Construct the path to the icon file
        icon_path = os.path.join(current_dir, "Icons", "Icon.ico")

        FIRST_TIME_USE = True

        # Set the icon
        icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        self.SetMinSize((800, 600))
        self.panel = wx.Panel(self)

        # Will hold reference to FileManagerWindow when opened
        self.file_manager = None

        # Center the window on the screen
        self.Centre()

        self.Data = Init_Measurement_Data(self)

        #
        self.previous_size = None

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
        self.scatter_size = 20
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

        # Load config if exists
        self.load_config()

        create_widgets(self)
        create_menu(self)
        load_recent_files_from_config(self)



        create_statusbar(self)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)


        self.canvas.mpl_connect("button_press_event", self.on_click)
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        self.canvas.mpl_connect('key_press_event', self.on_key_press)
        self.canvas.mpl_connect('scroll_event', self.on_mouse_wheel)
        self.canvas.mpl_connect('button_press_event', self.on_right_click)
        self.Bind(wx.EVT_CHAR_HOOK, self.on_key_press_global)
        # self.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)
        # self.add_cross_to_peak(self.selected_peak_index)

        self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_peak_params_cell_changed)
        self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGING, self.on_peak_params_cell_changed)
        self.Bind(wx.EVT_CLOSE, self.on_close)
        # Bind right-click events for peak_params_grid

        # self.peak_params_grid.Bind(wx.EVT_RIGHT_DOWN, self.on_peak_params_grid_right_click)  # For empty grid
        self.peak_params_grid.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_peak_params_right_click)
        # self.peak_params_grid.Bind(wx.EVT_CONTEXT_MENU, self.on_peak_params_context_menu)
        # Bind right-click events for peak_params_grid



        self.plot_manager.residuals_state = 2  # Set default state


        # Add this method too:
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



    # I DON'T THINK IT IS USED
    def on_listbox_selection(self, event):
        self.selected_indices = self.file_listbox.GetSelections()
        update_sheet_names(self)


    def on_change_sheet_name(self, event):
        new_sheet_name = self.change_to_textctrl.GetValue()
        if new_sheet_name:
            rename_sheet(self, new_sheet_name)
            wx.MessageBox(f'Sheet renamed to "{new_sheet_name}"', "Info", wx.OK | wx.ICON_INFORMATION)





    def add_peak_params(self):
        if hasattr(self, 'fitting_window'):
            self.selected_fitting_method = self.fitting_window.model_combobox.GetValue()
            print(f'Fitting method: {self.selected_fitting_method}')
        save_state(self)
        sheet_name = self.sheet_combobox.GetValue()

        if self.bg_min_energy is None or self.bg_max_energy is None:
            wx.MessageBox("Please create a background first.", "No Background", wx.OK | wx.ICON_WARNING)
            return None

        num_peaks = self.peak_params_grid.GetNumberRows() // 2

        # Update bg_min_energy and bg_max_energy from window.Data
        if sheet_name in self.Data['Core levels'] and 'Background' in self.Data['Core levels'][sheet_name]:
            background_data = self.Data['Core levels'][sheet_name]['Background']
            self.bg_min_energy = background_data.get('Bkg Low')
            self.bg_max_energy = background_data.get('Bkg High')

        # Ensure bg_min_energy and bg_max_energy are not None
        if self.bg_min_energy is None or self.bg_max_energy is None:
            wx.MessageBox("Background range is not set. Please set the background first.", "Warning",
                          wx.OK | wx.ICON_WARNING)
            return

        if num_peaks == 0:
            residual = self.y_values - np.array(self.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
            peak_y = residual[np.argmax(residual)]
            peak_x = self.x_values[np.argmax(residual)]
        else:

            # Call update_overall_fit_and_residuals to get the residuals
            residual = self.plot_manager.update_overall_fit_and_residuals(self)

            if residual is not None:
                peak_y = residual.max()
                peak_x = self.x_values[np.argmax(residual)]
            else:
                # Fallback if residuals couldn't be calculated
                wx.MessageBox("Unable to calculate residuals. Using default peak position.", "Warning",
                              wx.OK | wx.ICON_WARNING)
                peak_y = self.y_values.max()
                peak_x = self.x_values[np.argmax(self.y_values)]

        self.peak_count += 1

        # Add new rows to the grid
        self.peak_params_grid.AppendRows(2)
        self.add_choice_editor_to_new_row(self.peak_params_grid, self.peak_params_grid.GetNumberRows() - 2)
        row = self.peak_params_grid.GetNumberRows() - 2

        # Assign letter IDs
        letter_id = chr(64 + self.peak_count)


        # Set values in the grid
        self.peak_params_grid.SetCellValue(row, 0, letter_id)
        self.peak_params_grid.SetReadOnly(row, 0)
        self.peak_params_grid.SetCellValue(row, 1, f"{sheet_name} p{self.peak_count}")
        self.peak_params_grid.SetCellValue(row, 2, f"{peak_x:.2f}")
        self.peak_params_grid.SetCellValue(row, 3, f"{peak_y:.2f}")
        self.peak_params_grid.SetCellValue(row, 4, "1.6")
        self.peak_params_grid.SetCellValue(row, 5, "20")
        if self.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.peak_params_grid.SetCellValue(row, 6, f"{peak_y * 1.6 * 1.064:.2f}")
        else:
            self.peak_params_grid.SetCellValue(row, 6, f"{peak_y * 1.6 * 1.064:.2f}")
        if self.selected_fitting_method == "ExpGauss.(Area, \u03c3, \u03b3)":
            self.peak_params_grid.SetCellValue(row, 7, "0.3")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '1.2')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, '0.64')  # skew
        elif self.selected_fitting_method in ["LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.peak_params_grid.SetCellValue(row, 5, "50")
            self.peak_params_grid.SetCellValue(row, 7, "2.7")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '2.7')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, '0.64')  # skew
        elif self.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)"]:
            self.peak_params_grid.SetCellValue(row, 5, "50")
            self.peak_params_grid.SetCellValue(row, 7, "2.7")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '2.7')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, '0')  # skew
        elif self.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            self.peak_params_grid.SetCellValue(row, 5, "20")
            self.peak_params_grid.SetCellValue(row, 7, "1.2")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '0.4')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, '0.01')  # skew
        elif self.selected_fitting_method in ["D-parameter"]:
            self.peak_params_grid.SetCellValue(row, 5, "2")
            self.peak_params_grid.SetCellValue(row, 7, "1")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '1')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, '7')  # skew
        else:
            self.peak_params_grid.SetCellValue(row, 7, "1")  # sigma
            self.peak_params_grid.SetCellValue(row, 8, '0.15')  # gamma
            self.peak_params_grid.SetCellValue(row, 9, "0.64")  # Default value for skewed
        self.peak_params_grid.SetCellValue(row, 10, '')
        self.peak_params_grid.SetCellValue(row, 11, '')
        self.peak_params_grid.SetCellValue(row, 12, '')  # Split, initially empty
        self.peak_params_grid.SetCellValue(row, 13, self.selected_fitting_method)  # Fitting Model
        self.peak_params_grid.SetCellValue(row, 14, self.background_method)  # Bkg Type
        self.peak_params_grid.SetCellValue(row, 15,
                                           f"{self.bg_min_energy:.2f}" if self.bg_min_energy is not None else "")  # Bkg Low
        self.peak_params_grid.SetCellValue(row, 16,
                                           f"{self.bg_max_energy:.2f}" if self.bg_max_energy is not None else "")  # Bkg High
        self.peak_params_grid.SetCellValue(row, 17, f"{float(self.offset_l):.2f}")  # Bkg Offset Low
        self.peak_params_grid.SetCellValue(row, 18, f"{self.offset_h:.2f}")  # Bkg Offset High

        # Set position constraint to background range
        position_constraint = f"{self.bg_min_energy:.2f},{self.bg_max_energy:.2f}"
        self.peak_params_grid.SetCellValue(row + 1, 2, position_constraint)
        self.peak_params_grid.SetCellValue(row + 1, 3, "1:1e7")
        self.peak_params_grid.SetCellValue(row + 1, 4, "0.3:3.5")
        self.peak_params_grid.SetCellValue(row + 1, 5, "2:80")
        self.peak_params_grid.SetCellValue(row + 1, 6, "1:1e7")
        if self.selected_fitting_method == "ExpGauss.(Area, \u03c3, \u03b3)":
            self.peak_params_grid.SetCellValue(row + 1, 7, "0.01:1")
            self.peak_params_grid.SetCellValue(row + 1, 8, "0.01:3")
            self.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        elif self.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                                            "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            self.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")
            self.peak_params_grid.SetCellValue(row + 1, 7, "0.01:10")
            self.peak_params_grid.SetCellValue(row + 1, 8, "0.01:10")
            self.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        elif self.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            self.peak_params_grid.SetCellValue(row + 1, 5, "15:85")
            self.peak_params_grid.SetCellValue(row + 1, 7, "0.2:1.5")
            self.peak_params_grid.SetCellValue(row + 1, 8, "0.2:1.5")
            self.peak_params_grid.SetCellValue(row + 1, 9, '0.01:0.7')  # skew
        else:
            self.peak_params_grid.SetCellValue(row + 1, 7, "0.3:3")
            self.peak_params_grid.SetCellValue(row + 1, 8, "0.3:3")
            self.peak_params_grid.SetCellValue(row + 1, 9, '0.01:2')  # skew
        self.peak_params_grid.ForceRefresh()

        # Set constraint values
        self.peak_params_grid.SetReadOnly(row + 1, 0)
        for col in range(self.peak_params_grid.GetNumberCols()+1):  # Assuming you have 15 columns in total
            # self.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(230, 230, 230))
            self.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200,245,228))

        for col in [10, 11, 12]:  # Columns for Area, sigma and gamma
            self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(27, 140, 60))

        # Set background color for Height, FWHM, and L/G ratio cells if Voigt function
        if self.selected_fitting_method == "Voigt (Area, L/G, \u03c3)":
            for col in [3, 4, 8]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [5, 6, 7]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.peak_params_grid.SetCellValue(row, col, "0")
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.selected_fitting_method == "Voigt (Area, L/G, \u03c3, S)":
            for col in [3, 4, 8]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [5, 6, 7, 9]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.selected_fitting_method in ["Voigt (Area, \u03c3, \u03b3)", "ExpGauss.(Area, \u03c3, \u03b3)"]:
            for col in [3,4, 5, 9]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [6, 7, 8]:  # Columns for Height, FWHM
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.peak_params_grid.SetCellValue(row, col, "0")
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)" ]:
            for col in [3, 5]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 6, 7, 8]:  # Columns for Height, FWHM
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.peak_params_grid.SetCellValue(row, col, "0")
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.selected_fitting_method in ["LA (Area, \u03c3/\u03b3, \u03b3)"]:
            for col in [3,7,]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6, 8]:  # Columns for Height, FWHM
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
            for col in [9]:  # Columns for Area, sigma and gamma
                self.peak_params_grid.SetCellValue(row, col, "0")
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(255, 255, 255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200, 245, 228))
        elif self.selected_fitting_method in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            for col in [3,7]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6, 8, 9]:  # Columns for Height, FWHM
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        elif self.selected_fitting_method in  ["Pseudo-Voigt (Area)", "GL (Area)", "SGL (Area)"]:
            for col in [3]:  # Height
                # self.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(200,245,228))
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [7,8, 9]:  # Columns for Area, sigma and gamma
                # self.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                self.peak_params_grid.SetCellValue(row , col, "0")
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(255,255,255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [4, 5, 6]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))
        else:
            for col in [6]:  # Columns for Area, sigma and gamma
                # self.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(128, 128, 128))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [7,8, 9]:  # Columns for Area, sigma and gamma
                # self.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(240,240,240))
                self.peak_params_grid.SetCellValue(row + 1, col, "0")
                self.peak_params_grid.SetCellTextColour(row , col, wx.Colour(255,255,255))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(200,245,228))
            for col in [3, 4, 5]:  # Columns for Height, FWHM, L/G ratio
                self.peak_params_grid.SetCellTextColour(row, col, wx.Colour(0, 0, 0))
                self.peak_params_grid.SetCellTextColour(row + 1, col, wx.Colour(0, 0, 0))

        # Set selected_peak_index to the index of the new peak
        self.selected_peak_index = num_peaks

        # Update the Data structure with the new peak information
        if 'Fitting' not in self.Data['Core levels'][sheet_name]:
            self.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.Data['Core levels'][sheet_name]['Fitting']:
            self.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}



        if self.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': peak_y * 1.6 * 1.064,
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.01,
                'Fitting Model': self.selected_fitting_method,
                'Bkg Type': self.background_method,
                'Bkg Low': self.bg_min_energy,
                'Bkg High': self.bg_max_energy,
                'Bkg Offset Low': self.offset_l,
                'Bkg Offset High': self.offset_h,
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
        else:
            peak_data = {
                'Position': peak_x,
                'Height': peak_y,
                'FWHM': 1.6,
                'L/G': 20,
                'Area': peak_y * 1.6 * 1.064,
                'Sigma': 1.2,
                'Gamma': 0.4,
                'Skew': 0.64,
                'Fitting Model': self.selected_fitting_method,
                'Bkg Type': self.background_method,
                'Bkg Low': self.bg_min_energy,
                'Bkg High': self.bg_max_energy,
                'Bkg Offset Low': self.offset_l,
                'Bkg Offset High': self.offset_h,
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

        if self.selected_fitting_method in ["LA (Area, \u03c3, \u03b3)","LA (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_data.update({
                'L/G': 50,  # Default L/G ratio for LA model
                'Sigma': 2.75,
                'Gamma': 2.75,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",  # Wider range for L/G in LA model
                    'Area': '1:1e7',
                    'Sigma': "0.01:10",
                    'Gamma': "0.01:10"
                }
            })
        elif self.selected_fitting_method in ["LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
            peak_data.update({
                'L/G': 50,  # Default L/G ratio for LA model
                'Sigma': 2.75,
                'Gamma': 2.75,
                'Skew': 0.64,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "Fixed",  # Wider range for L/G in LA model
                    'Area': '1:1e7',
                    'Sigma': "0.01:4",
                    'Gamma': "0.01:4",
                    'Skew': "0.01:2"
                }
            })
        elif self.selected_fitting_method in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)"]:
            peak_data.update({
                'Sigma': 1,
                'Gamma': 0.5,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3"
                }
            })
        elif self.selected_fitting_method in ["Voigt (Area, L/G, \u03c3, S)"]:
            peak_data.update({
                'Sigma': 1,
                'Gamma': 0.5,
                'skew': 0.01,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:0.7"
                }
            })
        elif self.selected_fitting_method in ["SGL (Area)"]:
            # Value required to calculate area
            fwhm = 1.6
            fraction = 20
            sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
            gamma = fwhm / 2
            sgl_area = peak_y * ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (fraction / 100) * np.pi * gamma)
            peak_data.update({
                'Area': sgl_area,
                'Constraints': {
                    'Position': position_constraint,
                    'Height': "1:1e7",
                    'FWHM': "0.3:3.5",
                    'L/G': "1:80",  # Full range for Voigt models
                    'Area': '1:1e7',
                    'Sigma': "0.3:3",
                    'Gamma': "0.3:3",
                    'Skew': "0.01:0.7"
                }
            })

        self.Data['Core levels'][sheet_name]['Fitting']['Peaks'][sheet_name + f" p{self.peak_count}"] = peak_data
        # print(self.Data)
        self.show_hide_vlines()

        # Call the method to clear and replot everything
        self.clear_and_replot()

        return self.peak_count - 1  # Return the index of the added peak



    def number_to_letter(n):
        return chr(65 + n)  # 65 is the ASCII value for 'A'

    # --------------------------------------------------------------------------------------------------
    # OPEN WINDOW --------------------------------------------------------------------------------------
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
            self.deselect_all_peaks()

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

    def update_peak_fwhm(self, x):
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

                self.remove_cross_from_peak()
                self.highlight_selected_peak()  # Highlight the selected peak
            else:
                self.selected_peak_index = None
                self.deselect_all_peaks()
        else:
            # If peak fitting tab is not selected, don't allow peak selection
            self.selected_peak_index = None
            self.deselect_all_peaks()

        self.update_checkboxes_from_data()
        event.Skip()


    def remove_cross_from_peak(self):
        if hasattr(self, 'cross'):
            if self.cross in self.ax.lines:
                self.cross.remove()
            del self.cross
        self.canvas.mpl_disconnect('motion_notify_event')
        self.canvas.mpl_disconnect('button_release_event')

    def update_peak_grid(self, index, x, y):
        row = index * 2  # Assuming each peak uses two rows in the grid
        self.peak_params_grid.SetCellValue(row, 2, f"{x:.2f}")  # Update Position
        self.peak_params_grid.SetCellValue(row, 3, f"{y:.2f}")  # Update Height
        self.peak_params_grid.ForceRefresh()  # Refresh the grid to show changes

    def update_fwhm_grid(self, index, fwhm):
        row = index * 2  # Assuming each peak uses two rows in the grid
        self.peak_params_grid.SetCellValue(row, 4, f"{fwhm:.2f}")  # Update FWHM


    def on_cross_drag(self, event):
        if event.inaxes and self.selected_peak_index is not None:
            row = self.selected_peak_index * 2

            if row >= self.peak_params_grid.GetNumberRows():
                self.selected_peak_index = None
                return

            fitting_model = self.peak_params_grid.GetCellValue(row, 13)

            if event.button == 1:
                try:
                    if event.key == 'shift':
                        new_fwhm = self.update_peak_fwhm(event.xdata)
                        if new_fwhm is not None:
                            self.update_linked_fwhm_recursive(self.selected_peak_index, new_fwhm)

                    elif self.is_mouse_on_peak(event):
                        closest_index = np.argmin(np.abs(self.x_values - event.xdata))
                        bkg_y = self.background[closest_index]
                        new_x = event.xdata
                        new_height = max(event.ydata - bkg_y, 0)

                        if "LA" in fitting_model:
                            # Calculate new area for LA models
                            fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
                            sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                            gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                            skew = float(
                                self.peak_params_grid.GetCellValue(row, 9)) if "LA*G" in fitting_model else None
                            new_area = self.calculate_peak_area(fitting_model, new_height, fwhm, 0, sigma, gamma, skew)

                            self.update_peak(self.selected_peak_index, new_x, new_height, new_area)
                            self.update_linked_peaks_recursive(self.selected_peak_index, new_x, new_height, new_area)
                        else:
                            self.update_peak(self.selected_peak_index, new_x, new_height)
                            self.update_linked_peaks_recursive(self.selected_peak_index, new_x, new_height)

                    self.update_ratios()
                    self.clear_and_replot()
                    self.plot_manager.add_cross_to_peak(self, self.selected_peak_index)
                    self.canvas.draw_idle()

                except Exception as e:
                    print(f"Error during cross drag: {e}")

    def on_cross_release(self, event):
        save_state(self)
        if event.inaxes and self.selected_peak_index is not None:
            row = self.selected_peak_index * 2
            fitting_model = self.peak_params_grid.GetCellValue(row, 13)
            peak_label = self.peak_params_grid.GetCellValue(row, 1)
            sheet_name = self.sheet_combobox.GetValue()

            x = event.xdata
            y = event.ydata
            bkg_y = self.background[np.argmin(np.abs(self.x_values - x))]

            if event.button == 1:
                if event.key == 'shift':
                    new_fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
                    self.update_linked_fwhm_recursive(self.selected_peak_index, new_fwhm)
                else:
                    y = max(y - bkg_y, 0)

                    if "LA" in fitting_model:
                        current_area = float(self.peak_params_grid.GetCellValue(row, 6))
                        self.update_peak(self.selected_peak_index, x, y, current_area)
                        self.update_linked_peaks_recursive(self.selected_peak_index, x, y, current_area)
                    else:
                        self.update_peak(self.selected_peak_index, x, y)
                        self.update_linked_peaks_recursive(self.selected_peak_index, x, y)

            self.remove_cross_from_peak()
            self.cross = self.ax.plot(x, y + bkg_y, 'bx', markersize=15, markerfacecolor='none', picker=5, linewidth=3)[
                0]
            self.canvas.draw_idle()

        if hasattr(self, 'motion_cid'):
            self.canvas.mpl_disconnect(self.motion_cid)
            delattr(self, 'motion_cid')
        if hasattr(self, 'release_cid'):
            self.canvas.mpl_disconnect(self.release_cid)
            delattr(self, 'release_cid')

        self.refresh_peak_params_grid_release()

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

        if ("LA" in fitting_model or "GL (Area)" in fitting_model or "Voigt" in fitting_model or "ExpGauss" in fitting_model) and area_constraint.startswith(
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
            # Original height-based logic
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

            # # Create model instance first
            # skew_voigt_model = lmfit.models.SkewedVoigtModel()
            #
            # # Then evaluate with parameters
            # x_range = np.linspace(center - 10 * sigma, center + 10 * sigma, 1000)
            # y_values = skew_voigt_model.eval(x=x_range, amplitude=area, center=center, sigma=sigma, gamma=gamma,
            #                                  skew=skew)
            # height = np.max(y_values)
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
            # For Pseudo-Voigt, this is also an approximation
            return area / (fwhm * np.pi / 2)

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


    def update_peak(self, peak_index, new_x, new_height, area=None):
        row = peak_index * 2
        sheet_name = self.sheet_combobox.GetValue()
        peak_label = self.peak_params_grid.GetCellValue(row, 1)
        fitting_model = self.peak_params_grid.GetCellValue(row, 13)

        # Update the grid
        self.peak_params_grid.SetCellValue(row, 2, f"{new_x:.2f}")

        if "LA" in fitting_model and area is not None:
            self.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")  # Set area first
            self.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")  # Height is derived
        else:
            self.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")

        # Update the Data structure
        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            if peak_label in peaks:
                peaks[peak_label]['Position'] = new_x
                if "LA" in fitting_model and area is not None:
                    peaks[peak_label]['Area'] = area
                    peaks[peak_label]['Height'] = new_height
                else:
                    peaks[peak_label]['Height'] = new_height

        # Recalculate area if not LA model
        if not "LA" in fitting_model:
            self.recalculate_peak_area(peak_index)

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
                elif model in ["Voigt (Area, L/G, \u03c3, S)"] and new_sigma is not None:
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

                if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
                    original_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                    original_lg = float(self.peak_params_grid.GetCellValue(row, 5))
                    linked_sigma = original_sigma
                    linked_gamma = (original_lg / 100 * linked_sigma) / (1 - original_lg / 100)
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 7, f"{linked_sigma:.2f}")
                    self.peak_params_grid.SetCellValue(linked_peak * 2, 8, f"{linked_gamma:.2f}")
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
                self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
                new_fwhm = self.initial_fwhm
            elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
                current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                # new_sigma = max(current_sigma + delta_x * 0.5, 0.2)
                new_gamma = max(current_gamma + delta_x * 0.5, 0.2)
                self.peak_params_grid.SetCellValue(row, 7, f"{current_sigma:.2f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
                new_fwhm = self.initial_fwhm
            else:
                new_fwhm = max(self.initial_fwhm + delta_x * 1, 0.3)
                self.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")

            # Update FWHM in window.Data
            sheet_name = self.sheet_combobox.GetValue()
            if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][
                sheet_name] and 'Peaks' in self.Data['Core levels'][sheet_name]['Fitting']:
                peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                if peak_label in peaks:
                    peaks[peak_label]['FWHM'] = new_fwhm
                    if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                 "ExpGauss.(Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
                        peaks[peak_label]['Sigma'] = new_sigma
                        peaks[peak_label]['Gamma'] = new_gamma

            # Recalculate area
            self.recalculate_peak_area(self.selected_peak_index)

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
                        "\u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
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

    def show_hide_vlines_OLD(self):
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
            '''
            row = self.selected_peak_index * 2
            x_peak = float(self.peak_params_grid.GetCellValue(row, 2))  # Position
            y_peak = float(self.peak_params_grid.GetCellValue(row, 3))  # Height
            bkg_y = self.background[np.argmin(np.abs(self.x_values - x_peak))]
            y_peak += bkg_y
            distance = np.sqrt((event.xdata - x_peak) ** 2 + (event.ydata - y_peak) ** 2)            
            return distance < 3000  # Smaller tolerance for more precise detection
            '''
            return True
        return False


    # GET PEAK INDEX ------------------------------------------------------------
    # USED TO MOVE THE PE
    def get_peak_index_from_position(self, x, y):
        # Transform input coordinates (x, y) to display (pixel) coordinates
        x_display, y_display = self.ax.transData.transform((x, y))

        num_peaks = self.peak_params_grid.GetNumberRows() // 2  # Assuming each peak uses two rows

        for i in range(num_peaks):
            row = i * 2  # Assuming each peak has 2 rows (peak and constraints)
            if self.peak_params_grid.IsInSelection(row, 0):
                # Get the peak position in data coordinates
                x_peak = float(self.peak_params_grid.GetCellValue(row, 2))  # Position
                y_peak = float(self.peak_params_grid.GetCellValue(row, 3))  # Height
                bkg_y = self.background[np.argmin(np.abs(self.x_values - x_peak))]
                y_peak += bkg_y

                # Transform peak position to display (pixel) coordinates
                x_peak_display, y_peak_display = self.ax.transData.transform((x_peak, y_peak))

                # Calculate the Euclidean distance in display coordinates
                distance = np.sqrt((x_display - x_peak_display) ** 2 + (y_display - y_peak_display) ** 2)

                if distance < 10:  # Adjust the tolerance as needed (10 pixels here as an example)
                    return i
        return None


    def on_mouse_move(self, event):
        if event.inaxes:
            x, y = event.xdata, event.ydata
            if self.energy_scale == 'KE':
                self.SetStatusText(f"KE: {x:.1f} eV, I: {int(y)} CPS", 1)
            else:
                self.SetStatusText(f"BE: {x:.1f} eV, I: {int(y)} CPS", 1)

    def on_click(self, event):
        if event.inaxes:
            x_click = event.xdata
            if event.button == 1 and event.key == 'shift' and self.background_tab_selected:
                # Store current motion handler if it exists
                self.motion_notify_id = self.canvas.mpl_connect('motion_notify_event', self.on_motion)
                self.button_release_id = self.canvas.mpl_connect('button_release_event', self.on_release)

                x_click = event.xdata
                sheet_name = self.sheet_combobox.GetValue()
                if self.vline1 is not None and self.vline2 is not None:
                    vline1_x = self.vline1.get_xdata()[0]
                    vline2_x = self.vline2.get_xdata()[0]

                    # Determine which vline is at low/high BE
                    low_be_x = min(vline1_x, vline2_x)
                    high_be_x = max(vline1_x, vline2_x)

                    dist1 = abs(x_click - vline1_x)
                    dist2 = abs(x_click - vline2_x)

                    if dist1 < dist2:  # Closer to vline1
                        raw_y = self.y_values[np.argmin(np.abs(self.x_values - vline1_x))]
                        if vline1_x == low_be_x:
                            self.offset_l = event.ydata - raw_y
                            self.Data['Core levels'][sheet_name]['Background']['Bkg Offset Low'] = self.offset_l
                            self.fitting_window.offset_l_text.SetValue(f'{self.offset_l:.1f}')
                        else:
                            self.offset_h = event.ydata - raw_y
                            self.Data['Core levels'][sheet_name]['Background']['Bkg Offset High'] = self.offset_h
                            self.fitting_window.offset_h_text.SetValue(f'{self.offset_h:.1f}')
                    else:  # Closer to vline2
                        raw_y = self.y_values[np.argmin(np.abs(self.x_values - vline2_x))]
                        if vline2_x == low_be_x:
                            self.offset_l = event.ydata - raw_y
                            self.Data['Core levels'][sheet_name]['Background']['Bkg Offset Low'] = self.offset_l
                            self.fitting_window.offset_l_text.SetValue(f'{self.offset_l:.1f}')
                        else:
                            self.offset_h = event.ydata - raw_y
                            self.Data['Core levels'][sheet_name]['Background']['Bkg Offset High'] = self.offset_h
                            self.fitting_window.offset_h_text.SetValue(f'{self.offset_h:.1f}')

                    self.plot_manager.plot_background(self)
            elif event.button == 1:  # Left click
                if event.key == 'shift':  # SHIFT + left click
                    if self.peak_fitting_tab_selected and self.selected_peak_index is not None:
                        row = self.selected_peak_index * 2  # Each peak uses two rows in the grid
                        self.initial_fwhm = float(self.peak_params_grid.GetCellValue(row, 4))  # FWHM
                        self.initial_x = event.xdata
                        self.motion_cid = self.canvas.mpl_connect('motion_notify_event', self.on_cross_drag)
                        self.release_cid = self.canvas.mpl_connect('button_release_event', self.on_cross_release)
                elif self.background_tab_selected:  # Left click and background tab selected
                    self.deselect_all_peaks()
                    sheet_name = self.sheet_combobox.GetValue()
                    if sheet_name in self.Data['Core levels']:
                        core_level_data = self.Data['Core levels'][sheet_name]
                        if self.background_method == "Multi-Regions Smart":
                            # Check if click is close to vline1 or vline2
                            if self.vline1 is not None and self.vline2 is not None:
                                vline1_x = self.vline1.get_xdata()[0]
                                vline2_x = self.vline2.get_xdata()[0]

                                dist1 = abs(x_click - vline1_x)
                                dist2 = abs(x_click - vline2_x)

                                if dist1 < dist2 and dist1 < self.some_threshold:
                                    self.moving_vline = self.vline1
                                elif dist2 < self.some_threshold:
                                    self.moving_vline = self.vline2
                                else:
                                    self.moving_vline = None

                                if self.moving_vline is not None:
                                    self.motion_cid = self.canvas.mpl_connect('motion_notify_event', self.on_motion)
                                    self.release_cid = self.canvas.mpl_connect('button_release_event', self.on_release)
                                    return  # Exit to avoid creating new vlines

                        # Existing vline creation logic
                        if self.vline1 is None:
                            self.vline1 = self.ax.axvline(x_click, color='r', linestyle='--')
                            core_level_data['Background']['Bkg Low'] = float(x_click)
                        elif self.vline2 is None and abs(
                                x_click - core_level_data['Background']['Bkg Low']) > self.some_threshold:
                            self.vline2 = self.ax.axvline(x_click, color='r', linestyle='--')
                            core_level_data['Background']['Bkg High'] = float(x_click)
                            core_level_data['Background']['Bkg Low'], core_level_data['Background'][
                                'Bkg High'] = sorted([
                                core_level_data['Background']['Bkg Low'],
                                core_level_data['Background']['Bkg High']
                            ])
                        else:
                            self.moving_vline = self.vline1 if self.vline2 is None or abs(
                                x_click - core_level_data['Background']['Bkg Low']) < abs(
                                x_click - core_level_data['Background']['Bkg High']) else self.vline2
                            self.motion_cid = self.canvas.mpl_connect('motion_notify_event', self.on_motion)
                            self.release_cid = self.canvas.mpl_connect('button_release_event', self.on_release)
                elif self.noise_tab_selected:
                    if self.vline3 is None:
                        self.vline3 = self.ax.axvline(x_click, color='b', linestyle='--')
                        self.noise_min_energy = float(x_click)
                    elif self.vline4 is None and abs(x_click - self.noise_min_energy) > self.some_threshold:
                        self.vline4 = self.ax.axvline(x_click, color='b', linestyle='--')
                        self.noise_max_energy = float(x_click)
                        self.noise_min_energy, self.noise_max_energy = sorted(
                            [self.noise_min_energy, self.noise_max_energy])
                    else:
                        self.moving_vline = self.vline3 if self.vline4 is None or abs(
                            x_click - self.noise_min_energy) < abs(
                            x_click - self.noise_max_energy) else self.vline4
                        self.motion_cid = self.canvas.mpl_connect('motion_notify_event', self.on_motion)
                        self.release_cid = self.canvas.mpl_connect('button_release_event', self.on_release)
                elif self.peak_fitting_tab_selected:  # Only allow peak selection when peak fitting tab is selected
                    peak_index = self.get_peak_index_from_position(event.xdata, event.ydata)
                    if peak_index is not None:
                        self.selected_peak_index = peak_index
                        self.motion_cid = self.canvas.mpl_connect('motion_notify_event', self.on_cross_drag)
                        self.release_cid = self.canvas.mpl_connect('button_release_event', self.on_cross_release)
                        self.highlight_selected_peak()
                    else:
                        self.deselect_all_peaks()
                else:
                    self.deselect_all_peaks()

            self.show_hide_vlines()
            self.canvas.draw()



    def on_mouse_wheel(self, event):
        if event.step != 0:
            current_index = self.sheet_combobox.GetSelection()
            num_sheets = self.sheet_combobox.GetCount()

            if event.step > 0:
                # Scroll up, move to previous sheet
                new_index = (current_index - 1) % num_sheets
            else:
                # Scroll down, move to next sheet
                new_index = (current_index + 1) % num_sheets

            self.sheet_combobox.SetSelection(new_index)
            new_sheet = self.sheet_combobox.GetString(new_index)

            # Create a mock event to pass to on_sheet_selected
            mock_event = wx.CommandEvent(wx.EVT_COMBOBOX.typeId, self.sheet_combobox.GetId())
            mock_event.SetString(new_sheet)

            # Call on_sheet_selected with the mock event
            on_sheet_selected(self, mock_event)
            save_state(self)

    def highlight_selected_peak(self):
        if self.selected_peak_index is not None:
            num_peaks = self.peak_params_grid.GetNumberRows() // 2

            if self.selected_peak_index >= num_peaks:
                print(
                    f"Warning: selected_peak_index ({self.selected_peak_index}) is out of range. Max index: {num_peaks - 1}")
                self.selected_peak_index = None
                return

            for i in range(num_peaks):
                row = i * 2
                is_selected = (i == self.selected_peak_index)
                self.peak_params_grid.SetCellBackgroundColour(row, 0, wx.LIGHT_GREY if is_selected else wx.WHITE)
                self.peak_params_grid.SetCellBackgroundColour(row + 1, 0, wx.LIGHT_GREY if is_selected else wx.WHITE)

            row = self.selected_peak_index * 2

            if row < self.peak_params_grid.GetNumberRows():
                peak_label = self.peak_params_grid.GetCellValue(row, 1)
                x_str = self.peak_params_grid.GetCellValue(row, 2)
                y_str = self.peak_params_grid.GetCellValue(row, 3)

                if x_str and y_str:
                    try:
                        x = float(x_str)
                        y = float(y_str)
                        y += self.background[np.argmin(np.abs(self.x_values - x))]

                        self.remove_cross_from_peak()
                        if hasattr(self, 'peak_letter') and self.peak_letter:
                            self.peak_letter.remove()

                        self.cross, = self.ax.plot(x, y, 'bx', markersize=15, markerfacecolor='none', picker=5,
                                                   linewidth=3)

                        peak_letter = chr(65 + self.selected_peak_index)
                        max_y = self.ax.get_ylim()[1]
                        y_offset = max_y * 0.02
                        self.peak_letter = self.ax.text(x, y + y_offset, peak_letter, ha='center', va='bottom',
                                                        fontsize=12)

                        self.peak_params_grid.ClearSelection()
                        self.peak_params_grid.SelectRow(row, addToSelected=False)
                        self.peak_params_grid.Refresh()
                        self.canvas.draw_idle()

                        sheet_name = self.sheet_combobox.GetValue()
                        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][
                            sheet_name] and 'Peaks' in self.Data['Core levels'][sheet_name]['Fitting']:
                            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                            if self.selected_peak_index < len(peaks):
                                old_label = list(peaks.keys())[self.selected_peak_index]
                                if old_label != peak_label:
                                    peaks[peak_label] = peaks.pop(old_label)
                            else:
                                print(
                                    f"Warning: selected_peak_index ({self.selected_peak_index}) is out of range for peaks in Data structure")

                        self.canvas.mpl_connect('motion_notify_event', self.on_cross_drag)
                        self.canvas.mpl_connect('button_release_event', self.on_cross_release)
                    except ValueError as e:
                        print(f"Warning: Invalid data for selected peak. Cannot highlight. Error: {e}")
                else:
                    print(f"Warning: Empty data for selected peak. Cannot highlight.")
            else:
                print(f"Warning: Row {row} does not exist in peak_params_grid")

            self.peak_params_grid.Refresh()
        else:
            print("No peak selected (selected_peak_index is None)")

    def on_key_press(self, event):
        if self.selected_peak_index is not None:
            num_peaks = self.peak_params_grid.GetNumberRows() // 2  # Assuming each peak uses two rows

            if event.key == 'tab':
                if not self.peak_fitting_tab_selected:
                    self.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
                else:
                    self.change_selected_peak(1)  # Move to next peak
                return  # Prevent event from propagating
            elif event.key == 'q':
                # print("Local Key q")
                # self.change_selected_peak(-1)
                pass

            self.highlight_selected_peak()
            self.clear_and_replot()
            self.canvas.draw_idle()

    def on_key_press_global(self, event):
        keycode = event.GetKeyCode()
        if event.AltDown() and self.selected_peak_index is not None:
            if event.ShiftDown() and keycode in [wx.WXK_LEFT, wx.WXK_RIGHT] and self.selected_peak_index is not None:
                row = self.selected_peak_index * 2
                model = self.peak_params_grid.GetCellValue(row, 13)
                delta = 0.1 if keycode == wx.WXK_RIGHT else -0.1
                current_fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
                new_fwhm = current_fwhm

                if model in ["Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)", "Voigt (Area, L/G, \u03c3, S)"]:
                    current_sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                    lg_ratio = float(self.peak_params_grid.GetCellValue(row, 5))
                    new_sigma = max(current_sigma + delta, 0.4)
                    new_gamma = (lg_ratio / 100 * new_sigma) / (1 - lg_ratio / 100)
                    self.peak_params_grid.SetCellValue(row, 7, f"{new_sigma:.2f}")
                    self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
                    if model == "Voigt (Area, L/G, \u03c3, S)":
                        skew = float(self.peak_params_grid.GetCellValue(row, 9))
                elif model == "ExpGauss.(Area, \u03c3, \u03b3)":
                    current_gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                    new_gamma = max(current_gamma + delta, 0.2)
                    self.peak_params_grid.SetCellValue(row, 8, f"{new_gamma:.2f}")
                else:
                    new_fwhm = max(current_fwhm + delta, 0.3)
                    self.peak_params_grid.SetCellValue(row, 4, f"{new_fwhm:.2f}")

                self.recalculate_peak_area(self.selected_peak_index)
                self.update_linked_fwhm_recursive(self.selected_peak_index, new_fwhm)
                self.clear_and_replot()
                self.plot_manager.add_cross_to_peak(self, self.selected_peak_index)
                save_state(self)
                return
            elif keycode in [wx.WXK_LEFT, wx.WXK_RIGHT]:
                row = self.selected_peak_index * 2
                current_position = float(self.peak_params_grid.GetCellValue(row, 2))

                # Move 0.1 eV left or right
                delta = 0.1 if keycode == wx.WXK_LEFT else -0.1
                new_position = current_position + delta

                # Update peak position
                self.peak_params_grid.SetCellValue(row, 2, f"{new_position:.2f}")

                # Update linked peaks if any
                self.update_linked_peaks_recursive(self.selected_peak_index, new_position,
                                                   float(self.peak_params_grid.GetCellValue(row, 3)))

                # Refresh display
                self.clear_and_replot()
                self.plot_manager.add_cross_to_peak(self, self.selected_peak_index)
                save_state(self)
                return
            elif keycode in [wx.WXK_UP, wx.WXK_DOWN]:
                row = self.selected_peak_index * 2
                current_height = float(self.peak_params_grid.GetCellValue(row, 3))
                intensity_factor = 0.05

                delta = intensity_factor * current_height * (1 if keycode == wx.WXK_UP else -1)
                new_height = max(0, current_height + delta)

                self.peak_params_grid.SetCellValue(row, 3, f"{new_height:.2f}")

                # Recalculate area after height change
                self.recalculate_peak_area(self.selected_peak_index)

                self.update_linked_peaks_recursive(self.selected_peak_index,
                                                   float(self.peak_params_grid.GetCellValue(row, 2)), new_height)

                self.clear_and_replot()
                self.plot_manager.add_cross_to_peak(self, self.selected_peak_index)
                save_state(self)
                return
        elif event.ControlDown():
            if keycode == ord('B'):
                self.toggle_energy_scale()
                self.toggle_energy_scale()
                self.clear_and_replot()
            if keycode == ord('Z'):
                from libraries.Save import undo
                undo(self)
                return
            elif keycode == ord('Y'):
                from libraries.Save import redo
                redo(self)
                return
            # elif keycode == ord('C'):
            #     print('Control C')
            #     copy_cell(self.peak_params_grid)
            #     return
            # elif keycode == ord('V'):
            #     paste_cell(self.peak_params_grid)
            #     save_state(self)
            #     return
            elif keycode == ord('S'):
                print("Saving")
                on_save(self)
                return
            elif keycode == ord('O'):
                print("Opening")
                open_xlsx_file(self)
                return
            elif keycode == ord('Q'):
                print("Quiting")
                on_exit(self, event),
                return
            elif keycode == ord('K'):
                print("Shortcuts")
                self.show_popup_message2("Keyboard Shortcuts",
                                        "-Tab: Select next peak\n"
                                        "-Q: Select previous peak\n"
                                        "-Ctrl+Minus (-): Zoom out\n"
                                        "-Ctrl+Equal (=): Zoom in\n"
                                        "-Ctrl+Left bracket [: Select previous core level\n"
                                        "-Ctrl+Right bracket ]: Select next core level\n"
                                        "-Ctrl+Up: Increase plot intensity\n"
                                        "-Ctrl+Down: Decrease plot intensity\n"
                                        "-Ctrl+Left: Move plot to High BE\n"
                                        "-Ctrl+Right: Move plot to Low BE\n"
                                        "-SHIFT+Left: Decrease High BE\n"
                                        "-SHIFT+Right: Increase High BE\n"
                                        "-Ctrl+Z: Undo up to 30 events\n"
                                        "-Ctrl+Y: Redo\n"
                                        "-Ctrl+S: Save. Only works on the grid and not on the figure canvas\n"
                                        "-Ctrl+P: Open peak fitting window\n"
                                        "-Ctrl+A: Open Area window\n"
                                        "-Ctrl+K: Show Keyboard shortcut\n"
                                        "-Alt+Up: Increase peak intensity\n"
                                        "-Alt+Down: Decrease peak intensity\n"
                                        "-Alt+Left: Move peak to High BE\n"
                                        "-Alt+Right: Move peak to Low BE\n"
                                         )
                return
            elif keycode == ord('A'):
                print("Opening Area Window")
                self.on_open_background_window(),
                return
            elif keycode == ord('P'):
                print("Opening Fitting Window")
                self.on_open_fitting_window()
                return
        if event.ShiftDown() and keycode == wx.WXK_LEFT:
            self.adjust_plot_limits('high_be', 'increase')
            return
        elif event.ShiftDown() and keycode == wx.WXK_RIGHT:
            self.adjust_plot_limits('high_be', 'decrease')
            return

        if keycode == wx.WXK_TAB:
            if not self.peak_fitting_tab_selected:
                self.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
            else:
                self.change_selected_peak(1)  # Move to next peak
            return  # Prevent event from propagating
        elif keycode == ord('Q'):
            if not self.peak_fitting_tab_selected:
                self.show_popup_message("Open the Peak Fitting Tab to move or select a peak")
            else:
                self.change_selected_peak(-1)  # Move to next peak
            return  # Prevent event from propagating
        elif keycode in [ord('['), ord(']')] and event.ControlDown():
            current_index = self.sheet_combobox.GetSelection()
            num_sheets = self.sheet_combobox.GetCount()
            if keycode == ord('['):
                new_index = (current_index - 1) % num_sheets
            else:
                new_index = (current_index + 1) % num_sheets

            self.sheet_combobox.SetSelection(new_index)
            new_sheet = self.sheet_combobox.GetString(new_index)

            # Call on_sheet_selected with the new sheet name
            on_sheet_selected(self, new_sheet)
            save_state(self)
            return  # Prevent event from propagating
        elif keycode in [ord('-'), ord('=')] and event.ControlDown():
            sheet_name = self.sheet_combobox.GetValue()
            limits = self.plot_config.get_plot_limits(self, sheet_name)

            zoom_factor = 0.2
            if keycode == ord('-'):  # Zoom out
                limits['Xmin'] -= zoom_factor
                limits['Xmax'] += zoom_factor
            else:  # Zoom in
                limits['Xmin'] += zoom_factor
                limits['Xmax'] -= zoom_factor

            # Update the plot limits
            self.plot_config.update_plot_limits(self, sheet_name,
                                                x_min=limits['Xmin'],
                                                x_max=limits['Xmax'])

            # Update the plot
            self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis

            # Update subplot limits if it exists
            if hasattr(self, 'residuals_subplot') and self.residuals_subplot:
                self.residuals_subplot.set_xlim(limits['Xmax'], limits['Xmin'])

            # After zooming, update residuals
            self.plot_manager.update_overall_fit_and_residuals(self)

            self.canvas.draw_idle()
            return  # Prevent event from propagating
        elif keycode in [wx.WXK_LEFT, wx.WXK_RIGHT] and event.ControlDown():
            sheet_name = self.sheet_combobox.GetValue()
            limits = self.plot_config.get_plot_limits(self, sheet_name)
            move_factor = 0.1
            if keycode == wx.WXK_LEFT:
                limits['Xmin'] -= move_factor
                limits['Xmax'] -= move_factor
            else:  # Right key
                limits['Xmin'] += move_factor
                limits['Xmax'] += move_factor

            # Update the plot limits
            self.plot_config.update_plot_limits(self, sheet_name,
                                                x_min=limits['Xmin'],
                                                x_max=limits['Xmax'])

            # Update the plot
            self.ax.set_xlim(limits['Xmax'], limits['Xmin'])  # Reverse X-axis
            self.plot_manager.update_overall_fit_and_residuals(self)

            self.canvas.draw_idle()
            return
        elif keycode in [wx.WXK_UP, wx.WXK_DOWN] and event.ControlDown():
            sheet_name = self.sheet_combobox.GetValue()
            limits = self.plot_config.get_plot_limits(self, sheet_name)
            intensity_factor = 0.05
            max_intensity = max(self.y_values)

            if keycode == wx.WXK_DOWN:  # Decrease intensity
                limits['Ymax'] = max(limits['Ymax'] - intensity_factor * max_intensity, limits['Ymin'])
            else:  # Increase intensity
                limits['Ymax'] += intensity_factor * max_intensity

            # Update the plot limits
            self.plot_config.update_plot_limits(self, sheet_name, y_max=limits['Ymax'])

            # Update the plot
            self.ax.set_ylim(limits['Ymin'], limits['Ymax'])

            # Check RSD visibility
            if hasattr(self.plot_manager, 'residuals_state') and self.plot_manager.residuals_state == 1:
                residual_height = 1.07 * max(self.y_values)
                if residual_height > limits['Ymax']:
                    if hasattr(self.plot_manager, 'rsd_text') and self.plot_manager.rsd_text:
                        self.plot_manager.rsd_text.remove()
                        self.plot_manager.rsd_text = None
                else:
                    if hasattr(self.plot_manager, 'rsd_text') and self.plot_manager.rsd_text is None:
                        rsd = PeakFunctions.calculate_rsd(self.y_values, self.background)
                        if rsd is not None:
                            x_min = self.ax.get_xlim()[1] + 0.4
                            self.plot_manager.rsd_text = self.ax.text(x_min, residual_height, f'RSD: {rsd:.2f}',
                                                                      horizontalalignment='right',
                                                                      verticalalignment='center',
                                                                      fontsize=9,
                                                                      color=self.plot_manager.residual_color,
                                                                      alpha=self.plot_manager.residual_alpha + 0.2,
                                                                      bbox=dict(facecolor='white', edgecolor='none'))

            self.canvas.draw_idle()
            return  # Prevent event from propagating
        event.Skip()  # Let other key events propagate normally

    def show_popup_message(self, message):
        popup = wx.adv.RichToolTip("Are you trying to select a peak?", message)
        popup.ShowFor(self)

    def show_popup_message2(self, message1, message2):
        popup = wx.adv.RichToolTip(message1, message2)
        popup.ShowFor(self)

    def change_selected_peak(self, direction):


        num_peaks = self.peak_params_grid.GetNumberRows() // 2


        if self.selected_peak_index is None:
            self.selected_peak_index = 0 if direction > 0 else num_peaks - 1
        else:
            self.selected_peak_index = (self.selected_peak_index + direction) % num_peaks



        self.remove_cross_from_peak()
        if self.peak_fitting_tab_selected:
            self.highlight_selected_peak()

        self.canvas.draw_idle()

    def on_motion(self, event):
        if event.button == 1 and event.key == 'shift' and self.background_tab_selected:
            # print('ON MOTION')
            x_click = event.xdata
            sheet_name = self.sheet_combobox.GetValue()
            if self.vline1 is not None and self.vline2 is not None:
                vline1_x = self.vline1.get_xdata()[0]
                vline2_x = self.vline2.get_xdata()[0]

                # Determine which vline is at low/high BE
                low_be_x = min(vline1_x, vline2_x)
                high_be_x = max(vline1_x, vline2_x)

                dist1 = abs(x_click - vline1_x)
                dist2 = abs(x_click - vline2_x)

                if dist1 < dist2:  # Closer to vline1
                    raw_y = self.y_values[np.argmin(np.abs(self.x_values - vline1_x))]
                    if vline1_x == low_be_x:
                        self.offset_l = event.ydata - raw_y
                        self.Data['Core levels'][sheet_name]['Background']['Bkg Offset Low'] = self.offset_l
                        self.fitting_window.offset_l_text.SetValue(f'{self.offset_l:.1f}')
                    else:
                        self.offset_h = event.ydata - raw_y
                        self.Data['Core levels'][sheet_name]['Background']['Bkg Offset High'] = self.offset_h
                        self.fitting_window.offset_h_text.SetValue(f'{self.offset_h:.1f}')
                else:  # Closer to vline2
                    raw_y = self.y_values[np.argmin(np.abs(self.x_values - vline2_x))]
                    if vline2_x == low_be_x:
                        self.offset_l = event.ydata - raw_y
                        self.Data['Core levels'][sheet_name]['Background']['Bkg Offset Low'] = self.offset_l
                        self.fitting_window.offset_l_text.SetValue(f'{self.offset_l:.1f}')
                    else:
                        self.offset_h = event.ydata - raw_y
                        self.Data['Core levels'][sheet_name]['Background']['Bkg Offset High'] = self.offset_h
                        self.fitting_window.offset_h_text.SetValue(f'{self.offset_h:.1f}')

                self.plot_manager.plot_background(self)
        elif event.inaxes and self.moving_vline is not None:
            x_click = event.xdata
            self.moving_vline.set_xdata([x_click])

            sheet_name = self.sheet_combobox.GetValue()
            if sheet_name in self.Data['Core levels']:
                core_level_data = self.Data['Core levels'][sheet_name]

                if self.moving_vline in [self.vline1, self.vline2]:
                    if self.moving_vline == self.vline1:
                        core_level_data['Background']['Bkg Low'] = float(x_click)
                    else:
                        core_level_data['Background']['Bkg High'] = float(x_click)

                    # Ensure Bkg Low is always less than Bkg High
                    bkg_low = core_level_data['Background']['Bkg Low']
                    bkg_high = core_level_data['Background']['Bkg High']
                    core_level_data['Background']['Bkg Low'] = min(bkg_low, bkg_high)
                    core_level_data['Background']['Bkg High'] = max(bkg_low, bkg_high)

                    # If in Adaptive Smart mode, update the background in real-time
                    if self.background_method == "Multi-Regions Smart":
                        # plot_background(self)
                        pass

                elif self.moving_vline in [self.vline3, self.vline4]:
                    if self.moving_vline == self.vline3:
                        self.noise_min_energy = float(x_click)
                    else:
                        self.noise_max_energy = float(x_click)

                    # Ensure noise_min_energy is always less than noise_max_energy
                    self.noise_min_energy, self.noise_max_energy = sorted(
                        [self.noise_min_energy, self.noise_max_energy])

            self.canvas.draw_idle()

    def on_release(self, event):
        if self.moving_vline is not None:
            # Disconnect motion handler when mouse is released
            if hasattr(self, 'motion_notify_id'):
                self.canvas.mpl_disconnect(self.motion_notify_id)
                delattr(self, 'motion_notify_id')
            if hasattr(self, 'button_release_id'):
                self.canvas.mpl_disconnect(self.button_release_id)
                delattr(self, 'button_release_id')

            # self.canvas.mpl_disconnect(self.motion_cid)
            # self.canvas.mpl_disconnect(self.release_cid)
            self.moving_vline = None

            # Ensure the background range is correctly ordered in the data dictionary
            sheet_name = self.sheet_combobox.GetValue()
            if sheet_name in self.Data['Core levels']:
                core_level_data = self.Data['Core levels'][sheet_name]
                if 'Background' in core_level_data:
                    bg_low = core_level_data['Background'].get('Bkg Low')
                    bg_high = core_level_data['Background'].get('Bkg High')
                    if bg_low is not None and bg_high is not None:
                        core_level_data['Background']['Bkg Low'] = min(bg_low, bg_high)
                        core_level_data['Background']['Bkg High'] = max(bg_low, bg_high)

            # If in Adaptive Smart mode, update the background
            if self.background_method == "Multi-Regions Smart":
                # plot_background(self)
                pass
        # If a peak is being moved, update its position and height
        if self.selected_peak_index is not None:
            row = self.selected_peak_index * 2  # Each peak uses two rows in the grid
            peak_x = float(self.peak_params_grid.GetCellValue(row, 2))  # Position
            peak_y = float(self.peak_params_grid.GetCellValue(row, 3))  # Height
            self.update_peak_plot(peak_x, peak_y, remove_old_peaks=False)

            # Update the peak information in the data dictionary
            sheet_name = self.sheet_combobox.GetValue()
            if sheet_name in self.Data['Core levels']:
                core_level_data = self.Data['Core levels'][sheet_name]
                if 'Fitting' not in core_level_data:
                    core_level_data['Fitting'] = {}
                if 'Peaks' not in core_level_data['Fitting']:
                    core_level_data['Fitting']['Peaks'] = {}

                peak_label = self.peak_params_grid.GetCellValue(row, 1)


                def try_float(value, default=0.0):
                    try:
                        return float(value)
                    except ValueError:
                        return default

                core_level_data['Fitting']['Peaks'][peak_label] = {
                    'Position': peak_x,
                    'Height': peak_y,
                    'FWHM': float(self.peak_params_grid.GetCellValue(row, 4)),
                    'L/G': float(self.peak_params_grid.GetCellValue(row, 5)),
                    'Area': try_float(self.peak_params_grid.GetCellValue(row, 6), 0.0),
                    'Sigma': float(self.peak_params_grid.GetCellValue(row, 7)),
                    'Gamma': float(self.peak_params_grid.GetCellValue(row, 8)),
                    'Skew': float(self.peak_params_grid.GetCellValue(row, 9))
                }

        self.selected_peak_index = None
        self.canvas.draw_idle()

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

    def on_right_click(self, event):
        if event.button == 3:  # Right click
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

            # Check clipboard files
            clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_clipboard.json')
            has_clipboard_data = os.path.exists(clipboard_file)

            menu.AppendSeparator()

            copy_peak_table = menu.Append(-1, "Copy Peak Table")
            paste_peak_table = menu.Append(-1, "Paste Peak Table")
            peak_clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')

            has_peak_clipboard_data = os.path.exists(peak_clipboard_file)
            has_rows = self.peak_params_grid.GetNumberRows() > 0

            paste.Enable(has_clipboard_data)
            paste_peak_table.Enable(has_peak_clipboard_data)
            copy_peak_table.Enable(has_rows)

            self.Bind(wx.EVT_MENU, self.on_zoom_in_tool, zoom_in)
            self.Bind(wx.EVT_MENU, self.on_zoom_out, zoom_out)
            self.Bind(wx.EVT_MENU, self.on_drag_tool, drag)
            self.Bind(wx.EVT_MENU, lambda evt: copy_core_level(self), copy)
            self.Bind(wx.EVT_MENU, lambda evt: paste_core_level(self), paste)
            self.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self), copy_peak_table)
            self.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self), paste_peak_table)

            self.PopupMenu(menu)
            menu.Destroy()

    # In KherveFitting.py:

    def on_peak_params_right_click(self, event):
        row = event.GetRow()
        col = event.GetCol()

        menu = wx.Menu()
        copy_item = menu.Append(wx.ID_ANY, "Copy Peak Table")
        paste_item = menu.Append(wx.ID_ANY, "Paste Peak Table")

        # Add separator and propagate option
        menu.AppendSeparator()
        propagate_item = menu.Append(wx.ID_ANY, "Propagate to column")

        # Check if paste data exists
        import os
        import tempfile
        clipboard_file = os.path.join(tempfile.gettempdir(), 'khervefitting_peak_clipboard.json')

        # Enable/disable menu items based on context
        has_clipboard = os.path.exists(clipboard_file)
        has_rows = self.peak_params_grid.GetNumberRows() > 0

        copy_item.Enable(has_rows)
        paste_item.Enable(has_clipboard)
        propagate_item.Enable(col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1)  # Enable only for constraint rows

        from libraries.Save import copy_all_peak_parameters, paste_all_peak_parameters
        from libraries.Utilities import propagate_constraint

        # Bind menu events
        self.Bind(wx.EVT_MENU, lambda evt: copy_all_peak_parameters(self), copy_item)
        self.Bind(wx.EVT_MENU, lambda evt: paste_all_peak_parameters(self), paste_item)
        self.Bind(wx.EVT_MENU, lambda evt: propagate_constraint(self, row, col), propagate_item)

        # Show the menu
        self.peak_params_grid.PopupMenu(menu, event.GetPosition())
        menu.Destroy()

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
        if 'Results' in self.Data and 'Peak' in self.Data['Results']:
            for row in range(self.results_grid.GetNumberRows()):
                peak_label = f"Peak_{row}"
                peak_data = self.Data['Results']['Peak'].get(peak_label)
                if peak_data:
                    checkbox_state = peak_data.get('Checkbox', '0')
                    current_grid_state = self.results_grid.GetCellValue(row, 7)
                    if checkbox_state != current_grid_state:
                        self.results_grid.SetCellValue(row, 7, checkbox_state)
                        self.results_grid.RefreshAttr(row, 7)
        self.results_grid.ForceRefresh()

    def update_checkboxes_from_data_NEW(self):
        if 'Results' in self.Data and 'Peak' in self.Data['Results']:
            for row in range(self.results_grid.GetNumberRows()):
                peak_label = f"Peak_{row}"
                peak_data = self.Data['Results']['Peak'].get(peak_label)
                if peak_data:
                    atomic_conc = float(self.results_grid.GetCellValue(row, 6))
                    checkbox_state = '0' if atomic_conc == 0 else peak_data.get('Checkbox', '0')

                    current_grid_state = self.results_grid.GetCellValue(row, 7)
                    if checkbox_state != current_grid_state:
                        self.results_grid.SetCellValue(row, 7, checkbox_state)
                        self.results_grid.RefreshAttr(row, 7)
                        self.Data['Results']['Peak'][peak_label]['Checkbox'] = checkbox_state
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

    # def on_cell_changed_OLD(self, event):
    #     row = event.GetRow()
    #     col = event.GetCol()
    #
    #     try:
    #         if col in [1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]:
    #             event.Veto()
    #             return
    #         elif col in [8]:  # RSF column
    #             rsf = float(self.results_grid.GetCellValue(row, 8))
    #             area = float(self.results_grid.GetCellValue(row, 5))
    #             binding_energy = float(self.results_grid.GetCellValue(row, 1))
    #             kinetic_energy = self.photons - binding_energy
    #
    #             # Calculate ECF based on method
    #             if self.library_type == "Scofield":
    #                 ecf = kinetic_energy ** 0.6
    #             elif self.library_type == "Wagner":
    #                 ecf = kinetic_energy ** 1.0
    #             elif self.library_type == "TPP-2M":
    #                 # Calculate IMFP using TPP-2M using the average matrix
    #                 imfp = AtomicConcentrations.calculate_imfp_tpp2m(kinetic_energy)
    #
    #                 # 26.2 is a factor added by Avantage to match KE^0.6
    #                 ecf = imfp * 26.2
    #             elif window.library_type == "EAL":
    #                 z_avg = 50  # Default values
    #                 eal = (0.65 + 0.007 * kinetic_energy ** 0.93) / (z_avg ** 0.38)
    #                 ecf = eal
    #
    #             elif self.library_type == "None":
    #                 ecf = 1.0
    #             else:
    #                 ecf = 1.0
    #
    #             txfn = 1.0  # Transmission function because it is corrected in the Excel page
    #             if rsf == 0 or txfn == 0 or ecf == 0:
    #                 new_rel_area = 0
    #             else:
    #                 new_rel_area = area / (rsf * txfn * ecf)
    #             self.results_grid.SetCellValue(row, 13, f"{new_rel_area:.2f}")
    #
    #             # Update the atomic percentages
    #             self.update_atomic_percentages()
    #         elif col in [16, 17, 18, 19]:  # Constraint columns
    #             pass
    #
    #     except ValueError:
    #         wx.MessageBox("Invalid value entered", "Error", wx.OK | wx.ICON_ERROR)

    def on_cell_changed(self, event):
        return

    def update_atomic_percentages(self):
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
                print(f"Library Type:{self.library_type}")

                # Calculate ECF based on method selected
                if self.library_type == "Scofield":
                    ecf = kinetic_energy ** 0.6
                elif self.library_type == "Wagner":
                    ecf = kinetic_energy ** 1.0
                elif self.library_type == "TPP-2M":
                    # Calculate IMFP using TPP-2M using the average matrix
                    imfp = AtomicConcentrations.calculate_imfp_tpp2m(kinetic_energy)

                    # 26.2 is a factor added by Avantage to match KE^0.6
                    ecf = imfp * 26.2
                elif self.library_type == "None":
                    ecf = 1.0
                else:
                    ecf = 1.0  # Default no correction

                # Get raw area and RSF
                area = float(self.results_grid.GetCellValue(i, 5))
                rsf = float(self.results_grid.GetCellValue(i, 8))
                if rsf == 0:
                    self.results_grid.SetCellValue(i, 6, "0.00")
                    continue


                # Calculate Transmission function
                txfn = 1.0  # Transmission function

                # Angular correction
                angular_correction = 1.0
                if self.use_angular_correction:
                    angular_correction = AtomicConcentrations.calculate_angular_correction(
                        self,
                        peak_name,
                        self.analysis_angle
                    )

                # Calculate normalized area with ECF correction
                normalized_area = area / (rsf * txfn * ecf * angular_correction)

                total_normalized_area += normalized_area
                checked_indices.append((i, normalized_area))
            else:
                # Set the atomic percentage to 0 for unticked rows
                self.results_grid.SetCellValue(i, 6, "0.00")

        # Calculate atomic percentages
        for i, norm_area in checked_indices:
            atomic_percent = (norm_area / total_normalized_area) * 100 if total_normalized_area > 0 else 0
            self.results_grid.SetCellValue(i, 6, f"{atomic_percent:.2f}")

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
            if sigma is None or gamma is None:
                raise ValueError("Sigma and gamma are required for Voigt models")
            area = PeakFunctions.skewedvoigt_height_to_area(height, sigma / 2.355, gamma / 2, skew)
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


    def on_peak_params_cell_changed(self, event):
        row = event.GetRow()
        col = event.GetCol()
        new_value = self.peak_params_grid.GetCellValue(row, col)
        sheet_name = self.sheet_combobox.GetValue()
        peak_index = row // 2

        # Define default constraint values
        default_constraints = {
            2: '0:1000',  # Position
            3: '1:1e7',  # Height
            4: '0.3:3.5',  # FWHM
            5: '5:80',  # L/G
            6: '1:1e7',  # Area
            7: '0.3:3',  # Sigma
            8: '0.3:3',  # Gamma
            9: '0.01:2' # Skew
        }
        # Check if this is a constraint row and the value contains "="
        if row % 2 == 1 and col in [2, 3, 4, 5, 6, 7, 8, 9] and "=" in new_value:
            # Remove the "=" from the cell
            self.peak_params_grid.SetCellValue(row, col, new_value.replace("=", ""))

            # Import and call the propagate_constraint function
            from libraries.Utilities import propagate_constraint
            propagate_constraint(self, row, col)
            return  # Skip the rest of the function
        elif col == 1:  # Peak label column
            # Check for duplicate names
            existing_names = []
            for i in range(0, self.peak_params_grid.GetNumberRows(), 2):
                if i != row:  # Skip current row
                    existing_names.append(self.peak_params_grid.GetCellValue(i, 1))

            if new_value in existing_names:
                wx.MessageBox(f"Peak name '{new_value}' already exists. Cannot have duplicate peak names.",
                              "Duplicate Peak Name", wx.OK | wx.ICON_ERROR)
                event.Veto()
                return
        elif col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 1:  # Constraint rows

            if new_value.lower() in ['fi', 'fix', 'fixe', 'fixed']:
                new_value = 'Fixed'
                self.peak_params_grid.SetCellValue(row, col, new_value)
                return
            elif new_value == 'F':
                new_value = 'F*1'
                self.peak_params_grid.SetCellValue(row, col, new_value)
                return
            elif new_value.startswith('#'):
                # Check if '#' is not followed by at least one digit
                if len(new_value) == 1 or not new_value[1:].replace('.', '', 1).isdigit():
                    wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()  # Veto the event if '#' is on its own or not followed by a valid number
                    return

                # If '#' is followed by a valid number, proceed with the calculation
                peak_value = float(self.peak_params_grid.GetCellValue(row - 1, col))
                new_value = str(round(peak_value - float(new_value[1:]), 2)) + ':' + str(
                    round(peak_value + float(new_value[1:]), 2))
                self.peak_params_grid.SetCellValue(row, col, new_value)
                return
            elif new_value.startswith(('+', '/', '*')):
                wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                              wx.OK | wx.ICON_ERROR)
                event.Veto()
                return
            # Add validation for expressions starting with '-' and not containing ':'
            elif new_value.startswith('-') and ':' not in new_value:
                wx.MessageBox(f"Wrong Value entered", "Wrong Value",
                              wx.OK | wx.ICON_ERROR)
                event.Veto()
                return

            # Pattern to match all possible formats
            pattern = r'^([A-P])([+\-*/])(\d+\.?\d*)(?:#(\d+\.?\d*))?$'
            match = re.match(pattern, new_value)

            if not new_value:  # If empty string
                if col == 2:
                    sheet_name = self.sheet_combobox.GetValue()
                    x_values = self.Data['Core levels'][sheet_name]['B.E.']
                    new_value = f"{min(x_values):.2f}:{max(x_values):.2f}"
                elif col == 4:
                    new_value = "0.3:35"
                elif col == 5:
                    new_value = "1:80"
                elif col == 6:
                    new_value = "1:1e7"
                elif col == 7:
                    new_value = "0.2:3"
                elif col == 8:
                    new_value = "0.2:3"
                elif col == 9:
                    new_value = "0.1:1"
                self.peak_params_grid.SetCellValue(row, col, new_value)
                return

            if match:
                referenced_peak = match.group(1)
                letter_index = ord(referenced_peak) - 65
                current_peak = row // 2

                if letter_index == current_peak:
                    wx.MessageBox(f"Cannot reference the same peak ({referenced_peak}).", "Invalid Self Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                if letter_index * 2 >= self.peak_params_grid.GetNumberRows():
                    wx.MessageBox(f"Peak {referenced_peak} does not exist.", "Invalid Peak Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

            elif new_value.upper() in 'ABCDEFGHIJKLMNOP':
                letter_index = ord(new_value.upper()) - 65
                current_peak = row // 2

                if letter_index == current_peak:
                    wx.MessageBox(f"Cannot reference the same peak ({new_value.upper()}).", "Invalid Self Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                if letter_index * 2 >= self.peak_params_grid.GetNumberRows():
                    wx.MessageBox(f"Peak {new_value.upper()} does not exist.", "Invalid Peak Reference",
                                  wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return
        elif col in [0, 10, 11, 12]:
            event.Veto()
            return
        elif col not in [13, 14] and row % 2 == 1:  # Constraint row
            if not new_value:  # If the cell is empty
                new_value = default_constraints.get(col, '')
                self.peak_params_grid.SetCellValue(row, col, new_value)
        # Allow only numeric input for specific columns in non-constraint rows
        elif col not in [1, 13, 14] and row % 2 == 0:
            try:
                float(new_value)  # This will handle integers, floats, and scientific notation
            except ValueError:
                event.Veto()
                return


        if col == 2 and new_value.upper() in 'ABCDEFGHIJKLMNOP':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                new_value = "0:1000"  # Default value
            else:
                current_split = float(self.peak_params_grid.GetCellValue(row-1, 12))
                print(f'Current split: {current_split}')
                ref_split = float(self.peak_params_grid.GetCellValue(letter_index * 2, 12))
                split_diff = current_split - ref_split
                if split_diff != 0:
                    new_value = f"{new_value.upper()}+{split_diff:.2f}#0.1"
                else:
                    new_value = new_value.upper() + '*1'
        elif col == 6 and new_value.upper() in 'ABCDEFGHIJKLMNOP':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                new_value = "1:1e7"  # Default value
            else:
                current_ratio = float(self.peak_params_grid.GetCellValue(row - 1, 11))
                ref_ratio = float(self.peak_params_grid.GetCellValue(letter_index * 2, 11))
                ratio = current_ratio / ref_ratio
                if ratio != 1:
                    new_value = f"{new_value.upper()}*{ratio:.2f}#0.01"
                else:
                    new_value = new_value.upper() + '*1'
        elif new_value.lower() in 'abcdefghijklmnop':
            letter_index = ord(new_value.upper()) - 65
            if letter_index * 2 == row - 1:  # Same peak
                if col == 2:
                    new_value = "0:1000"
                elif col ==4:
                    new_value = "0.3:35"
                elif col ==5:
                    new_value = "1:80"
                elif col ==6:
                    new_value = "1:1e7"
                elif col ==7:
                    new_value = "0.2:3"
                elif col ==8:
                    new_value = "0.2:3"
                elif col ==8:
                    new_value = "0.1:1"
            else:
                new_value = new_value.upper() + '*1'
            self.peak_params_grid.SetCellValue(row, col, new_value)

        # Convert lowercase to uppercase in expressions like a*0.5
        if '*' in new_value or '+' in new_value:
            parts = new_value.split('*' if '*' in new_value else '+')
            if len(parts) == 2 and parts[0].lower() in 'abcdefghij':
                parts[0] = parts[0].upper()
                new_value = ('*' if '*' in new_value else '+').join(parts)
                self.peak_params_grid.SetCellValue(row, col, new_value)

        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            peak_keys = list(peaks.keys())

            if peak_index < len(peak_keys):
                correct_peak_key = peak_keys[peak_index]

                if row % 2 == 0:  # Main parameter row
                    if col == 1:  # Label
                        # Update the label while preserving order
                        new_peaks = {}
                        for i, (key, value) in enumerate(peaks.items()):
                            if i == peak_index:
                                new_peaks[new_value] = value
                            else:
                                new_peaks[key] = value
                        self.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks
                    elif col == 2:  # Position
                        peaks[correct_peak_key]['Position'] = float(new_value)
                    elif col in [3, 4, 5, 6, 7, 8,9]:  # Height, FWHM, L/G, Area, Sigma, Gamma changed
                        def try_float(value, default=0.0):
                            try:
                                return float(value)
                            except (ValueError, TypeError):
                                return default


                        model = peaks[correct_peak_key]['Fitting Model']
                        height = float(self.peak_params_grid.GetCellValue(row, 3))
                        fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
                        fraction = float(self.peak_params_grid.GetCellValue(row, 5))
                        area = float(self.peak_params_grid.GetCellValue(row, 6))
                        sigma = try_float(self.peak_params_grid.GetCellValue(row, 7), 0.0)
                        gamma = try_float(self.peak_params_grid.GetCellValue(row, 8), 0.0)
                        skew = float(self.peak_params_grid.GetCellValue(row, 9))

                        if model in ["LA (Area, \u03c3/\u03b3, \u03b3)", "LA*G (Area, \u03c3/\u03b3, \u03b3)"]:
                            if col == 5:  # L/G ratio changed
                                gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                            elif col == 7:  # Sigma changed
                                fraction = 100 * sigma / (sigma + gamma)
                                self.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                            elif col == 8:  # Gamma changed
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                            elif col == 9:
                                pass
                        elif model in ["Voigt (Area, L/G, \u03c3)"]:
                            if col == 5: # L/G ratio changed
                                sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                            elif col == 7:  # Sigma changed
                                fraction = float(self.peak_params_grid.GetCellValue(row, 5))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                            elif col == 8:
                                pass
                        elif model in ["Voigt (Area, L/G, \u03c3, S)"]:
                            if col == 5: # L/G ratio changed
                                sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                            elif col == 7:  # Sigma changed
                                fraction = float(self.peak_params_grid.GetCellValue(row, 5))
                                gamma = (fraction / 100) * sigma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 8, f"{gamma:.2f}")
                            elif col == 8:
                                pass
                            elif col == 9:
                                value = float(self.peak_params_grid.GetCellValue(row, 9))
                                if value == 0:
                                    skew = 0.1
                                    self.peak_params_grid.SetCellValue(row, 9, f"{skew:.2f}")
                        elif model in ["LA (Area, \u03c3, \u03b3)"]:
                            if col == 5:  # L/G ratio changed
                                pass
                            elif col == 7:  # Sigma changed
                                fraction = 100 * sigma / (sigma + gamma)
                                self.peak_params_grid.SetCellValue(row, 5, f"{fraction:.2f}")
                            elif col == 8:  # Gamma changed
                                sigma = (fraction / 100) * gamma / (1 - fraction / 100)
                                self.peak_params_grid.SetCellValue(row, 7, f"{sigma:.2f}")
                            elif col == 9:
                                pass
                            elif model in ["GL (Area)"]:
                                # For Gaussian-Lorentzian area-based model
                                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                                height = area / (sigma * np.sqrt(2 * np.pi))
                                self.peak_params_grid.SetCellValue(row, 3, f"{height:.2f}")
                            elif model in ["SGL (Area)"]:
                                # For Sum of Gaussian-Lorentzian area-based model
                                sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))
                                gamma = fwhm / 2
                                height = area / ((1 - fraction / 100) * sigma * np.sqrt(2 * np.pi) + (
                                            fraction / 100) * np.pi * gamma)
                                self.peak_params_grid.SetCellValue(row, 3, f"{height:.2f}")
                        elif model == "D-parameter":
                            return
                        else:
                            # Recalculate area
                            area = self.calculate_peak_area(model, height, fwhm, fraction, sigma, gamma,skew)
                            self.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")

                        self.update_ratios()
                        # Update grid and data
                        peaks[correct_peak_key].update({
                            'Height': round(height, 2),
                            'FWHM': round(fwhm, 2),
                            'L/G': round(fraction, 2),
                            'Area': round(area, 2),
                            'Sigma': round(sigma, 2),
                            'Gamma': round(gamma, 2),
                            'Skew': round(skew, 2)
                        })
                    elif col == 13:  # Fitting Model changed
                        peaks[correct_peak_key]['Fitting Model'] = new_value

                        # Recalculate area with new model
                        model = peaks[correct_peak_key]['Fitting Model']
                        height = float(self.peak_params_grid.GetCellValue(row, 3))
                        fwhm = float(self.peak_params_grid.GetCellValue(row, 4))
                        fraction = float(self.peak_params_grid.GetCellValue(row, 5))
                        sigma = float(self.peak_params_grid.GetCellValue(row, 7))
                        gamma = float(self.peak_params_grid.GetCellValue(row, 8))
                        skew = float(self.peak_params_grid.GetCellValue(row, 9))
                        area = self.calculate_peak_area(new_value, height, fwhm, fraction, sigma, gamma)

                        self.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")
                        peaks[correct_peak_key]['Area'] = area

                        on_sheet_selected(self, sheet_name)
                else:  # Constraint row
                    if col in [2, 3, 4, 5, 6, 7, 8, 9]:
                        constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
                        column_to_constraint = {2: 0, 3: 1, 4: 2, 5: 3, 6: 4, 7: 5, 8: 6, 9:7}
                        constraint_key = constraint_keys[column_to_constraint[col]]

                        if 'Constraints' not in peaks[correct_peak_key]:
                            peaks[correct_peak_key]['Constraints'] = {}

                        # Set default constraints if cell is empty
                        if not new_value:
                            default_constraints = {
                                'Position': '0:1000',
                                'Height': '1:1e7',
                                'FWHM': '0.3:3.5',
                                'L/G': '5:80',
                                'Area': '1:1e7',
                                'Sigma': '0.01:4',
                                'Gamma': '0.01:4',
                                'Skew': '0.02:2'
                            }
                            new_value = default_constraints[col]
                            self.peak_params_grid.SetCellValue(row, col, new_value)

                        peaks[correct_peak_key]['Constraints'][constraint_key] = new_value

            # Ensure numeric values are displayed with 2 decimal places
        if col in [2, 3, 4, 5, 6, 7, 8, 9] and row % 2 == 0:  # Only for main parameter rows, not constraint rows
            try:
                formatted_value = f"{float(new_value):.2f}"
                self.peak_params_grid.SetCellValue(row, col, formatted_value)
            except ValueError:
                pass

        event.Skip()

        # Update all data in window.Data for all peaks
        for i in range(self.peak_params_grid.GetNumberRows() // 2):
            row_data = i * 2  # Data row
            row_constraint = i * 2 + 1  # Constraint row
            peak_label = self.peak_params_grid.GetCellValue(row_data, 1)

            # Update main parameters
            peaks[peak_label] = {
                'Position': float(self.peak_params_grid.GetCellValue(row_data, 2)),
                'Height': float(self.peak_params_grid.GetCellValue(row_data, 3)),
                'FWHM': float(self.peak_params_grid.GetCellValue(row_data, 4)),
                'L/G': float(self.peak_params_grid.GetCellValue(row_data, 5)),
                'Area': float(self.peak_params_grid.GetCellValue(row_data, 6)),
                'Sigma': float(self.peak_params_grid.GetCellValue(row_data, 7)),
                'Gamma': float(self.peak_params_grid.GetCellValue(row_data, 8)),
                'Skew': float(self.peak_params_grid.GetCellValue(row_data, 9)),
                'Fitting Model': self.peak_params_grid.GetCellValue(row_data, 13)
            }

            # Update constraints
            if 'Constraints' not in peaks[peak_label]:
                peaks[peak_label]['Constraints'] = {}

            constraint_keys = ['Position', 'Height', 'FWHM', 'L/G', 'Area', 'Sigma', 'Gamma', 'Skew']
            for col_idx, key in enumerate(constraint_keys, start=2):
                value = self.peak_params_grid.GetCellValue(row_constraint, col_idx)
                peaks[peak_label]['Constraints'][key] = value

        # Refresh the grid to ensure it reflects the current state of self.Data
        self.refresh_peak_params_grid()

        # Replot the peaks with updated parameters
        self.clear_and_replot()

        save_state(self)


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
    def refresh_peak_params_grid(self):
        sheet_name = self.sheet_combobox.GetValue()
        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            for i, (peak_label, peak_data) in enumerate(peaks.items()):
                row = i * 2
                self.peak_params_grid.SetCellValue(row, 1, peak_label)
                self.peak_params_grid.SetCellValue(row, 2, f"{peak_data['Position']:.2f}")
                self.peak_params_grid.SetCellValue(row, 3, f"{peak_data['Height']:.2f}")
                self.peak_params_grid.SetCellValue(row, 4, f"{peak_data['FWHM']:.2f}")
                self.peak_params_grid.SetCellValue(row, 5, f"{peak_data['L/G']:.2f}")
                try:
                    area_value = float(peak_data['Area'])
                    self.peak_params_grid.SetCellValue(row, 6, f"{area_value:.2f}")
                except (ValueError, KeyError):
                    self.peak_params_grid.SetCellValue(row, 6, "ER! REFRESH PEAK")
                # self.peak_params_grid.SetCellValue(row, 6, f"{float(peak_data['Area']):.2f}")
                # self.peak_params_grid.SetCellValue(row, 6, f"{peak_data['Area']:.2f}")
                self.peak_params_grid.SetCellValue(row, 7, f"{peak_data['Sigma']:.2f}")
                self.peak_params_grid.SetCellValue(row, 8, f"{peak_data['Gamma']:.2f}")
                self.peak_params_grid.SetCellValue(row, 9, f"{peak_data.get('Skew', 0):.2f}")
                if 'Constraints' in peak_data:
                    self.peak_params_grid.SetCellValue(row + 1, 2, str(peak_data['Constraints'].get('Position', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 3, str(peak_data['Constraints'].get('Height', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 4, str(peak_data['Constraints'].get('FWHM', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 5, str(peak_data['Constraints'].get('L/G', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 6, str(peak_data['Constraints'].get('Area', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 7, str(peak_data['Constraints'].get('Sigma', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 8, str(peak_data['Constraints'].get('Gamma', '')))
                    self.peak_params_grid.SetCellValue(row + 1, 9, str(peak_data['Constraints'].get('Skew', '')))
        self.peak_params_grid.ForceRefresh()

    def refresh_peak_params_grid_release(self):
        sheet_name = self.sheet_combobox.GetValue()
        if sheet_name in self.Data['Core levels'] and 'Fitting' in self.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            for i, (peak_label, peak_data) in enumerate(peaks.items()):
                row = i * 2
                self.peak_params_grid.SetCellValue(row, 2, f"{peak_data['Position']:.2f}")
                self.peak_params_grid.SetCellValue(row, 3, f"{peak_data['Height']:.2f}")
                try:
                    area_value = float(peak_data['Area'])
                    self.peak_params_grid.SetCellValue(row, 6, f"{area_value:.2f}")
                except (ValueError, KeyError):
                    self.peak_params_grid.SetCellValue(row, 6, "ER! REFRESH PEAK")
            self.peak_params_grid.ForceRefresh()


    def on_checkbox_update(self, event):
        row = event.GetRow()
        col = event.GetCol()

        if col == 7:
            try:
                # Force current_value to be either '0' or '1'
                current_value = self.results_grid.GetCellValue(row, col)
                current_value = '1' if current_value == '1' else '0'

                new_value = '0' if current_value == '1' else '1'

                self.results_grid.SetCellEditor(row, col, wx.grid.GridCellBoolEditor())
                self.results_grid.SetCellRenderer(row, col, wx.grid.GridCellBoolRenderer())
                self.results_grid.SetCellValue(row, col, new_value)

                peak_label = f"Peak_{row}"
                if 'Results' in self.Data and 'Peak' in self.Data['Results'] and peak_label in self.Data['Results'][
                    'Peak']:
                    self.Data['Results']['Peak'][peak_label]['Checkbox'] = new_value

                self.update_atomic_percentages()
                self.results_grid.RefreshAttr(row, col)
                # self.results_grid.ForceRefresh()

                # Simulate plot click behavior
                self.update_checkboxes_from_data()
                self.plot_manager.clear_and_replot(self)
                self.canvas.draw_idle()
                save_state(self)

            except Exception as e:
                print(f"Error updating checkbox: {e}")

        event.Skip()

    def on_key_down(self, event):
        keycode = event.GetKeyCode()
        if keycode == wx.WXK_DELETE:
            selected_rows = self.get_selected_rows()
            if selected_rows:
                selected_rows.sort(reverse=True)
                for row in selected_rows:
                    peak_name = self.results_grid.GetCellValue(row, 0)  # Get the peak name from the first column
                    self.results_grid.DeleteRows(row)

                    # Remove the peak from self.Data
                    for key, value in list(self.Data['Results']['Peak'].items()):
                        if value.get('Name') == peak_name:
                            del self.Data['Results']['Peak'][key]
                            break

                # Renumber the remaining peaks in self.Data
                new_data = {}
                for i, (key, value) in enumerate(self.Data['Results']['Peak'].items()):
                    new_key = f"Peak_{i}"
                    new_data[new_key] = value
                self.Data['Results']['Peak'] = new_data

                self.results_grid.ForceRefresh()
                self.update_atomic_percentages()
                save_state(self)
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





    def deselect_all_peaks(self):
        self.selected_peak_index = None
        self.remove_cross_from_peak()

        if hasattr(self, 'peak_letter') and self.peak_letter:
            self.peak_letter.remove()
            self.peak_letter = None

        # Clear any selections in the peak_params_grid
        self.peak_params_grid.ClearSelection()

        # If you want to uncheck any checkboxes in the results_grid
        for row in range(self.results_grid.GetNumberRows()):
            self.results_grid.SetCellValue(row, 7, '0')  # Assuming column 7 is the checkbox column

        self.update_peak_plot(None, None, remove_old_peaks=True)

        # Refresh both grids
        self.peak_params_grid.ForceRefresh()
        self.results_grid.ForceRefresh()



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
        if grid_fitting_method != "D-parameter":
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

                        for collection in self.ax.collections:
                            if collection.get_label().startswith(sheet_name):
                                data['individual_peak_fits'].append(collection.get_paths()[0].vertices[:, 1])

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

        else:
            config = {}
            print("No config file found, using default values.")
            self.recent_files = []

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

            #Tines opened
            'times_opened': getattr(self, 'times_opened', 0),
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

    # In KherveFitting.py, modify on_be_correction_change:
    def on_be_correction_change(self, event):
        new_correction = self.be_correction_spinbox.GetValue()

        # Find which row the current sheet belongs to
        sheet_name = self.sheet_combobox.GetValue()
        sample_row = None

        if hasattr(self, 'file_manager') and self.file_manager is not None:
            try:
                grid = self.file_manager.grid
                if grid and grid.IsShown():
                    for row in range(grid.GetNumberRows()):
                        for col in range(1, grid.GetNumberCols() - 2):
                            if grid.GetCellValue(row, col) == sheet_name:
                                sample_row = str(row)

                                # Update the BE correction in Data
                                if 'BeCorrections' not in self.Data:
                                    self.Data['BeCorrections'] = {}
                                self.Data['BeCorrections'][sample_row] = new_correction

                                # Update file manager grid
                                be_col = len(self.file_manager.core_levels) + 1
                                grid.SetCellValue(row, be_col, str(new_correction))
                                break
                        if sample_row:
                            break
            except (RuntimeError, wx.PyDeadObjectError):
                pass

        # Apply correction to the current sheet
        self.apply_be_correction(new_correction)

    def on_auto_be(self, event):
        c1s_correction = self.calculate_c1s_correction()
        if c1s_correction is not None:
            self.be_correction_spinbox.SetValue(c1s_correction)
            self.apply_be_correction(c1s_correction)

    def apply_be_correction_SINGLEBE(self, correction):
        """
        Apply binding energy correction to all data and update displays.

        Args:
            correction (float): New BE correction value in eV
        """
        delta_correction = correction - self.be_correction
        self.be_correction = correction
        self.Data['BEcorrection'] = correction


        # Update all sheets
        for sheet_name, sheet_data in self.Data['Core levels'].items():
            # Update B.E. values handling negative numbers
            try:
                sheet_data['B.E.'] = [float(str(be).strip()) + delta_correction for be in sheet_data['B.E.']]
            except (ValueError, TypeError):
                print(f"Warning: Invalid BE data in sheet {sheet_name}")
                continue

            # Update Background
            if 'Background' in sheet_data:
                if 'Bkg Low' in sheet_data['Background'] and sheet_data['Background']['Bkg Low'] != '':
                    sheet_data['Background']['Bkg Low'] += delta_correction
                if 'Bkg High' in sheet_data['Background'] and sheet_data['Background']['Bkg High'] != '':
                    sheet_data['Background']['Bkg High'] += delta_correction

            # Update peak positions
            if 'Fitting' in sheet_data and 'Peaks' in sheet_data['Fitting']:
                for peak in sheet_data['Fitting']['Peaks'].values():
                    peak['Position'] += delta_correction
                    if 'Constraints' in peak:
                        pos_constraint = peak['Constraints'].get('Position', '')
                        if pos_constraint and ',' in pos_constraint and not any(
                                c in pos_constraint for c in 'ABCDEFGHIJKLMNOP'):
                            min_val, max_val = map(float, pos_constraint.split(','))
                            peak['Constraints'][
                                'Position'] = f"{min_val + delta_correction:.2f},{max_val + delta_correction:.2f}"

        # Update plot limits
        for sheet_name in self.Data['Core levels']:
            if sheet_name in self.plot_config.plot_limits:
                limits = self.plot_config.plot_limits[sheet_name]
                limits['Xmin'] += delta_correction
                limits['Xmax'] += delta_correction

                # Also store in main Data structure
                if 'Plot_Limits' not in self.Data['Core levels'][sheet_name]:
                    self.Data['Core levels'][sheet_name]['Plot_Limits'] = {}
                self.Data['Core levels'][sheet_name]['Plot_Limits'] = limits.copy()

        # Update Results grid
        for row in range(self.results_grid.GetNumberRows()):
            try:
                pos = float(self.results_grid.GetCellValue(row, 1))
                self.results_grid.SetCellValue(row, 1, f"{pos + delta_correction:.2f}")
            except ValueError:
                continue

        # Save state for undo/redo
        save_state(self)

        # Update current sheet display
        current_sheet = self.sheet_combobox.GetValue()
        on_sheet_selected(self, current_sheet)

        # # Update plots
        # self.plot_manager.update_plots_be_correction(self, delta_correction)

        print(f"Applied BE correction of {correction:.2f} eV")

    def apply_be_correction(self, correction):
        """
        Apply binding energy correction to all data and update displays.

        Args:
            correction (float): New BE correction value in eV
        """
        delta_correction = correction - self.be_correction
        self.be_correction = correction
        self.Data['BEcorrection'] = correction

        # Find which sample/row this sheet belongs to
        sheet_name = self.sheet_combobox.GetValue()
        sample_row = None

        # Search grid in FileManager if it exists
        if hasattr(self, 'file_manager') and self.file_manager is not None:
            try:
                grid = self.file_manager.grid
                if grid and grid.IsShown():
                    for row in range(grid.GetNumberRows()):
                        for col in range(1, grid.GetNumberCols() - 2):
                            if grid.GetCellValue(row, col) == sheet_name:
                                sample_row = str(row)

                                # Update BE correction in grid
                                be_col = len(self.file_manager.core_levels) + 1
                                grid.SetCellValue(row, be_col, f"{correction:.2f}")
                                break
                        if sample_row:
                            break
            except (RuntimeError, wx.PyDeadObjectError):
                pass

        # Update BeCorrections for this sample if found
        if sample_row is not None and 'BeCorrections' in self.Data:
            self.Data['BeCorrections'][sample_row] = correction

        # Update current sheet
        if sheet_name in self.Data['Core levels']:
            sheet_data = self.Data['Core levels'][sheet_name]

            # Update B.E. values
            sheet_data['B.E.'] = [be + delta_correction for be in sheet_data['B.E.']]

            # Update Background range
            if 'Background' in sheet_data:
                if 'Bkg Low' in sheet_data['Background'] and sheet_data['Background']['Bkg Low'] != '':
                    sheet_data['Background']['Bkg Low'] += delta_correction
                if 'Bkg High' in sheet_data['Background'] and sheet_data['Background']['Bkg High'] != '':
                    sheet_data['Background']['Bkg High'] += delta_correction

            # Update peak positions
            if 'Fitting' in sheet_data and 'Peaks' in sheet_data['Fitting']:
                for peak in sheet_data['Fitting']['Peaks'].values():
                    peak['Position'] += delta_correction
                    if 'Constraints' in peak:
                        pos_constraint = peak['Constraints'].get('Position', '')
                        if pos_constraint and ',' in pos_constraint and not any(
                                c in pos_constraint for c in 'ABCDEFGHIJKLMNOP'):
                            min_val, max_val = map(float, pos_constraint.split(','))
                            peak['Constraints'][
                                'Position'] = f"{min_val + delta_correction:.2f},{max_val + delta_correction:.2f}"

        # Update plot limits
        if sheet_name in self.plot_config.plot_limits:
            limits = self.plot_config.plot_limits[sheet_name]
            limits['Xmin'] += delta_correction
            limits['Xmax'] += delta_correction

        # Update peak_params_grid with corrected positions
        num_peaks = self.peak_params_grid.GetNumberRows() // 2
        for i in range(num_peaks):
            row = i * 2
            try:
                pos = float(self.peak_params_grid.GetCellValue(row, 2))
                self.peak_params_grid.SetCellValue(row, 2, f"{pos + delta_correction:.2f}")
            except ValueError:
                continue

        # Update Results grid
        for row in range(self.results_grid.GetNumberRows()):
            try:
                pos = float(self.results_grid.GetCellValue(row, 1))
                self.results_grid.SetCellValue(row, 1, f"{pos + delta_correction:.2f}")
            except ValueError:
                continue

        # Update current sheet display
        on_sheet_selected(self, sheet_name)

        # Update FileManager's BE corrections if open
        if hasattr(self, 'file_manager') and self.file_manager is not None:
            try:
                self.file_manager.save_be_corrections()
            except Exception as e:
                print(f"Error updating FileManager: {e}")

    def calculate_c1s_correction(self):
        ref_sheet = next((sheet for sheet in self.Data['Core levels'] if self.ref_peak_name[0] in sheet), None)
        if ref_sheet:
            peaks = self.Data['Core levels'][ref_sheet]['Fitting']['Peaks']
            ref_peak = next((peak for label, peak in peaks.items() if self.ref_peak_name in label), None)
            if ref_peak:
                return self.ref_peak_be - ref_peak['Position']
        return None

    def load_be_correction(self):
        if 'BEcorrection' in self.Data:
            self.be_correction = self.Data['BEcorrection']
            self.be_correction_spinbox.SetValue(self.be_correction)

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

    def on_open_file_manager_OLD(self, event):
        """Open the file manager window"""
        from libraries.FileManager import FileManagerWindow
        if not hasattr(self, 'file_manager_window') or not self.file_manager_window:
            self.file_manager_window = FileManagerWindow(self)
        self.file_manager_window.Show()
        self.file_manager_window.Raise()

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

    # Create Splash Screen
    splash = show_splash(duration=2000, delay=0)

    # Detect OS
    os_name = platform.system()

    if os_name == "Darwin":  # Mac OS
        frame = MyFrame(None, "KherveFitting-v1.42 25c18")
    elif os_name == "Windows":
        frame = MyFrame(None, "KherveFitting v1.42 25c18")
    else:
        frame = MyFrame(None, "KherveFitting v1.42")
    frame.Show(True)

    if splash:
        splash.Destroy()

    check_first_time_use(frame)

    updater = UpdateChecker()
    updater.check_update_delayed(frame)

    app.MainLoop()
    sys.exit(0)

