# In libraries/Widgets_Toolbars.py
import os
import sys
import wx
import webbrowser
import subprocess
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wxagg import NavigationToolbar2WxAgg as NavigationToolbar
from libraries.Sheet_Operations import CheckboxRenderer, on_sheet_selected
from libraries.Open import ExcelDropTarget, open_xlsx_file, open_avg_file, import_mrs_file, import_multiple_mrs_files, \
    open_vg_microtech_file_dialog
from libraries.Plot_Operations import PlotManager
from Functions import toggle_Col_1
from libraries.Save import update_undo_redo_state
from libraries.Save import save_state, undo, redo
from libraries.Save import save_peaks_library, load_peaks_library
from libraries.Save import on_save_as, save_plot_only_to_excel
from libraries.Save import export_sheet_to_txt, export_sheet_to_csv, export_sheet_to_dat
from libraries.Open import open_vg_microtech_file, import_multiple_vg_microtech_files
from libraries.Open import open_vamas_file_dialog, open_kal_file_dialog, import_mrs_file, open_spe_file_dialog, open_file_location
from libraries.Open import import_raman_txt_file, import_multiple_raman_files, import_xps_asc_file, import_multiple_xps_asc_files
from libraries.Open import import_xps_csv_file, import_multiple_xps_csv_files
from libraries.Export import export_word_report
from libraries.Utilities import CropWindow, PlotModWindow, on_delete_sheet, copy_sheet, JoinSheetsWindow
from libraries.MarketResearch import check_registration_needed, launch_registration_form
from libraries.Help import show_libraries_used, show_version_log, report_bug
from Functions import (import_avantage_file, on_save, save_all_sheets_with_plots, save_results_table, open_avg_file,
                       import_multiple_avg_files, create_plot_script_from_excel, on_save_plot, \
    on_save_plot_pdf, on_save_plot_svg, on_exit, undo, redo, toggle_plot, show_shortcuts, show_mini_game, on_about)
from libraries.Utilities import add_draggable_text
from Functions import refresh_sheets, on_sheet_selected_wrapper, toggle_plot, on_save, on_save_plot, on_save_all_sheets, toggle_Col_1, undo, redo
from libraries.LibraryID import PeriodicTableXPS
# from libraries.Asteroid import main as asteroid_main
# from libraries.Solitaire import SolitaireGame
# from libraries.Flappybird import main as flappybird_main
# from libraries.ChemistryLab import ChemistryLabGame
from libraries.Save import on_backup_main
from libraries.Save import save_json_only
from libraries.Utilities import sort_excel_sheets
from libraries.DownloadStats import show_download_stats_window
from libraries.Open import import_multiple_avantage_files
from libraries.Open import import_multiple_kfitting_files

# With conditional imports:
import platform
IS_MAC = platform.system() == 'Darwin'

# Only import games if not on Mac
if not IS_MAC:
    try:
        print("Starting pygame as no macOS has been detected")
        from libraries.Asteroid import main as asteroid_main
        from libraries.Flappybird import main as flappybird_main
        from libraries.TetrisGame import Tetris
        from libraries.Solitaire import SolitaireGame
        from libraries.ChemistryLab import ChemistryLabGame
    except ImportError:
        # Fallback if games can't be imported
        asteroid_main = None
        flappybird_main = None
        Tetris = None
        SolitaireGame = None
        ChemistryLabGame = None
else:
    # Set all game functions to None on Mac
    asteroid_main = None
    flappybird_main = None
    Tetris = None
    SolitaireGame = None
    ChemistryLabGame = None

def show_tetris_game(window):
    """Launch the Tetris game"""
    if IS_MAC:
        wx.MessageBox("Games are not available on Mac due to system compatibility issues.",
                     "Not Available", wx.OK | wx.ICON_INFORMATION)
        return
    try:
        from libraries.TetrisGame import Tetris
        game = Tetris()
        game.run()
    except Exception as e:
        print(f"Error launching Tetris game: {e}")

def show_chemistry_lab_game(window):
    """Launch the Chemistry Lab game"""
    if IS_MAC:
        wx.MessageBox("Games are not available on Mac due to system compatibility issues.",
                     "Not Available", wx.OK | wx.ICON_INFORMATION)
        return
    try:
        from libraries.ChemistryLab import ChemistryLabGame
        game = ChemistryLabGame()
        game.run()
    except Exception as e:
        print(f"Error launching Chemistry Lab game: {e}")

def show_flappybird_game(window):
    """Launch the Flappybird game"""
    if IS_MAC:
        wx.MessageBox("Games are not available on Mac due to system compatibility issues.",
                     "Not Available", wx.OK | wx.ICON_INFORMATION)
        return
    try:
        from libraries.Flappybird import main as flappybird_main
        flappybird_main()
    except Exception as e:
        print(f"Error launching Flappybird game: {e}")

def show_asteroid_game(window):
    """Launch the Asteroid game"""
    if IS_MAC:
        wx.MessageBox("Games are not available on Mac due to system compatibility issues.",
                     "Not Available", wx.OK | wx.ICON_INFORMATION)
        return
    try:
        if asteroid_main:
            asteroid_main()
    except Exception as e:
        print(f"Error launching Asteroid game: {e}")

def launch_solitaire(window):
    if IS_MAC:
        wx.MessageBox("Games are not available on Mac due to system compatibility issues.",
                     "Not Available", wx.OK | wx.ICON_INFORMATION)
        return
    try:
        from libraries.Solitaire import SolitaireGame
        game = SolitaireGame()
        game.run()
    except Exception as e:
        wx.MessageBox(f"Error launching solitaire: {e}", "Error", wx.OK | wx.ICON_ERROR)
def create_widgets(window):
    # Main sizer
    # main_sizer = wx.BoxSizer(wx.HORIZONTAL)
    main_sizer = wx.BoxSizer(wx.VERTICAL)

    # Create horizontal toolbar first so it stays on top
    toolbar_panel = wx.Panel(window.panel)
    toolbar_sizer = wx.BoxSizer(wx.VERTICAL)
    window.toolbar = create_horizontal_toolbar(toolbar_panel, window)  # Pass panel instead of window
    toolbar_sizer.Add(window.toolbar, 0, wx.EXPAND)
    toolbar_panel.SetSizer(toolbar_sizer)
    main_sizer.Add(toolbar_panel, 0, wx.EXPAND)

    # Content sizer for the rest (vertical toolbar and plot area)
    content_sizer = wx.BoxSizer(wx.HORIZONTAL)

    # Create the vertical toolbar as a child of the panel
    window.v_toolbar = create_vertical_toolbar(window.panel, window)

    # # Content sizer (everything except vertical toolbar)  IT WAS USED JUST BEFORE
    # content_sizer = wx.BoxSizer(wx.HORIZONTAL)

    # Create a splitter window
    window.splitter = wx.SplitterWindow(window.panel, style=wx.SP_LIVE_UPDATE)

    # Right frame for the plot
    window.right_frame = wx.Panel(window.splitter)
    # window.right_frame.SetBackgroundColour(wx.Colour(255, 255, 255)) # To change
    right_frame_sizer = wx.BoxSizer(wx.VERTICAL)

    # Create the FigureCanvas
    window.canvas = FigureCanvas(window.right_frame, -1, window.figure)

    # Set up drag and drop for Excel files
    file_drop_target = ExcelDropTarget(window)
    window.canvas.SetDropTarget(file_drop_target)

    plt.tight_layout()
    right_frame_sizer.Add(window.canvas, 1, wx.EXPAND | wx.ALL, 0)

    # Initialize plot_manager
    window.plot_manager = PlotManager(window.ax, window.canvas)
    window.plot_manager.plot_initial_logo()

    # Update plot manager with loaded or default values
    window.update_plot_preferences()

    # Create a hidden NavigationToolbar
    # window.navigation_toolbar = NavigationToolbar(window.canvas)
    # window.navigation_toolbar.Hide()
    window.create_navigation_toolbar()

    window.right_frame.SetSizer(right_frame_sizer)

    # Create grids panel
    grids_panel = create_grids_panel(window)

    for grid in [window.peak_params_grid, window.results_grid]:
        label_font = grid.GetLabelFont()
        label_font.SetPointSize(8)
        grid.SetLabelFont(label_font)

        cell_font = grid.GetDefaultCellFont()
        cell_font.SetPointSize(8)  # Change size as needed
        grid.SetDefaultCellFont(cell_font)


    # Set up the splitter
    window.splitter.SplitVertically(window.right_frame, grids_panel)
    window.splitter.SetMinimumPaneSize(0)
    window.splitter.SetSashGravity(0.5)

    # Set initial sash position
    window.initial_sash_position = 800
    window.splitter.SetSashPosition(window.initial_sash_position)

    # Add splitter to content sizer
    content_sizer.Add(window.v_toolbar, 0, wx.EXPAND)
    content_sizer.Add(window.splitter, 1, wx.EXPAND | wx.ALL, 5)

    # Add content sizer to main sizer
    main_sizer.Add(content_sizer, 1, wx.EXPAND)

    window.panel.SetSizer(main_sizer)

    # # Create the horizontal toolbar .... IT WAS USED JUST BEFORE HORIZ
    # window.toolbar = create_horizontal_toolbar(window)

    update_undo_redo_state(window)
    toggle_Col_1(window)

    # Bind events
    bind_events_widgets(window)


def create_grids_panel(window):
    grids_panel = wx.Panel(window.splitter)

    # Create splitter window
    inner_splitter = wx.SplitterWindow(grids_panel, style=wx.SP_LIVE_UPDATE)

    # Store reference to inner_splitter for later access
    window.inner_splitter = inner_splitter

    # Create peak params panel and grid
    peak_params_panel = wx.Panel(inner_splitter)
    peak_params_sizer = create_peak_params_grid(window, peak_params_panel)
    peak_params_panel.SetSizer(peak_params_sizer)

    # Create results panel and grid
    results_panel = wx.Panel(inner_splitter)
    results_sizer = create_results_grid(window, results_panel)
    results_panel.SetSizer(results_sizer)

    # Add splitter to main sizer
    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(inner_splitter, 1, wx.EXPAND)
    grids_panel.SetSizer(sizer)

    # Split horizontally
    window_height = grids_panel.GetSize().GetHeight()
    split_position = window_height // 2
    inner_splitter.SplitHorizontally(peak_params_panel, results_panel, split_position)
    inner_splitter.SetMinimumPaneSize(100)

    # Bind size event to maintain 50-50 split
    def on_size(event):
        size = inner_splitter.GetSize()
        # Change this value to adjust the ratio (e.g., 0.6 gives 60% to peak params grid)
        inner_splitter.SetSashPosition(int(size.GetHeight() * 0.6))
        event.Skip()

    inner_splitter.Bind(wx.EVT_SIZE, on_size)

    return grids_panel

def create_peak_params_grid(window, parent):
    peak_params_frame_box = wx.StaticBox(parent, label="Peak Fitting Parameters")
    peak_params_sizer = wx.StaticBoxSizer(peak_params_frame_box, wx.VERTICAL)

    window.peak_params_frame = wx.Panel(peak_params_frame_box)
    # window.peak_params_frame.SetBackgroundColour(wx.Colour(255, 255, 255)) # To change back
    peak_params_sizer_inner = wx.BoxSizer(wx.VERTICAL)

    window.peak_params_grid = wx.grid.Grid(window.peak_params_frame)
    window.peak_params_grid.CreateGrid(0, 19)

    # Set column labels and sizes
    column_labels = ["ID", "Peak\nLabel", "Position\n(eV)", "Height\n(CPS)", "FWHM\n(eV)", "\u03c3/\u03b3 (%)\nL/G \n", "Area\n(CPS.eV)",
                     "\u03c3\nW_g", "\u03b3\nW_l", "W_g\nSkew",
                     "Conc.\n(%)", "A/A\u1D00", "Split\n(eV)", "Fitting Model", "Bkg Type", "Bkg Low\n(eV)",
                     "Bkg High\n(eV)", "Bkg Offset Low\n(CPS)", "Bkg Offset High\n(CPS)"]
    for i, label in enumerate(column_labels):
        window.peak_params_grid.SetColLabelValue(i, label)


    # Set grid properties
    default_row_size = 25
    window.peak_params_grid.SetDefaultRowSize(default_row_size)
    window.peak_params_grid.SetColLabelSize(35)
    window.peak_params_grid.SetDefaultColSize(60)
    # window.peak_params_grid.SetLabelBackgroundColour(wx.Colour(230, 250, 230))  # Green background for column labels
    # -------------------------------------------------------To change below to white
    # window.peak_params_grid.SetDefaultCellBackgroundColour(wx.WHITE)  # White background for all cells
    window.peak_params_grid.SetRowLabelSize(25)

    # Ensure all cells have white background
    for row in range(window.peak_params_grid.GetNumberRows()):
        for col in range(window.peak_params_grid.GetNumberCols()):
            window.peak_params_grid.SetCellBackgroundColour(row, col, wx.WHITE)


    # Adjust individual column sizes
    col_sizes = [20, 90, 80, 60, 60, 50, 70, 50, 50, 50, 40, 40, 40, 130, 130, 80, 80, 100, 100]
    for i, size in enumerate(col_sizes):
        window.peak_params_grid.SetColSize(i, size)

    # Store the fitting models as a window attribute
    window.fitting_models = [
        "GL (Area)",
        "SGL (Area)",
        "LA (Area, \u03c3/\u03b3, \u03b3)",
        "Voigt (Area, L/G, \u03c3)",
        "Voigt (Area, \u03c3, \u03b3)",
        "Voigt (Area, L/G, \u03c3, S)",
        "LA (Area, \u03c3, \u03b3)",
        "LA*G (Area, \u03c3/\u03b3, \u03b3)",
        "Pseudo-Voigt (Area)",
        "ExpGauss.(Area, \u03c3, \u03b3)",
        "GL (Height)",
        "SGL (Height)"
    ]

    # For initial setup, create a new editor for each cell
    for row in range(0, window.peak_params_grid.GetNumberRows(), 2):
        # Create a fresh editor instance for each cell
        fresh_editor = wx.grid.GridCellChoiceEditor(window.fitting_models.copy(), allowOthers=False)
        window.peak_params_grid.SetCellEditor(row, 13, fresh_editor)

    # Define the helper function with the same approach
    def set_model_choice_editors(window):
        """Apply choice editors to the fitting model column (13) for all parameter rows."""
        for row in range(0, window.peak_params_grid.GetNumberRows(), 2):
            # Create a new editor instance for each cell
            fresh_editor = wx.grid.GridCellChoiceEditor(window.fitting_models.copy(), allowOthers=False)
            window.peak_params_grid.SetCellEditor(row, 13, fresh_editor)

    # Save the function as an attribute of the window
    window.set_model_choice_editors = set_model_choice_editors

    # Similar approach for the new row handler
    def add_choice_editor_to_new_row(grid, row_num):
        if row_num % 2 == 0:  # Only for parameter rows
            # Create a new editor instance
            fresh_editor = wx.grid.GridCellChoiceEditor(window.fitting_models.copy(), allowOthers=False)
            grid.SetCellEditor(row_num, 13, fresh_editor)

    # Store this function in the window object
    window.add_choice_editor_to_new_row = add_choice_editor_to_new_row



    peak_params_sizer_inner.Add(window.peak_params_grid, 1, wx.EXPAND | wx.ALL, 5)
    window.peak_params_frame.SetSizer(peak_params_sizer_inner)
    peak_params_sizer.Add(window.peak_params_frame, 1, wx.EXPAND | wx.ALL, 5)

    return peak_params_sizer


def create_results_grid(window, parent):
    results_frame_box = wx.StaticBox(parent, label="Results")
    results_sizer = wx.StaticBoxSizer(results_frame_box, wx.VERTICAL)

    window.results_frame = wx.Panel(results_frame_box)
    results_sizer_inner = wx.BoxSizer(wx.VERTICAL)

    window.results_grid = wx.grid.Grid(window.results_frame)
    window.results_grid.CreateGrid(0, 31)

    # Set column labels and properties for results grid
    column_labels = ["Peak\nLabel", "Position\n(eV)", "Height\n(CPS)", "FWHM\n(eV)", "L/G \n\u03c3/\u03b3 (%)",
                     "Area\n(CPS.eV)", "Atomic\n(%)", " ", "RSF", "TXFN", "ECF", "Instr.","Fitting Model",
                     "Corr. Area\n(a.u.)",
                     "\u03c3 or \u03B1\nW_g", "\u03b3 or \u03B2\nW_l", "Bkg Type", "Bkg Low\n(eV)", "Bkg High\n(eV)", "Bkg Offset Low\n(CPS)",
                     "Bkg Offset High\n(CPS)", "Sheetname", "Position\nConstraint", "Height\nConstraint",
                     "FWHM\nConstraint", "L/G\nConstraint", "Area\nConstraint", "\u03c3\nConstraint",
                     "\u03b3\nConstraint", "Weight\n(%)","Mass\n(amu)"]
    for i, label in enumerate(column_labels):
        window.results_grid.SetColLabelValue(i, label)

    window.results_grid.SetDefaultRowSize(25)
    window.results_grid.SetDefaultColSize(60)
    window.results_grid.SetRowLabelSize(25)
    window.results_grid.SetColLabelSize(35)
    # window.results_grid.SetLabelBackgroundColour(wx.Colour(255, 255, 255))
    # ------------------------------------------> To change below for white
    # window.results_grid.SetDefaultCellBackgroundColour(wx.WHITE)

    # Adjust specific column sizes
    col_sizes = [100, 60, 60, 50, 50, 80, 50, 20, 30, 30, 50, 80,120, 60, 80, 70, 70, 100, 100, 80, 80, 80, 120, 120,
                 120,70,70,70,50,45,40]
    for i, size in enumerate(col_sizes):
        window.results_grid.SetColSize(i, size)

    # Set renderer for checkbox column
    checkbox_renderer = CheckboxRenderer()
    for row in range(window.results_grid.GetNumberRows()):
        window.results_grid.SetCellRenderer(row, 7, checkbox_renderer)

    results_sizer_inner.Add(window.results_grid, 1, wx.EXPAND | wx.ALL, 5)
    window.results_frame.SetSizer(results_sizer_inner)
    results_sizer.Add(window.results_frame, 1, wx.EXPAND | wx.ALL, 5)

    return results_sizer


def bind_events_widgets(window):
    window.results_grid.Bind(wx.EVT_KEY_DOWN, window.on_key_down)
    window.peak_params_grid.Bind(wx.grid.EVT_GRID_SELECT_CELL, window.on_grid_select)
    window.splitter.Bind(wx.EVT_SPLITTER_SASH_POS_CHANGED, window.on_splitter_changed)
    window.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, window.on_checkbox_update)
    window.canvas.mpl_connect('button_release_event', window.on_plot_mouse_release)
    window.peak_params_grid.Bind(wx.EVT_LEFT_UP, window.on_peak_params_mouse_release)



def create_menu(window):
    menubar = wx.MenuBar()

    # Create menus
    file_menu = wx.Menu()
    import_menu = wx.Menu()
    export_menu = wx.Menu()
    edit_menu = wx.Menu()
    view_menu = wx.Menu()
    tools_menu = wx.Menu()
    help_menu = wx.Menu()
    save_menu = wx.Menu()

    # File menu items

    # Add "New Instance" option to File menu
    new_instance_item = file_menu.Append(wx.NewId(), "New Window\tCtrl+N")
    window.Bind(wx.EVT_MENU, lambda event: launch_new_instance(), new_instance_item)

    open_item = file_menu.Append(wx.ID_OPEN, "Open \tCtrl+O")
    window.Bind(wx.EVT_MENU, lambda event: open_xlsx_file(window), open_item)

    # Open KherveFitting submenu
    open_kfitting_menu = wx.Menu()

    open_single_kfitting_item = open_kfitting_menu.Append(wx.NewId(), "Open KFitting file (.xlsx)")
    window.Bind(wx.EVT_MENU, lambda event: open_xlsx_file(window), open_single_kfitting_item)

    open_multiple_kfitting_item = open_kfitting_menu.Append(wx.NewId(), "Open Multiple KFitting files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_kfitting_files(window), open_multiple_kfitting_item)

    file_menu.AppendSubMenu(open_kfitting_menu, "Open")

    # Recent files submenu
    window.recent_files_menu = wx.Menu()
    file_menu.AppendSubMenu(window.recent_files_menu, "Recent Files")

    # Save submenu items
    save_json_item = save_menu.Append(wx.NewId(), "Save Data (.json)")
    window.Bind(wx.EVT_MENU, lambda event: save_json_only(window), save_json_item)

    save_Excel_item = save_menu.Append(wx.NewId(), "Export/Save this Core Level to Excel")
    window.Bind(wx.EVT_MENU, lambda event: on_save(window), save_Excel_item)

    save_all_item = save_menu.Append(wx.NewId(), "Export/Save all Core Levels to Excel")
    window.Bind(wx.EVT_MENU, lambda event: save_all_sheets_with_plots(window), save_all_item)

    save_plot_only_item = save_menu.Append(wx.NewId(), "Save Plot Only in Excel")
    window.Bind(wx.EVT_MENU, lambda event: save_plot_only_to_excel(window), save_plot_only_item)

    file_menu.AppendSubMenu(save_menu, "Save")

    save_as_item = file_menu.Append(wx.ID_SAVEAS, "Save As...\tCtrl+Shift+S")
    window.Bind(wx.EVT_MENU, lambda event: on_save_as(window), save_as_item)

    # Import submenu items
    import_vamas_item = import_menu.Append(wx.NewId(), "Import Vamas Data file (.vms)")
    window.Bind(wx.EVT_MENU, lambda event: open_vamas_file_dialog(window), import_vamas_item)

    import_avantage_item = import_menu.Append(wx.NewId(), "Import Avantage Data file (.xlsx or .xls)")
    window.Bind(wx.EVT_MENU, lambda event: import_avantage_file(window), import_avantage_item)

    import_multiple_avantage_item = import_menu.Append(wx.NewId(), "Import Multiple Avantage xlsx files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_avantage_files(window), import_multiple_avantage_item)

    import_kal_item = import_menu.Append(wx.NewId(), "Import Kratos Data file (.kal)")
    window.Bind(wx.EVT_MENU, lambda event: open_kal_file_dialog(window), import_kal_item)

    import_spe_item = import_menu.Append(wx.NewId(), "Import Phi Data file (.spe)")
    window.Bind(wx.EVT_MENU, lambda event: open_spe_file_dialog(window), import_spe_item)

    # Add MRS file import items
    import_mrs_item = import_menu.Append(wx.NewId(), "Import MRS Data file (.mrs)")
    window.Bind(wx.EVT_MENU, lambda event: import_mrs_file(window), import_mrs_item)

    import_multiple_mrs_item = import_menu.Append(wx.NewId(), "Import Multiple MRS files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_mrs_files(window), import_multiple_mrs_item)

    import_avg_item = import_menu.Append(wx.NewId(), "Import AVG file (.avg)")
    window.Bind(wx.EVT_MENU, lambda event: open_avg_file(window), import_avg_item)

    import_multiple_avg_item = import_menu.Append(wx.NewId(), "Import Multiple AVG files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_avg_files(window), import_multiple_avg_item)

    # Import ASC files
    import_xps_asc_item = import_menu.Append(wx.NewId(), "Import XPS .asc file")
    window.Bind(wx.EVT_MENU, lambda event: import_xps_asc_file(window), import_xps_asc_item)

    import_multiple_xps_asc_item = import_menu.Append(wx.NewId(), "Import Multiple XPS .asc files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_xps_asc_files(window), import_multiple_xps_asc_item)

    # Import VG Microtech
    import_vg_microtech_item = import_menu.Append(wx.NewId(), "Import VG-Microtech file (.1)")
    window.Bind(wx.EVT_MENU, lambda event: open_vg_microtech_file_dialog(window), import_vg_microtech_item)

    import_multiple_vg_microtech_item = import_menu.Append(wx.NewId(), "Import Multiple VG-Microtech files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_vg_microtech_files(window),
                import_multiple_vg_microtech_item)

    # Import CSV files
    import_xps_csv_item = import_menu.Append(wx.NewId(), "Import XPS .csv file")
    window.Bind(wx.EVT_MENU, lambda event: import_xps_csv_file(window), import_xps_csv_item)

    import_multiple_xps_csv_item = import_menu.Append(wx.NewId(), "Import Multiple XPS .csv files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_xps_csv_files(window), import_multiple_xps_csv_item)

    import_raman_item = import_menu.Append(wx.NewId(), "Import Raman .txt file")
    window.Bind(wx.EVT_MENU, lambda event: import_raman_txt_file(window), import_raman_item)

    import_multiple_raman_item = import_menu.Append(wx.NewId(), "Import Multiple Raman .txt files (folder)")
    window.Bind(wx.EVT_MENU, lambda event: import_multiple_raman_files(window), import_multiple_raman_item)

    # Export submenu items
    export_python_plot_item = export_menu.Append(wx.NewId(), "Python Plot")
    window.Bind(wx.EVT_MENU, lambda event: create_plot_script_from_excel(window), export_python_plot_item)

    save_plot_item = export_menu.Append(wx.NewId(), "Export plot as PNG")
    window.Bind(wx.EVT_MENU, lambda event: on_save_plot(window), save_plot_item)

    save_plot_item_pdf = export_menu.Append(wx.NewId(), "Export plot as PDF")
    window.Bind(wx.EVT_MENU, lambda event: on_save_plot_pdf(window), save_plot_item_pdf)

    save_plot_item_svg = export_menu.Append(wx.NewId(), "Export plot as SVG")
    window.Bind(wx.EVT_MENU, lambda event: on_save_plot_svg(window), save_plot_item_svg)

    export_txt_item = export_menu.Append(wx.ID_ANY, "Export data as TXT",
                                         "Export current core level to TXT file")
    export_csv_item = export_menu.Append(wx.ID_ANY, "Export data as CSV",
                                         "Export current core level to CSV file")
    export_dat_item = export_menu.Append(wx.ID_ANY, "Export data as DAT",
                                         "Export current core level to DAT file")

    window.Bind(wx.EVT_MENU, lambda event: export_sheet_to_txt(window), export_txt_item)
    window.Bind(wx.EVT_MENU, lambda event: export_sheet_to_csv(window), export_csv_item)
    window.Bind(wx.EVT_MENU, lambda event: export_sheet_to_dat(window), export_dat_item)

    export_menu.AppendSeparator()
    word_report_item = export_menu.Append(wx.NewId(), "Create Report (.docx)")
    window.Bind(wx.EVT_MENU, lambda event: export_word_report(window), word_report_item)

    open_location_item = file_menu.Append(wx.NewId(), "Open File Location")
    window.Bind(wx.EVT_MENU, lambda event: open_file_location(window), open_location_item)

    # Append submenus to file menu
    file_menu.AppendSubMenu(import_menu, "Import")
    file_menu.AppendSubMenu(export_menu, "Export")

    # Exit item
    file_menu.AppendSeparator()
    exit_item = file_menu.Append(wx.ID_EXIT, "Exit\tCtrl+Q")
    window.Bind(wx.EVT_MENU, lambda event: on_exit(window, event), exit_item)

    # Edit menu items
    undo_item = edit_menu.Append(wx.ID_UNDO, "Undo\tCtrl+Z")
    redo_item = edit_menu.Append(wx.ID_REDO, "Redo\tCtrl+Y")
    window.Bind(wx.EVT_MENU, lambda event: undo(window), undo_item)
    window.Bind(wx.EVT_MENU, lambda event: redo(window), redo_item)
    edit_menu.AppendSeparator()

    preferences_item = edit_menu.Append(wx.ID_PREFERENCES, "Preferences")
    window.Bind(wx.EVT_MENU, window.on_preferences, preferences_item)

    # View menu items
    sample_manager_item = view_menu.Append(wx.NewId(), "Sample Manager")
    window.Bind(wx.EVT_MENU, window.on_open_file_manager, sample_manager_item)

    ToggleFitting_item = view_menu.Append(wx.NewId(), "Toggle Peak Fitting")
    window.Bind(wx.EVT_MENU, lambda event: toggle_plot(window), ToggleFitting_item)

    ToggleLegend_item = view_menu.Append(wx.NewId(), "Toggle Legend")
    window.Bind(wx.EVT_MENU, lambda event: window.plot_manager.toggle_legend(), ToggleLegend_item)

    ToggleFit_item = view_menu.Append(wx.NewId(), "Toggle Fit Results")
    window.Bind(wx.EVT_MENU, lambda event: window.plot_manager.toggle_fitting_results(), ToggleFit_item)

    ToggleRes_item = view_menu.Append(wx.NewId(), "Toggle Residuals")
    window.Bind(wx.EVT_MENU, lambda event: window.plot_manager.toggle_residuals(), ToggleRes_item)

    toggle_energy_item = view_menu.AppendCheckItem(wx.NewId(), "Show Kinetic Energy\tCtrl+B")
    window.Bind(wx.EVT_MENU, lambda event: window.toggle_energy_scale(), toggle_energy_item)
    window.toggle_energy_item = toggle_energy_item

    # Tools menu items
    Area_item = tools_menu.Append(wx.NewId(), "Calculate Area Under Curve\tCtrl+A")
    window.Bind(wx.EVT_MENU, lambda event: window.on_open_background_window(), Area_item)

    Fitting_item = tools_menu.Append(wx.NewId(), "Create Peak Model\tCtrl+P")
    window.Bind(wx.EVT_MENU, lambda event: window.on_open_fitting_window(), Fitting_item)

    Dparam_item = tools_menu.Append(wx.NewId(), "D-parameter\tCtrl+D")
    window.Bind(wx.EVT_MENU, window.on_differentiate, Dparam_item)

    ID_item = tools_menu.Append(wx.NewId(), "Element ID\tCtrl+I")
    window.Bind(wx.EVT_MENU, window.open_periodic_table, ID_item)

    Noise_item = tools_menu.Append(wx.NewId(), "Noise Analysis")
    window.Bind(wx.EVT_MENU, lambda event: window.on_open_noise_analysis_window, Noise_item)

    # Help menu items
    # mini_help_item = help_menu.Append(wx.NewId(), "Help")
    # window.Bind(wx.EVT_MENU, window.on_mini_help, mini_help_item)

    shortcuts_item = help_menu.Append(wx.NewId(), "List of Shortcuts\tCtrl+K")
    window.Bind(wx.EVT_MENU, lambda event: show_shortcuts(window), shortcuts_item)

    manual_item = help_menu.Append(wx.NewId(), "Open Full Manual\tCtrl+M")
    window.Bind(wx.EVT_MENU, lambda event: open_manual(window), manual_item)

    yt_videos_item = help_menu.Append(wx.NewId(), "KherveFitting Videos")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open("https://www.youtube.com/@xpsexamples-imperialcolleg6571"),
                yt_videos_item)


    # Create a submenu for Knowledge links
    knowledge_menu = wx.Menu()

    paper_menu = wx.Menu()


    multiplet_item = paper_menu.Append(wx.NewId(), "Multiplet splitting")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://analyticalsciencejournals.onlinelibrary.wiley.com/doi/epdf/10.1002/sia.7383"),
                multiplet_item)

    d_parameter_item = paper_menu.Append(wx.NewId(), "Carbon sp2")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://www.mdpi.com/2311-5629/7/3/51"),
                d_parameter_item)

    Carbonpeaktable_item = paper_menu.Append(wx.NewId(), "Fitting Carbon")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://drive.google.com/file/d/1fyXNfX46cN7q2sYRqwBM-jaj7C2cNSPA/view"),
                Carbonpeaktable_item)

    Fepeaktable_item = paper_menu.Append(wx.NewId(), "Fitting Transition Metal Cr/Mn/Fe/Co/Ni")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://drive.google.com/file/d/1Kxx_j2kCpj8Hrd3XwbmEcJ16qHpgDmuN/view"),
                Fepeaktable_item)

    Fepeaktable_item = paper_menu.Append(wx.NewId(), "Fitting Transition Metal Cu/Ti/V/Sc/Zn")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://drive.google.com/file/d/1YYw7O1JVW4Ni_3GJv72uTE9KVE4Cg1S9/view"),
                Fepeaktable_item)




    knowledge_menu.AppendSubMenu(paper_menu, "Useful papers")



    # Add items to the Knowledge submenu
    thermo_item = knowledge_menu.Append(wx.NewId(), "Thermo Knowledge")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://www.thermofisher.com/uk/en/home/materials-science/learning-center/periodic-table.html"),
                thermo_item)

    # Add Guide to XPS link
    guide_xps_item = knowledge_menu.Append(wx.NewId(), "Guide to XPS")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://pubs.aip.org/jva/collection/1440/Special-Topic-Collection-Reproducibility"),
                guide_xps_item)

    biesinger_item = knowledge_menu.Append(wx.NewId(), "XPSfitting by M. Biesinger")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://www.xpsfitting.com/"),
                biesinger_item)

    yt_videos_item2 = knowledge_menu.Append(wx.NewId(), "M. Biesinger Videos")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open("https://www.youtube.com/@markbiesinger/videos"),
                yt_videos_item2)

    nist_item = knowledge_menu.Append(wx.NewId(), "NIST XPS")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://srdata.nist.gov/xps"), nist_item)

    dream_nist_item = knowledge_menu.Append(wx.NewId(), "KherveNIST")
    window.Bind(wx.EVT_MENU, lambda event: window.open_dream_nist(), dream_nist_item)

    harwell_item = knowledge_menu.Append(wx.NewId(), "HarwellXPS Guru")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open(
        "https://www.harwellxps.guru/"),
                harwell_item)

    yt_videos_item3 = knowledge_menu.Append(wx.NewId(), "HarwellXPS Guru Videos")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open("https://www.youtube.com/@HarwellXPS"),
                yt_videos_item3)

    yt_videos_item4 = knowledge_menu.Append(wx.NewId(), "Casa XPS Videos")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open("https://www.youtube.com/@casaxpscasasoftware4605/videos"),
                yt_videos_item4)


    # Add the Knowledge submenu to the Help menu
    help_menu.AppendSubMenu(knowledge_menu, "Knowledge")

    # Create Bored submenu
    bored_menu = wx.Menu()

    # Create Bored submenu only if not on Mac
    if not IS_MAC:
        bored_menu = wx.Menu()

        atoms_item = bored_menu.Append(wx.NewId(), "Atoms")
        window.Bind(wx.EVT_MENU, lambda event: show_mini_game(window), atoms_item)

        chemistry_lab_item = bored_menu.Append(wx.NewId(), "Material Lab")
        window.Bind(wx.EVT_MENU, lambda event: show_chemistry_lab_game(window), chemistry_lab_item)

        asteroid_item = bored_menu.Append(wx.NewId(), "Meteors Smash")
        window.Bind(wx.EVT_MENU, lambda event: show_asteroid_game(window), asteroid_item)

        solitaire_item = bored_menu.Append(wx.NewId(), "Kherve Solitaire")
        window.Bind(wx.EVT_MENU, lambda event: launch_solitaire(window), solitaire_item)

        tetris_item = bored_menu.Append(wx.NewId(), "Kherve Tetris")
        window.Bind(wx.EVT_MENU, lambda event: show_tetris_game(window), tetris_item)

        flappybird_item = bored_menu.Append(wx.NewId(), "Khervey the Flappy Bird")
        window.Bind(wx.EVT_MENU, lambda event: show_flappybird_game(window), flappybird_item)

        help_menu.AppendSubMenu(bored_menu, "Take a Break")

    resubmit_form_item = help_menu.Append(wx.NewId(), "Registration Form")
    window.Bind(wx.EVT_MENU, lambda event: launch_registration_form(window), resubmit_form_item)

    report_bug_item = help_menu.Append(wx.NewId(), "Report Bug")
    window.Bind(wx.EVT_MENU, lambda event: report_bug(window), report_bug_item)

    # version_log_item = help_menu.Append(wx.NewId(), "Version Log")
    # window.Bind(wx.EVT_MENU, lambda event: show_version_log(window), version_log_item)

    download_stats_item = help_menu.Append(wx.NewId(), "Download Stats")
    window.Bind(wx.EVT_MENU, lambda event: show_download_stats_window(window), download_stats_item)

    coffee_item = help_menu.Append(wx.NewId(), "Buy Me a Coffee")
    window.Bind(wx.EVT_MENU, lambda event: webbrowser.open("https://buymeacoffee.com/gkerherve"), coffee_item)

    # libraries_item = help_menu.Append(wx.NewId(), "Libraries Used")
    # window.Bind(wx.EVT_MENU, lambda event: show_libraries_used(window), libraries_item)

    about_item = help_menu.Append(wx.ID_ABOUT, "About")
    window.Bind(wx.EVT_MENU, lambda event: on_about(window, event), about_item)



    # Add menus to menubar
    menubar.Append(file_menu, "&File")
    menubar.Append(edit_menu, "&Edit")
    menubar.Append(view_menu, "&View")
    menubar.Append(tools_menu, "&Tools")
    menubar.Append(help_menu, "&Help")

    window.SetMenuBar(menubar)


def create_rightside_toolbar(parent, window):
    r_toolbar = wx.ToolBar(parent, style= wx.TB_RIGHT)
    r_toolbar.SetToolBitmapSize(wx.Size(25, 25))

    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, "Icons")

    r_save_peaks_tool = r_toolbar.AddTool(wx.ID_ANY, 'Save Peaks Library',
                                          wx.Bitmap(os.path.join(icon_path, "LibSave.png"), wx.BITMAP_TYPE_PNG),
                                          shortHelp="Save peaks parameters to library")

    r_open_peaks_tool = r_toolbar.AddTool(wx.ID_ANY, 'Open Peaks Library',
                                          wx.Bitmap(os.path.join(icon_path, "LibOpen.png"), wx.BITMAP_TYPE_PNG),
                                          shortHelp="Load peaks parameters from library")

    r_toolbar.Realize()
    window.Bind(wx.EVT_TOOL, lambda event: save_peaks_library(window), r_save_peaks_tool)
    window.Bind(wx.EVT_TOOL, lambda event: load_peaks_library(window), r_open_peaks_tool)

    return r_toolbar

def launch_new_instance():
    """Launch a new instance of the application"""
    if getattr(sys, 'frozen', False):
        # Running as executable
        executable = sys.executable
        subprocess.Popen([executable])
    else:
        # Running as script
        script_path = sys.argv[0]
        subprocess.Popen([sys.executable, script_path])


def create_horizontal_toolbar(parent, window):
    # # To use with normal toolbar
    # toolbar = window.CreateToolBar(style=  wx.TB_FLAT)
    # toolbar.SetToolBitmapSize(wx.Size(25, 25))

    # Create toolbar as a panel instead of using window.CreateToolBar()
    # toolbar_panel = wx.Panel(window.panel)
    toolbar = wx.ToolBar(parent, style=wx.TB_FLAT)
    toolbar.SetToolBitmapSize(wx.Size(25, 25))

    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, "Icons")

    separators = []

    # File operations
    open_file_tool = toolbar.AddTool(wx.ID_ANY, 'Open File', wx.Bitmap(os.path.join(icon_path, "open-folder-3.png"),
                                                                       wx.BITMAP_TYPE_PNG), shortHelp="Open File\tCtrl+O")

    # Change the default save tool
    quick_save_tool = toolbar.AddTool(wx.ID_ANY, 'Save Data',
                                      wx.Bitmap(os.path.join(icon_path, "Save_Json-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Save Data - Default (JSON only)\tCtrl+S")
    window.Bind(wx.EVT_TOOL, lambda event: save_json_only(window), quick_save_tool)

    save_tool = toolbar.AddTool(wx.ID_ANY, 'Export to Excel', wx.Bitmap(os.path.join(icon_path, "Save-excel-3.png"),
                                wx.BITMAP_TYPE_PNG), shortHelp="Export this core level to Excel")

    save_all_tool = toolbar.AddTool(wx.ID_ANY, 'Save All Sheets',
                                wx.Bitmap(os.path.join(icon_path, "save-Multi-3.png"), wx.BITMAP_TYPE_PNG),
                                    shortHelp="Export all core level to Excel (Not recommended for 50 above core "
                                              "levels)")

    # save_plot_tool = toolbar.AddTool(wx.ID_ANY, 'Export Plot as PNG', wx.Bitmap(os.path.join(icon_path, "save-PNG-25.png"), wx.BITMAP_TYPE_PNG), shortHelp="Export Plot as PNG")

    # toolbar.AddSeparator()
    window.undo_tool = toolbar.AddTool(wx.ID_ANY, 'Undo', wx.Bitmap(os.path.join(icon_path, "undo-3.png"),
                                                                    wx.BITMAP_TYPE_PNG), shortHelp="Undo -- For peaks properties only")
    window.redo_tool = toolbar.AddTool(wx.ID_ANY, 'Redo', wx.Bitmap(os.path.join(icon_path, "redo-3.png"),
                                                                    wx.BITMAP_TYPE_PNG), shortHelp="Redo -- For peaks properties only")
    # toolbar.AddSeparator()

    # Add sort sheets button
    sort_icon = os.path.join(icon_path, "Sort-3.png")
    if os.path.exists(sort_icon):
        sort_bmp = wx.Bitmap(sort_icon)
    else:
        sort_bmp = wx.ArtProvider.GetBitmap(wx.ART_SORT_ASC, wx.ART_TOOLBAR)
    sort_tool = toolbar.AddTool(wx.ID_ANY, "Sort Sheets", sort_bmp, "Sort sheets by sample groups")
    window.Bind(wx.EVT_TOOL, lambda event: sort_excel_sheets(window), sort_tool)

    # Add File Manager button to toolbar
    file_manager_bmp = wx.Bitmap(os.path.join(icon_path, "list-view-3.png"), wx.BITMAP_TYPE_PNG)
    file_manager_tool = toolbar.AddTool(wx.ID_ANY, "Sample/Experiment Manager", file_manager_bmp,
                                        "Open Sample/Experiment Manager. Make sure to use F2 or Ctrl+2 to plot with peak models")
    window.Bind(wx.EVT_TOOL, window.on_open_file_manager, file_manager_tool)


    # Sheet selection
    window.sheet_combobox = wx.ComboBox(toolbar, style=wx.CB_READONLY)
    window.sheet_combobox.SetToolTip("Select Sheet")
    toolbar.AddControl(window.sheet_combobox)
    window.sheet_combobox.Bind(wx.EVT_COMBOBOX, lambda event: on_sheet_selected(window, event))

    refresh_folder_tool = toolbar.AddTool(wx.ID_ANY, 'Refresh Excel File', wx.Bitmap(os.path.join(icon_path,
                                                                                                  "Refresh-3.png"),
                                                                                     wx.BITMAP_TYPE_PNG),
                                          shortHelp="Refresh Excel File and json file. Used when the Excel File has "
                                                    "more sheets or when the file does not work well")

    delete_sheet_tool = toolbar.AddTool(wx.ID_ANY, 'Delete Core Level/Survey',
                                        wx.Bitmap(os.path.join(icon_path, "delete-3.png"), wx.BITMAP_TYPE_PNG),
                                        shortHelp="Delete current Core Level/Survey")
    window.Bind(wx.EVT_TOOL, lambda event: on_delete_sheet(window, event), delete_sheet_tool)

    copy_sheet_tool = toolbar.AddTool(wx.ID_ANY, 'Copy/Paste Core Level',
                                      wx.Bitmap(os.path.join(icon_path, "copy-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Copy/Paste this Core Level/Survey at the end of this file")

    join_sheets_tool = toolbar.AddTool(wx.ID_ANY, 'Join Core Level/Survey',
                                       wx.Bitmap(os.path.join(icon_path, "join-3.png"), wx.BITMAP_TYPE_PNG),
                                       shortHelp="Join Multiple Core Level/Survey")


    rename_sheet_tool = toolbar.AddTool(wx.ID_ANY, 'Rename Core Level/Survey',
                                        wx.Bitmap(os.path.join(icon_path, "rename-3.png"), wx.BITMAP_TYPE_PNG),
                                        shortHelp="Rename current Core Level/Survey")
    window.Bind(wx.EVT_TOOL, lambda evt: _show_rename_dialog(window), rename_sheet_tool)

    def _show_rename_dialog(window):
        # Get the current sheet name
        current_sheet_name = window.sheet_combobox.GetValue()

        if hasattr(window, 'file_manager') and window.file_manager is not None:
            try:
                # Close existing file manager
                window.file_manager.Close()
                window.file_manager.Destroy()
                window.file_manager = None
            except Exception as e:
                print(f"Error refreshing file manager: {e}")
                pass

        # Show dialog with current sheet name pre-populated
        dlg = wx.TextEntryDialog(window, 'Enter new sheet name:', 'Rename Sheet', current_sheet_name)
        if dlg.ShowModal() == wx.ID_OK:
            new_name = dlg.GetValue()
            if new_name and new_name != current_sheet_name:
                from libraries.Utilities import rename_sheet
                rename_sheet(window, new_name)
        dlg.Destroy()

    crop_tool = toolbar.AddTool(wx.ID_ANY, 'Crop',
                                wx.Bitmap(os.path.join(icon_path, "Crop-3.png"), wx.BITMAP_TYPE_PNG),
                                shortHelp="Crop data to new sheet")
    window.Bind(wx.EVT_TOOL, lambda event: CropWindow(window).Show(), crop_tool)

    toolbar.AddSeparator()

    # BE correction
    window.be_correction_spinbox = wx.SpinCtrlDouble(toolbar, value='0.00', min=-10000.00, max=10000.00, inc=0.01,
                                                     size=(70, -1))
    window.be_correction_spinbox.SetDigits(2)
    window.be_correction_spinbox.SetToolTip("BE Correction")
    toolbar.AddControl(window.be_correction_spinbox)

    auto_be_button = toolbar.AddTool(wx.ID_ANY, 'Auto BE', wx.Bitmap(os.path.join(icon_path, "BEcorrect-3.png"),
                                                                     wx.BITMAP_TYPE_PNG), shortHelp="Automatic binding energy correction")


    toolbar.AddSeparator()

    # Analysis tools
    bkg_tool = toolbar.AddTool(wx.ID_ANY, 'Background', wx.Bitmap(os.path.join(icon_path, "BKG-3.png"),
                                                                  wx.BITMAP_TYPE_PNG), shortHelp="Calculate Area "
                                                                                                 "Under Curve\tCtrl+A")
    fitting_tool = toolbar.AddTool(wx.ID_ANY, 'Fitting', wx.Bitmap(os.path.join(icon_path, "C1s-3.png"),
                                                                   wx.BITMAP_TYPE_PNG), shortHelp="Create Peaks "
                                                                                                  "Model \tCtrl+P")


    diff_tool = toolbar.AddTool(wx.ID_ANY, 'Differentiate',
                                wx.Bitmap(os.path.join(icon_path, "Dpara-3.png"), wx.BITMAP_TYPE_PNG),
                                shortHelp="D-parameter Calculation")







    plot_mod_tool = toolbar.AddTool(wx.ID_ANY, 'Plot Modifications',
                                    wx.Bitmap(os.path.join(icon_path, "Mod-3.png"), wx.BITMAP_TYPE_PNG),
                                    shortHelp="Plot modifications window")
    window.Bind(wx.EVT_TOOL, lambda evt: PlotModWindow(window).Show(), plot_mod_tool)
    window.Bind(wx.EVT_TOOL, window.on_differentiate, diff_tool)

    toolbar.AddSeparator()

    # noise_analysis_tool = toolbar.AddTool(wx.ID_ANY, 'Noise Analysis',
    #                                       wx.Bitmap(os.path.join(icon_path, "Noise-25.png"), wx.BITMAP_TYPE_PNG),
    #                                       shortHelp="Open Noise Analysis Window")


    id_tool = toolbar.AddTool(wx.ID_ANY, 'ID', wx.Bitmap(os.path.join(icon_path, "ID-3.png"), wx.BITMAP_TYPE_PNG),
                              shortHelp="Element identifications (ID)")

    nist_tool = toolbar.AddTool(wx.ID_ANY, 'NIST Database',
                                wx.Bitmap(os.path.join(icon_path, "NIST-3.png"), wx.BITMAP_TYPE_PNG),
                                shortHelp="Open KherveDB database")
    # window.Bind(wx.EVT_TOOL, lambda event: window.open_dream_nist(), nist_tool)
    # window.Bind(wx.EVT_TOOL, lambda event: open_dream_nist_handler(window), nist_tool)
    window.Bind(wx.EVT_TOOL, lambda event: window.open_kherve_db(), nist_tool)

    toolbar.AddStretchableSpace()




    # Export and toggle tools
    save_peaks_tool = toolbar.AddTool(wx.ID_ANY, 'Save Peaks Library',
                                      wx.Bitmap(os.path.join(icon_path, "LibSave-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Save peaks parameters to library")

    open_peaks_tool = toolbar.AddTool(wx.ID_ANY, 'Open Peaks Library',
                                      wx.Bitmap(os.path.join(icon_path, "LibOpen-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Load peaks parameters from library")

    export_tool = toolbar.AddTool(wx.ID_ANY, 'Export Results', wx.Bitmap(os.path.join(icon_path, "Export-3.png"),
                                                                         wx.BITMAP_TYPE_PNG), shortHelp="Export to Results Grid")

    # Create delete toolbar instance
    window.delete_toolbar = DeleteToolbar(window)

    # Add delete master toggle tool
    delete_master_tool = toolbar.AddTool(wx.ID_ANY, 'Delete',
                                         wx.Bitmap(os.path.join(icon_path, "DeleteRow-3.png"), wx.BITMAP_TYPE_PNG),
                                         shortHelp="Delete Options for Results Grid")

    # Add backup tool
    backup_icon = os.path.join(icon_path, "backup-3.png")
    if os.path.exists(backup_icon):
        backup_bmp = wx.Bitmap(backup_icon)
    else:
        backup_bmp = wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE_AS, wx.ART_TOOLBAR)
    backup_tool = toolbar.AddTool(wx.ID_ANY, "Backup", backup_bmp, "Create a backup of current files")
    window.Bind(wx.EVT_TOOL, lambda event: on_backup_main(window), backup_tool)

    # Add quick settings if enabled
    if hasattr(window, 'enable_quick_settings') and window.enable_quick_settings:
        window.quick_settings.create_quick_settings_tool(toolbar)

    Setting_tool = toolbar.AddTool(wx.ID_ANY, 'Load Settings',
                                      wx.Bitmap(os.path.join(icon_path, "Settings-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Open Preference Window")


    toggle_Col_1_tool = toolbar.AddTool(wx.ID_ANY, 'Toggle Residuals',
                                        wx.Bitmap(os.path.join(icon_path, "HideColumn-3.png"), wx.BITMAP_TYPE_PNG),
                                        shortHelp="Toggle Columns Peak Fitting Parameters")
    window.toggle_right_panel_tool = window.add_toggle_tool(toolbar, "Toggle Right Panel",
                                                            wx.ArtProvider.GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR))


    def show_delete_toolbar(event):
        if not window.delete_toolbar.IsShown():
            pos = toolbar.GetScreenPosition()
            window.delete_toolbar.SetPosition((pos.x + toolbar.GetSize().width - 100, pos.y + 30))
            window.delete_toolbar.Show()
        else:
            window.delete_toolbar.Hide()


    window.Bind(wx.EVT_TOOL, show_delete_toolbar, delete_master_tool)
    window.Bind(wx.EVT_TOOL, lambda evt: on_delete_all(window, evt), window.delete_toolbar.delete_all_tool)
    window.Bind(wx.EVT_TOOL, lambda evt: on_delete_last(window, evt), window.delete_toolbar.delete_last_tool)
    window.Bind(wx.EVT_TOOL, lambda evt: on_delete_first(window, evt), window.delete_toolbar.delete_first_tool)

    toolbar.Realize()

    # Bind events
    bind_toolbar_events(window, open_file_tool, refresh_folder_tool, bkg_tool, fitting_tool,
                        # noise_analysis_tool,
                        # toggle_legend_tool, toggle_fit_results_tool, toggle_residuals_tool, plot_tool, toggle_peak_fill_tool,
                        save_tool, #save_plot_tool,
                        save_all_tool, toggle_Col_1_tool, export_tool, auto_be_button, id_tool)
    # toolbar.Bind(wx.EVT_TOOL, lambda event: window.plot_manager.toggle_y_axis(), toggle_y_axis_tool)
    window.Bind(wx.EVT_MENU, window.on_preferences, Setting_tool)
    window.Bind(wx.EVT_TOOL, lambda event: save_peaks_library(window), save_peaks_tool)
    window.Bind(wx.EVT_TOOL, lambda event: load_peaks_library(window), open_peaks_tool)
    window.Bind(wx.EVT_TOOL, lambda event: copy_sheet(window), copy_sheet_tool)
    window.Bind(wx.EVT_TOOL, lambda event: JoinSheetsWindow(window).Show(), join_sheets_tool)
    window.be_correction_spinbox.Bind(wx.EVT_SPINCTRLDOUBLE, window.on_be_correction_change)
    # window.Bind(wx.EVT_TOOL, lambda event: save_peaks_to_github(window), save_peaks_tool)
    # window.Bind(wx.EVT_TOOL, lambda event: load_peaks_library(window), open_peaks_tool)

    toolbar.Realize()

    # Mac-specific styling
    if 'wxMac' in wx.PlatformInfo:
        # Remove default border and set background color
        toolbar.SetWindowStyle(toolbar.GetWindowStyle() | wx.BORDER_NONE)
        toolbar.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))

        # Add custom grey border on the bottom side
        border_panel = wx.Panel(toolbar)
        border_panel.SetBackgroundColour(wx.Colour(200, 200, 200))  # Light grey

        def on_toolbar_size(event):
            # Set the border panel to be full width but only 1px tall on the bottom
            size = toolbar.GetSize()
            border_panel.SetSize(0, size.height - 1, size.width, 1)
            event.Skip()

        toolbar.Bind(wx.EVT_SIZE, on_toolbar_size)

    return toolbar





# Bind the delete toolbar tools
def on_delete_all(window, event):
    # Get current sheet's row number
    sheet_name = window.sheet_combobox.GetValue()
    row_number = "0"

    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = match.group(1)

    results_table_key = f'Results Table{row_number}'

    # Check if table exists
    if results_table_key in window.Data:
        window.results_grid.DeleteRows(0, window.results_grid.GetNumberRows())
        window.Data[results_table_key]['Peak'] = {}
        save_state(window)


def on_delete_last(window, event):
    # Get current sheet's row number
    sheet_name = window.sheet_combobox.GetValue()
    row_number = "0"

    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = match.group(1)

    results_table_key = f'Results Table{row_number}'

    last_row = window.results_grid.GetNumberRows() - 1
    if last_row >= 0:
        window.results_grid.DeleteRows(last_row)

        if results_table_key in window.Data:
            peak_keys = list(window.Data[results_table_key]['Peak'].keys())
            if peak_keys:
                del window.Data[results_table_key]['Peak'][peak_keys[-1]]
        save_state(window)


def on_delete_first(window, event):
    # Get current sheet's row number
    sheet_name = window.sheet_combobox.GetValue()
    row_number = "0"

    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = match.group(1)

    results_table_key = f'Results Table{row_number}'

    if window.results_grid.GetNumberRows() > 0:
        window.results_grid.DeleteRows(0)

        if results_table_key in window.Data:
            peak_keys = list(window.Data[results_table_key]['Peak'].keys())
            if peak_keys:
                del window.Data[results_table_key]['Peak'][peak_keys[0]]

                # Renumber remaining peaks
                new_data = {}
                for i, (key, value) in enumerate(window.Data[results_table_key]['Peak'].items()):
                    if i > 0:  # Skip the first one that was deleted
                        new_key = f"Peak_{i - 1}"
                        new_data[new_key] = value

                window.Data[results_table_key]['Peak'] = new_data
        save_state(window)

def add_vertical_separator(toolbar, separators):
    separators.append(wx.StaticLine(toolbar, style=wx.LI_VERTICAL))
    separators[-1].SetSize((2, 24))
    toolbar.AddControl(separators[-1])

def bind_toolbar_events(window, open_file_tool, refresh_folder_tool, bkg_tool, fitting_tool,
                        # noise_analysis_tool,
                        save_tool, #save_plot_tool,
                        save_all_tool, toggle_Col_1_tool, export_tool, auto_be_button, id_tool
                        # toggle_legend_tool, toggle_fit_results_tool, toggle_residuals_tool, toggle_peak_fill_tool, plot_tool,
                        ):
    window.Bind(wx.EVT_TOOL, lambda event: open_xlsx_file(window), open_file_tool)
    window.Bind(wx.EVT_TOOL, lambda event: refresh_sheets(window, on_sheet_selected_wrapper), refresh_folder_tool)
    window.Bind(wx.EVT_TOOL, lambda event: window.on_open_background_window(), bkg_tool)
    window.Bind(wx.EVT_TOOL, lambda event: window.on_open_fitting_window(), fitting_tool)
    window.sheet_combobox.Bind(wx.EVT_COMBOBOX, lambda event: on_sheet_selected_wrapper(window, event))
    window.Bind(wx.EVT_TOOL, lambda event: on_save(window), save_tool)
    # window.Bind(wx.EVT_TOOL, lambda event: on_save_plot(window), save_plot_tool)
    window.Bind(wx.EVT_TOOL, lambda event: on_save_all_sheets(window, event), save_all_tool)
    window.Bind(wx.EVT_TOOL, lambda event: toggle_Col_1(window), toggle_Col_1_tool)
    window.Bind(wx.EVT_TOOL, lambda event: window.export_results(), export_tool)
    # window.be_correction_spinbox.Bind(wx.EVT_SPINCTRLDOUBLE, window.on_be_correction_change)
    window.Bind(wx.EVT_TOOL, window.on_auto_be, auto_be_button)
    window.Bind(wx.EVT_TOOL, lambda event: undo(window), window.undo_tool)
    window.Bind(wx.EVT_TOOL, lambda event: redo(window), window.redo_tool)
    window.Bind(wx.EVT_TOOL, window.open_periodic_table, id_tool)
    window.Bind(wx.EVT_TOOL, window.on_toggle_right_panel, window.toggle_right_panel_tool)


def create_vertical_toolbar(parent, frame):
    # v_toolbar = wx.ToolBar(parent, style=wx.TB_VERTICAL | wx.TB_DEFAULT_STYLE)
    v_toolbar = wx.ToolBar(parent, style=wx.TB_VERTICAL | wx.TB_FLAT )
    v_toolbar.SetToolBitmapSize(wx.Size(25, 25))

    # Check if running on macOS
    is_mac = 'wxMac' in wx.PlatformInfo

    # Only apply the custom border styling on Mac
    if is_mac:
        # Remove default border and set background color
        v_toolbar.SetWindowStyle(v_toolbar.GetWindowStyle() | wx.BORDER_NONE)
        v_toolbar.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))

        # Add custom grey border on the right side
        border_panel = wx.Panel(v_toolbar)
        border_panel.SetBackgroundColour(wx.Colour(200, 200, 200))  # Light grey

        def on_toolbar_size(event):
            # Set the border panel to be full height but only 1px wide on the right side
            size = v_toolbar.GetSize()
            border_panel.SetSize(size.width - 1, 0, 1, size.height)
            event.Skip()

        v_toolbar.Bind(wx.EVT_SIZE, on_toolbar_size)

    # Get the correct path for icons
    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, "Icons")

    # Create toggle toolbar instance
    frame.toggle_toolbar = ToggleToolbar(frame)

    # Add master toggle tool
    toggle_master_tool = v_toolbar.AddTool(wx.ID_ANY, 'Toggles',
                                           wx.Bitmap(os.path.join(icon_path, "Toggles-3.png"), wx.BITMAP_TYPE_PNG),
                                           shortHelp="Toggle Options")

    def show_toggle_toolbar(event):
        if not frame.toggle_toolbar.IsShown():
            pos = v_toolbar.GetScreenPosition()
            frame.toggle_toolbar.SetPosition((pos.x + v_toolbar.GetSize().width, pos.y))
            frame.toggle_toolbar.Show()
        else:
            frame.toggle_toolbar.Hide()
    frame.Bind(wx.EVT_TOOL, show_toggle_toolbar, toggle_master_tool)

    # Bind the toggle toolbar tools
    frame.Bind(wx.EVT_TOOL, lambda event: toggle_plot(frame), frame.toggle_toolbar.plot_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_toggle_peak_fill, frame.toggle_toolbar.peak_fill_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.plot_manager.toggle_y_axis(), frame.toggle_toolbar.y_axis_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.plot_manager.toggle_legend(), frame.toggle_toolbar.legend_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.plot_manager.toggle_fitting_results(), frame.toggle_toolbar.fit_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.plot_manager.toggle_residuals(frame), frame.toggle_toolbar.residuals_tool)


    # v_toolbar.AddSeparator()

    # Zoom tools
    zoom_in_tool = v_toolbar.AddTool(wx.ID_ANY, 'Zoom In',
                                     wx.Bitmap(os.path.join(icon_path, "ZoomIN-3.png"), wx.BITMAP_TYPE_PNG),
                                     shortHelp="Zoom In")
    zoom_out_tool = v_toolbar.AddTool(wx.ID_ANY, 'Zoom Out',
                                      wx.Bitmap(os.path.join(icon_path, "ZoomOUT-3.png"), wx.BITMAP_TYPE_PNG),
                                      shortHelp="Zoom Out")
    drag_tool = v_toolbar.AddTool(wx.ID_ANY, 'Drag',
                                  wx.Bitmap(os.path.join(icon_path, "Drag-25.png"), wx.BITMAP_TYPE_PNG),
                                  shortHelp="Drag Plot")

    # Plot limits tool
    plot_limits_tool = v_toolbar.AddTool(wx.ID_ANY, 'Plot Limits',
                                         wx.Bitmap(os.path.join(icon_path, "PlotLimits-3.png"), wx.BITMAP_TYPE_PNG),
                                         shortHelp="Set Plot Limits")

    # Bind the plot limits tool
    frame.Bind(wx.EVT_TOOL, lambda event: show_plot_limits_window(frame), plot_limits_tool)

    v_toolbar.AddSeparator()

    # BE adjustment tools
    high_be_increase_tool = v_toolbar.AddTool(wx.ID_ANY, 'High BE +',
                                              wx.Bitmap(os.path.join(icon_path, "Right-Red-25g.png"), wx.BITMAP_TYPE_PNG),
                                              shortHelp="Decrease High BE")
    high_be_decrease_tool = v_toolbar.AddTool(wx.ID_ANY, 'High BE -',
                                              wx.Bitmap(os.path.join(icon_path, "Left-Red-25g.png"), wx.BITMAP_TYPE_PNG),
                                              shortHelp="Increase High BE")

    # v_toolbar.AddSeparator()

    low_be_increase_tool = v_toolbar.AddTool(wx.ID_ANY, 'Low BE +',
                                             wx.Bitmap(os.path.join(icon_path, "Left-blue-25g.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Increase Low BE")
    low_be_decrease_tool = v_toolbar.AddTool(wx.ID_ANY, 'Low BE -',
                                             wx.Bitmap(os.path.join(icon_path, "Right-blue-25g.png"),
                                                       wx.BITMAP_TYPE_PNG),
                                             shortHelp="Decrease Low BE")

    # v_toolbar.AddSeparator()

    # Intensity adjustment tools
    high_int_increase_tool = v_toolbar.AddTool(wx.ID_ANY, 'High Int +',
                                               wx.Bitmap(os.path.join(icon_path, "Up-Red-25g.png"), wx.BITMAP_TYPE_PNG),
                                               shortHelp="Increase High Intensity")
    high_int_decrease_tool = v_toolbar.AddTool(wx.ID_ANY, 'High Int -',
                                               wx.Bitmap(os.path.join(icon_path, "Down-Red-25g.png"), wx.BITMAP_TYPE_PNG),
                                               shortHelp="Decrease High Intensity")

    # v_toolbar.AddSeparator()

    low_int_increase_tool = v_toolbar.AddTool(wx.ID_ANY, 'Low Int +',
                                              wx.Bitmap(os.path.join(icon_path, "Up-Blue-25g.png"), wx.BITMAP_TYPE_PNG),
                                              shortHelp="Increase Low Intensity")
    low_int_decrease_tool = v_toolbar.AddTool(wx.ID_ANY, 'Low Int -',
                                              wx.Bitmap(os.path.join(icon_path, "Down-Blue-25g.png"), wx.BITMAP_TYPE_PNG),
                                              shortHelp="Decrease Low Intensity")
    v_toolbar.AddSeparator()



    # Add text size increase/decrease tools
    text_increase_tool = v_toolbar.AddTool(wx.ID_ANY, 'Increase Font Size',
                                          wx.Bitmap(os.path.join(icon_path, "A+_25.png"), wx.BITMAP_TYPE_PNG),
                                          shortHelp="Increase All Font Sizes")
    text_decrease_tool = v_toolbar.AddTool(wx.ID_ANY, 'Decrease Font Size',
                                          wx.Bitmap(os.path.join(icon_path, "A-_25.png"), wx.BITMAP_TYPE_PNG),
                                          shortHelp="Decrease All Font Sizes")

    # v_toolbar.AddSeparator()

    # Add text annotation tool after other tools
    text_tool = v_toolbar.AddTool(wx.ID_ANY, 'Add Text',
        wx.Bitmap(os.path.join(icon_path, "AddText-25.png"), wx.BITMAP_TYPE_PNG),
        shortHelp="Add draggable text annotation")

    # Add to the binding section
    frame.Bind(wx.EVT_TOOL, lambda evt: add_draggable_text(frame), text_tool)

    labels_tool = v_toolbar.AddTool(wx.ID_ANY, 'Labels Manager',
                                    wx.Bitmap(os.path.join(icon_path, "ListText2-25.png"), wx.BITMAP_TYPE_PNG),
                                    # wx.Bitmap(os.path.join(icon_path, "AddText-25.png"), wx.BITMAP_TYPE_PNG),
                                    shortHelp="Open Labels Manager")
    frame.Bind(wx.EVT_TOOL, frame.open_labels_window, labels_tool)

    # v_toolbar.AddSeparator()

    v_toolbar.Realize()

    # Bind events to the frame
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('high_be', 'increase'), high_be_increase_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('high_be', 'decrease'), high_be_decrease_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('low_be', 'increase'), low_be_increase_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('low_be', 'decrease'), low_be_decrease_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('high_int', 'increase'), high_int_increase_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('high_int', 'decrease'), high_int_decrease_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('low_int', 'increase'), low_int_increase_tool)
    frame.Bind(wx.EVT_TOOL, lambda event: frame.adjust_plot_limits('low_int', 'decrease'), low_int_decrease_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_text_size_increase, text_increase_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_text_size_decrease, text_decrease_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_zoom_in_tool, zoom_in_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_zoom_out, zoom_out_tool)
    frame.Bind(wx.EVT_TOOL, frame.on_drag_tool, drag_tool)

    return v_toolbar


def create_statusbar(window):
    """
    Create a status bar for the main window.

    Args:
    window: The main application window.
    """
    # Create a status bar with two fields
    window.CreateStatusBar(2)

    # Set the widths of the status bar fields
    window.SetStatusWidths([-1, 200])

    # Set initial text for the status bar fields
    window.SetStatusText("Working Directory: " + window.Working_directory, 0)
    window.SetStatusText("BE: 0 eV, I: 0 CPS", 1)


def update_statusbar(window, message):
    """
    Update the first field of the status bar with a new message.

    Args:
    window: The main application window.
    message: The new message to display in the status bar.
    """
    window.SetStatusText("Working Directory: " + message)


# def open_manual2(window):
#     import os
#     current_dir = os.path.dirname(os.path.abspath(__file__))
#     root_dir = os.path.dirname(current_dir)
#     manual_path = os.path.join(root_dir, "Manual.pdf")
#     import webbrowser
#     webbrowser.open(manual_path)

def open_manualOLD1(window):
    import os
    import sys
    import webbrowser

    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, get the path of the executable
        application_path = os.path.dirname(sys.executable)
    else:
        # If the application is run as a script, get the path of the script
        application_path = os.path.dirname(os.path.abspath(__file__))

    manual_path = os.path.join(application_path, "Manual.pdf")
    webbrowser.open(manual_path)


def open_manualOLD2(window):
    import os
    import sys
    import webbrowser
    import platform
    import subprocess

    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, get the path of the executable
        application_path = os.path.dirname(sys.executable)
    else:
        # If the application is run as a script, get the path of the script
        application_path = os.path.dirname(os.path.abspath(__file__))

    manual_path = os.path.join(application_path, "Manual.pdf")

    # Use the appropriate method based on the operating system
    if platform.system() == 'Windows':
        os.startfile(manual_path)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.call(['open', manual_path])
    else:  # Linux and other Unix-like systems
        subprocess.call(['xdg-open', manual_path])


def open_manual(window):
    import os
    import sys
    import platform
    import subprocess
    import datetime

    # Log file setup
    log_path = os.path.expanduser("~/khervefitting_log.txt")
    with open(log_path, "a") as log:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log.write(f"\n--- {timestamp} ---\n")

        # Get the correct base path
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
            log.write(f"Running from binary: {base_path}\n")
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            log.write(f"Running from source: {base_path}\n")

        # Check possible manual locations
        possible_paths = [
            os.path.join(base_path, "Manual.pdf"),
            os.path.join(base_path, "resources", "Manual.pdf"),
            os.path.join(os.path.dirname(base_path), "Manual.pdf"),
            os.path.join(base_path, "..", "Resources", "Manual.pdf")  # Mac app bundle path
        ]

        log.write(f"Checking paths: {possible_paths}\n")

        manual_path = None
        for path in possible_paths:
            if os.path.exists(path):
                manual_path = path
                log.write(f"Found manual at: {path}\n")
                break

        if manual_path:
            try:
                if platform.system() == 'Darwin':
                    log.write(f"Attempting to open with 'open' command\n")
                    result = subprocess.run(['open', manual_path], capture_output=True, text=True)
                    log.write(f"Result: {result.returncode}, Output: {result.stdout}, Error: {result.stderr}\n")
                elif platform.system() == 'Windows':
                    log.write(f"Attempting to open with startfile\n")
                    os.startfile(manual_path)
                else:
                    log.write(f"Attempting to open with xdg-open\n")
                    result = subprocess.run(['xdg-open', manual_path], capture_output=True, text=True)
                    log.write(f"Result: {result.returncode}, Output: {result.stdout}, Error: {result.stderr}\n")
                return True
            except Exception as e:
                log.write(f"Error opening manual: {e}\n")
                return False
        else:
            log.write("Manual not found in any expected locations\n")
            return False

def show_plot_limits_window(window):
    from libraries.PlotConfig import PlotLimitsWindow
    if not hasattr(window, 'plot_limits_window') or not window.plot_limits_window:
        window.plot_limits_window = PlotLimitsWindow(window)
    window.plot_limits_window.Show()
    window.plot_limits_window.Raise()


class ToggleToolbar(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, style=wx.FRAME_NO_TASKBAR | wx.FRAME_FLOAT_ON_PARENT)

        self.toolbar = wx.ToolBar(self, style=wx.TB_HORIZONTAL | wx.TB_FLAT)
        self.toolbar.SetToolBitmapSize(wx.Size(25, 25))

        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icons")

        # Add tools
        self.plot_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Plot',
                                              wx.Bitmap(os.path.join(icon_path, "scatter-plot-25.png"),
                                                        wx.BITMAP_TYPE_PNG))
        self.peak_fill_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Peak Fill',
                                                   wx.Bitmap(os.path.join(icon_path, "STO-25-2.png"),
                                                             wx.BITMAP_TYPE_PNG))
        self.y_axis_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Y Axis',
                                                wx.Bitmap(os.path.join(icon_path, "Y-25.png"), wx.BITMAP_TYPE_PNG))
        self.legend_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Legend',
                                                wx.Bitmap(os.path.join(icon_path, "Legend-25.png"), wx.BITMAP_TYPE_PNG))
        self.fit_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Fit Results',
                                             wx.Bitmap(os.path.join(icon_path, "ToggleFit-25.png"), wx.BITMAP_TYPE_PNG))
        self.residuals_tool = self.toolbar.AddTool(wx.ID_ANY, 'Toggle Residuals',
                                                   wx.Bitmap(os.path.join(icon_path, "Res-25.png"), wx.BITMAP_TYPE_PNG))

        self.toolbar.Realize()
        self.SetSize(self.toolbar.GetBestSize())

        # Bind close event
        self.Bind(wx.EVT_KILL_FOCUS, self.on_lose_focus)



    def on_lose_focus(self, event):
        self.Hide()
        event.Skip()


class DeleteToolbar(wx.Frame):
    def __init__(self, parent):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "Icons")
        super().__init__(parent, style=wx.FRAME_NO_TASKBAR | wx.FRAME_FLOAT_ON_PARENT)

        self.toolbar = wx.ToolBar(self, style=wx.TB_HORIZONTAL | wx.TB_FLAT)
        self.toolbar.SetToolBitmapSize(wx.Size(25, 25))

        # Add tools
        self.delete_all_tool = self.toolbar.AddTool(wx.ID_ANY, 'Delete All Results',
                                                    wx.Bitmap(os.path.join(icon_path, "AllRow-25.png"),
                                                              wx.BITMAP_TYPE_PNG), shortHelp="Delete All Rows of "
                                                                                             "the Results Grid")

        self.delete_last_tool = self.toolbar.AddTool(wx.ID_ANY, 'Delete Last Row',
                                                     wx.Bitmap(os.path.join(icon_path, "LastRow-25.png"),
                                                               wx.BITMAP_TYPE_PNG), shortHelp="Delete Last Row of "
                                                                                              "the Results Grid")

        self.delete_first_tool = self.toolbar.AddTool(wx.ID_ANY, 'Delete First Row',
                                                      wx.Bitmap(os.path.join(icon_path, "TopRow-25.png"),
                                                                wx.BITMAP_TYPE_PNG), shortHelp="Delete First Row of "
                                                                                               "the Results Grid")

        self.toolbar.Realize()
        self.SetSize(self.toolbar.GetBestSize())

        self.Bind(wx.EVT_KILL_FOCUS, self.on_lose_focus)

    def on_lose_focus(self, event):
        self.Hide()
        event.Skip()