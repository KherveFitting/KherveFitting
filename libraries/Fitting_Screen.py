
import re
import wx
from Functions import fit_peaks, remove_peak
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import matplotlib.pyplot as plt
import lmfit
from libraries.Peak_Functions import BackgroundCalculations
from libraries.Save import save_state
from libraries.Plot_Operations import PlotManager
from libraries.Open import load_library_data

class FittingWindow(wx.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent  # Store reference to MainFrame


        self.SetTitle("Peak Fitting")
        self.SetSize((325, 495))  # Increased height to accommodate new elements
        self.SetMinSize((325, 495))
        self.SetMaxSize((325, 495))

        #305 480

        self.init_ui()

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.library_data = self.parent.library_data
        self.doublet_splittings = self.load_doublet_splittings(self.parent.library_data)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)




    def init_ui(self):
        panel = wx.Panel(self)
        # panel.SetBackgroundColour(wx.Colour(255, 255, 255))
        # panel.SetBackgroundColour(wx.Colour(250, 250, 250))
        panel.SetBackgroundColour(wx.Colour(230, 250, 250))

        main_sizer = wx.BoxSizer(wx.VERTICAL)


        notebook = wx.Notebook(panel)
        notebook.SetBackgroundColour(wx.Colour(240, 250, 250))
        self.init_background_tab(notebook)
        self.init_fitting_tab(notebook)

        main_sizer.Add(notebook, 1, wx.EXPAND,border=5)
        panel.SetSizer(main_sizer)

        notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_tab_change)

        self.parent.background_tab_selected = True
        self.parent.peak_fitting_tab_selected = False
        self.parent.show_hide_vlines()
        self.parent.deselect_all_peaks()



    def init_background_tab(self, notebook):
        """Initialize the background tab in the notebook."""
        self.parent.background_method = "Multi-Regions Smart"
        self.parent.offset_h = 0
        self.parent.offset_l = 0


        self.background_panel = wx.Panel(notebook)
        background_sizer = wx.GridBagSizer(hgap=0, vgap=0)

        method_label = wx.StaticText(self.background_panel, label="Method:")
        self.method_combobox = wx.ComboBox(self.background_panel, choices=["Multi-Regions Smart", "Smart", "Shirley",
                                            "Linear", '1x U4-Tougaard', "2x U4-Tougaard", "3x U4-Tougaard"],
                                           style=wx.CB_READONLY)
        method_index = self.method_combobox.FindString(self.parent.background_method)
        self.method_combobox.SetSelection(method_index)
        self.method_combobox.Bind(wx.EVT_COMBOBOX, self.on_bkg_method_change)

        info_button = self.create_info_button(self.background_panel,
                                              self.get_background_description(self.parent.background_method))

        offset_h_label = wx.StaticText(self.background_panel, label="Offset (H):")
        self.offset_h_text = wx.TextCtrl(self.background_panel, value=str(self.parent.offset_h))
        self.offset_h_text.Bind(wx.EVT_TEXT, self.on_offset_h_change)

        offset_l_label = wx.StaticText(self.background_panel, label="Offset (L):")
        self.offset_l_text = wx.TextCtrl(self.background_panel, value=str(self.parent.offset_l))
        self.offset_l_text.Bind(wx.EVT_TEXT, self.on_offset_l_change)

        averaging_points_label = wx.StaticText(self.background_panel, label="Averaging Points:")
        self.averaging_points_text = wx.TextCtrl(self.background_panel, value="5")
        self.averaging_points_text.Bind(wx.EVT_TEXT, self.on_averaging_points_change)

        self.cross_section_label = wx.StaticText(self.background_panel, label = 'Tougaard1: B,C,D,T0')
        self.cross_section = wx.TextCtrl(self.background_panel, value="2866,1643,1,0")
        self.cross_section.Bind(wx.EVT_TEXT, self.on_cross_section_change)

        # Add two more labels and TextCtrls
        self.cross_section2_label = wx.StaticText(self.background_panel, label='Tougaard2: B,C,D,T0')
        self.cross_section2 = wx.TextCtrl(self.background_panel, value="2866,1643,1,0")
        self.cross_section2.Bind(wx.EVT_TEXT, self.on_cross_section2_change)

        self.cross_section3_label = wx.StaticText(self.background_panel, label='Tougaard3: B,C,D,T0')
        self.cross_section3 = wx.TextCtrl(self.background_panel, value="2866,1643,1,0")
        self.cross_section3.Bind(wx.EVT_TEXT, self.on_cross_section3_change)

        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Background' in self.parent.Data['Core levels'][sheet_name]:
            bg_data = self.parent.Data['Core levels'][sheet_name]['Background']
            saved_values = [
                bg_data.get('Tougaard_B', 2866),
                bg_data.get('Tougaard_C', 1643),
                bg_data.get('Tougaard_D', 1),
                bg_data.get('Tougaard_T0', 0)
            ]
            saved_values2 = [
                bg_data.get('Tougaard_B2', 2866),
                bg_data.get('Tougaard_C2', 1643),
                bg_data.get('Tougaard_D2', 1),
                bg_data.get('Tougaard_T02', 0)
            ]
            saved_values3 = [
                bg_data.get('Tougaard_B2', 2866),
                bg_data.get('Tougaard_C2', 1643),
                bg_data.get('Tougaard_D2', 1),
                bg_data.get('Tougaard_T02', 0)
            ]
            self.cross_section.SetValue(','.join(map(str, saved_values)))
            self.cross_section2.SetValue(','.join(map(str, saved_values2)))
            self.cross_section2.SetValue(','.join(map(str, saved_values3)))

        background_button = wx.Button(self.background_panel, label="Create\nBackground")
        background_button.SetMinSize((125, 40))
        background_button.Bind(wx.EVT_BUTTON, self.on_background)

        clear_background_button = wx.Button(self.background_panel, label="Clear\nAll")
        clear_background_button.SetMinSize((125, 40))
        clear_background_button.Bind(wx.EVT_BUTTON, self.on_clear_background)

        reset_vlines_button = wx.Button(self.background_panel, label="Reset \nVertical Lines")
        reset_vlines_button.SetMinSize((125, 40))
        reset_vlines_button.Bind(wx.EVT_BUTTON, self.on_reset_vlines)

        clear_between_vlines_button = wx.Button(self.background_panel, label="Clear Between\nVertical Lines")
        clear_between_vlines_button.SetMinSize((125, 40))
        clear_between_vlines_button.Bind(wx.EVT_BUTTON, self.on_clear_between_vlines)

        clear_background_only_button = wx.Button(self.background_panel, label="Clear\nBackground")
        clear_background_only_button.SetMinSize((125, 40))
        clear_background_only_button.Bind(wx.EVT_BUTTON, self.on_clear_background_only)

        self.tougaard_fit_btn = wx.Button(self.background_panel, label="Create Tougaard\n Model")
        self.tougaard_fit_btn.SetMinSize((125, 40))
        self.tougaard_fit_btn.Bind(wx.EVT_BUTTON, lambda evt: TougaardFitWindow(self).Show())

        # Layout Background Tab
        background_sizer.Add(method_label, pos=(0, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        background_sizer.Add(self.method_combobox, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=0)
        background_sizer.Add(info_button, pos=(1, 1), flag=wx.ALL, border=0)
        background_sizer.Add(offset_h_label, pos=(2, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        background_sizer.Add(self.offset_h_text, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(offset_l_label, pos=(3, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        background_sizer.Add(self.offset_l_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(averaging_points_label, pos=(4, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        background_sizer.Add(self.averaging_points_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section_label,  pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section2_label,  pos=(6, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section2, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section3_label,  pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.cross_section3, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=5)


        background_sizer.Add(reset_vlines_button, pos=(9, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(clear_between_vlines_button, pos=(9, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(self.tougaard_fit_btn, pos=(10, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(clear_background_only_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(background_button, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=5)
        background_sizer.Add(clear_background_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=5)

        self.background_panel.SetSizer(background_sizer)
        notebook.AddPage(self.background_panel, "Background")

        self.update_tougaard_controls_visibility(self.parent.background_method)

    def init_fitting_tab(self, notebook):
        """Initialize the peak fitting tab in the notebook."""
        self.fitting_panel = wx.Panel(notebook)
        fitting_sizer = wx.GridBagSizer(hgap=0, vgap=0)


        self.model_combobox = CustomComboBox(self.fitting_panel, style=wx.CB_READONLY)

        # Set items and green items
        items = ["Preferred Models-----",
                 "GL (Area)",
                 "SGL (Area)",
                 "LA (Area, \u03c3/\u03b3, \u03b3)",
                 "Voigt (Area, L/G, \u03c3)",
                 "Voigt (Area, L/G, \u03c3, skew)",
                 "Others---------------",
                 "Area Based----------",
                 # "GL (Area)",
                 # "SGL (Area)",
                 "Voigt (Area, \u03c3, \u03b3)",
                 # "Voigt (Area, L/G, \u03c3)",
                 "LA (Area, \u03c3, \u03b3)",
                 # "LA (Area, \u03c3/\u03b3, \u03b3)",
                 "LA*G (Area, \u03c3/\u03b3, \u03b3)",
                 "Pseudo-Voigt (Area)",
                 "ExpGauss.(Area, \u03c3, \u03b3)",
                 "Height Based---------",
                 "GL (Height)",
                 "SGL (Height)",
                 # "Under Test -----------",
                 ]


        green_items = ["Others---------------","Area Based----------", "Height Based---------", "Under Test "
                                                                                               "-----------", "Preferred Models-----"]
        self.model_combobox.SetItems(items)
        self.model_combobox.SetGreenItems(green_items)

        # Set default selection to "GL (Area)"
        default_index = items.index("GL (Area)")
        self.model_combobox.SetSelection(default_index)
        self.model_combobox.SetValue("GL (Area)")
        self.parent.set_fitting_method("GL (Area)")



        # model_index = self.model_combobox.FindString(self.parent.selected_fitting_method)
        # self.model_combobox.SetSelection(model_index)
        self.model_combobox.Bind(wx.EVT_COMBOBOX, self.on_method_change)

        info_button = self.create_info_button(self.fitting_panel,
                                              self.get_fitting_description(self.parent.selected_fitting_method))

        self.optimization_method = wx.ComboBox(self.fitting_panel, choices=[
            "leastsq",
            "least_squares",
            # "differential_evolution",
            "nelder",
            # "lbfgsb",
            "powell",
            # "cg",
            # "newton",
            "cobyla",
            # "basinhopping",
            # "slsqp",
            # "tnc",
            # "trust-krylov",
            "trust-constr"
            # "dogleg",
            # "shgo",
            # "dual_annealing"
        ], style=wx.CB_READONLY)
        self.optimization_method.SetSelection(1)  # Default value

        self.max_iter_spin = wx.SpinCtrl(self.fitting_panel, value=str(self.parent.max_iterations), min=20, max=200)
        self.max_iter_spin.Bind(wx.EVT_SPINCTRL, self.on_max_iter_change)

        self.fit_iterations_spin = wx.SpinCtrl(self.fitting_panel, value="20", min=3, max=100)

        self.r_squared_label = wx.StaticText(self.fitting_panel, label="R²:")
        self.r_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)


        # self.chi_squared_label = wx.StaticText(self.fitting_panel, label="Chi²:")
        # self.chi_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        # self.rsd_label = wx.StaticText(self.fitting_panel, label="RSD:")
        # self.rsd_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)


        self.red_chi_squared_label = wx.StaticText(self.fitting_panel, label="Red. Chi²:")
        self.red_chi_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        # self.actual_iter_label = wx.StaticText(self.fitting_panel, label="Actual Iterations:")
        # self.actual_iter_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        self.current_fit_label = wx.StaticText(self.fitting_panel, label="Current Fit:")
        self.current_fit_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        fit_report_button = wx.Button(self.fitting_panel, label="Report")
        fit_report_button.SetMinSize((125, 30))
        fit_report_button.Bind(wx.EVT_BUTTON, self.on_fit_report)

        add_peak_button = wx.Button(self.fitting_panel, label="Add 1 Peak\nSinglet")
        add_peak_button.SetMinSize((125, 40))
        add_peak_button.Bind(wx.EVT_BUTTON, self.on_add_peak)

        add_doublet_button = wx.Button(self.fitting_panel, label="Add 2 Peaks\nDoublet")
        add_doublet_button.SetMinSize((125, 40))
        add_doublet_button.Bind(wx.EVT_BUTTON, self.on_add_doublet)

        remove_peak_button = wx.Button(self.fitting_panel, label="Remove\nLast Peak")
        remove_peak_button.SetMinSize((125, 40))
        remove_peak_button.Bind(wx.EVT_BUTTON, self.on_remove_peak)

        export_button = wx.Button(self.fitting_panel, label="Export to\nResults Grid")
        export_button.SetMinSize((125, 40))
        export_button.Bind(wx.EVT_BUTTON, self.on_export_results)

        fit_button = wx.Button(self.fitting_panel, label="Fit \nOne Time")
        fit_button.SetMinSize((125, 40))
        fit_button.Bind(wx.EVT_BUTTON, self.on_fit_peaks)

        fit_multi_button = wx.Button(self.fitting_panel, label="Fit \nN# Times")
        fit_multi_button.SetMinSize((125, 40))
        fit_multi_button.Bind(wx.EVT_BUTTON, self.on_fit_multi)

        fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Fitting Model:"), pos=(0, 0),
                          flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.model_combobox, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=0)
        fitting_sizer.Add(info_button, pos=(1, 1), flag=wx.ALL, border=0)

        fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Method:"), pos=(2, 0),
                          flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.optimization_method, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=0)

        fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Convergence Limit:"), pos=(3, 0),
                          flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.max_iter_spin, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="N# of Iterations:"), pos=(4, 0),
                          flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.fit_iterations_spin, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=5)

        fitting_sizer.Add(self.r_squared_label, pos=(5, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.r_squared_text, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=5)
        # fitting_sizer.Add(self.rsd_label, pos=(6, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        # fitting_sizer.Add(self.rsd_text, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(self.red_chi_squared_label, pos=(6, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.red_chi_squared_text, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=5)

        # fitting_sizer.Add(self.actual_iter_label, pos=(8, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        # fitting_sizer.Add(self.actual_iter_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(self.current_fit_label, pos=(7, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
        fitting_sizer.Add(self.current_fit_text, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=5)

        fitting_sizer.Add(fit_report_button, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)

        fitting_sizer.Add(add_peak_button, pos=(9, 0), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(add_doublet_button, pos=(9, 1), flag=wx.ALL | wx.EXPAND, border=5)

        fitting_sizer.Add(remove_peak_button, pos=(10, 0), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(export_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=5)

        fitting_sizer.Add(fit_button, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=5)
        fitting_sizer.Add(fit_multi_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=5)

        self.fitting_panel.SetSizer(fitting_sizer)
        notebook.AddPage(self.fitting_panel, "Peak Fitting")

    def on_fit_report(self, event):
        if hasattr(self.parent, 'fit_results') and 'result' in self.parent.fit_results:
            dlg = wx.Frame(self, title="Fit Report", size=(600, 400))
            text = wx.TextCtrl(dlg, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)

            result = self.parent.fit_results['result']
            report = result.fit_report()
            report += "\n\nAdditional Information:"
            report += f"\nNumber of function evaluations: {result.nfev}"
            report += f"\nNumber of variables: {result.nvarys}"
            report += f"\nNumber of data points: {result.ndata}"
            report += f"\nDegrees of freedom: {result.nfree}"
            report += f"\nAborted: {result.aborted}"
            report += f"\nError flags: {result.errorbars}"
            report += f"\nCovariance matrix: \n{result.covar}"

            text.SetValue(report)

            sizer = wx.BoxSizer(wx.VERTICAL)
            sizer.Add(text, 1, wx.EXPAND | wx.ALL, 5)

            dlg.SetSizer(sizer)
            dlg.Show()
        else:
            wx.MessageBox("No fit report available. Please perform a fit first.", "No Report")

    def get_optimization_method(self):
        selection = self.optimization_method.GetValue()
        return selection.split()[0]  # Get just the method name without description

    def on_method_change_OLD(self, event):
        new_method = self.model_combobox.GetValue()
        old_method = self.parent.selected_fitting_method
        self.parent.set_fitting_method(new_method)
        self.update_fitting_info_button()

    def on_method_change(self, event):
        new_method = self.model_combobox.GetValue()
        if "skew" in new_method:
            self.optimization_method.SetValue("leastsq")
            self.optimization_method.Enable(True)
        else:
            self.optimization_method.SetValue("least_squares")
            self.optimization_method.Enable(True)
        self.parent.set_fitting_method(new_method)
        self.update_fitting_info_button()

    def on_max_iter_change(self, event):
        self.parent.set_max_iterations(self.max_iter_spin.GetValue())

    # def on_method_change(self, event):
    #     new_method = self.model_combobox.GetValue()
    #     self.parent.set_fitting_method(new_method)

    def on_bkg_method_change(self, event):
        new_method = self.method_combobox.GetValue()
        self.parent.set_background_method(new_method)
        self.update_background_info_button()
        self.update_tougaard_controls_visibility(new_method)

    def update_tougaard_controls_visibility(self,new_method):
        if new_method.startswith("1x U4-Tougaard"):
            self.cross_section.Enable(True)
            self.cross_section_label.Enable(True)
            self.tougaard_fit_btn.Enable(True)
            self.cross_section2.Enable(False)
            self.cross_section2_label.Enable(False)
            self.cross_section3.Enable(False)
            self.cross_section3_label.Enable(False)
        elif new_method.startswith("2x U4-Tougaard"):
            self.cross_section.Enable(True)
            self.cross_section_label.Enable(True)
            self.tougaard_fit_btn.Enable(True)
            self.cross_section2.Enable(True)
            self.cross_section2_label.Enable(True)
            self.cross_section3.Enable(False)
            self.cross_section3_label.Enable(False)
        elif new_method.startswith("3x U4-Tougaard"):
            self.cross_section.Enable(True)
            self.cross_section_label.Enable(True)
            self.tougaard_fit_btn.Enable(True)
            self.cross_section2.Enable(True)
            self.cross_section2_label.Enable(True)
            self.cross_section3.Enable(True)
            self.cross_section3_label.Enable(True)
        else:
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(False)
            self.cross_section2.Enable(False)
            self.cross_section2_label.Enable(False)
            self.cross_section3.Enable(False)
            self.cross_section3_label.Enable(False)
        self.background_panel.Layout()


    def on_clear_background_only(self, event):
        self.parent.plot_manager.clear_background_only(self.parent)

    def on_averaging_points_change(self, event):
        try:
            self.parent.averaging_points = int(self.averaging_points_text.GetValue())
        except ValueError:
            pass

    def on_reset_vlines(self, event):
        self.parent.vline1 = None
        self.parent.vline2 = None
        self.parent.show_hide_vlines()
        self.parent.plot_manager.clear_and_replot(self.parent)


    def on_add_peak(self, event):
        self.parent.add_peak_params()

    def on_remove_peak(self, event):
        remove_peak(self.parent)

    def on_fit_multi(self, event):
        if self.parent.peak_params_grid.GetNumberRows() == 0:
            wx.MessageBox("No peaks to fit. Add at least one peak first.", "Error", wx.OK | wx.ICON_ERROR)
            return
        save_state(self.parent)
        iterations = self.fit_iterations_spin.GetValue()
        for i in range(1, iterations + 1):
            self.current_fit_text.SetValue(f"{i}/{iterations}")
            result = fit_peaks(self.parent, self.parent.peak_params_grid)
            if result:
                r_squared, rsd, red_chi_square = result
                self.update_fit_indicators(r_squared, rsd, red_chi_square)
                # self.actual_iter_text.SetValue(str(self.parent.fit_results['nfev']))
            self.parent.clear_and_replot()
            wx.Yield()
        self.current_fit_text.SetValue("Complete")
        save_state(self.parent)

    def on_fit_peaks(self, event):
        save_state(self.parent)
        result = fit_peaks(self.parent, self.parent.peak_params_grid)
        if result:
            r_squared, rsd, red_chi_squared = result
            self.update_fit_indicators(r_squared, rsd, red_chi_squared)
        else:
            print("Fitting failed or was cancelled.")

    def update_fit_indicators(self, r_squared, rsd, red_chi_squared):
        self.r_squared_text.SetValue(f"{r_squared:.5f}")
        # self.rsd_text.SetValue(f"{rsd:.5f}")
        self.red_chi_squared_text.SetValue(f"{red_chi_squared:.2f}")
        self.Layout()

    def on_export_results(self, event):
        self.parent.export_results()

    def load_doublet_splittings(self, library_data):
        splittings = {}
        return
        for key, value in library_data.items():
            element, orbital = key
            if orbital.endswith(('p', 'd', 'f')):
                try:
                    splittings[f"{element}{orbital}"] = value['C-Al1486']['ds']  # Must match Excel exactly
                    print(f'Loading Splittings for {element} {orbital} Value {value['C-Al1486']['ds']}, split'
                          f' {splittings[f"{element}{orbital}"]}')
                    # print(f'Element: {element} , Orbital:{orbital}, Split: {splittings}')
                except KeyError:
                    print(f"Error\nAvailable instruments for {element}{orbital}: {list(value.keys())}")
                    # print(f"Full value data: {value}")
                    continue
        return splittings

    def get_doublet_splitting(self, element, orbital, instrument):
        # Extract number and letter from orbital (e.g., "3d" -> "3" and "d")
        orbital_match = re.match(r'(\d)([spdf])', orbital)
        if orbital_match:
            num, letter = orbital_match.groups()
            key = (element, f"{num}{letter}")

            if key in self.library_data and 'C-Al1486' in self.library_data[key]:
                print(f'Found splitting for {element}{num}{letter}: {self.library_data[key]["C-Al1486"]["ds"]}')
                return self.library_data[key]['C-Al1486']['ds']
            else:
                print(f'No splitting found for key: {key}')

        print(f"Warning: No doublet splitting found for {element} {orbital}")
        return 0.0


    def on_add_doublet(self, event):
        if self.parent.bg_min_energy is None or self.parent.bg_max_energy is None:
            wx.MessageBox("Please create a background first.", "No Background", wx.OK | wx.ICON_WARNING)
            return

        sheet_name = self.parent.sheet_combobox.GetValue()
        # first_word = sheet_name.split()[0]  # Get the first word of the sheet name
        # orbital = re.search(r'[spdf]', first_word)

        # Looking into library
        first_word = sheet_name.split()[0]
        print(f'First word {first_word}')
        # element = re.match(r'([A-Z][a-z]*)', first_word).group(1)
        orbital = re.search(r'([2-5][spdf])', first_word)

        if not orbital:
            wx.MessageBox("Cannot fit doublet peak on a S orbital core level.",
                          "Error", wx.OK | wx.ICON_ERROR)
            return

        orbital = orbital.group()


        if orbital[-1] == 's':
            self.parent.add_peak_params()
        else:
            first_peak = self.parent.add_peak_params()
            second_peak = self.parent.add_peak_params()


            self.parent.peak_fill_types[second_peak] = self.parent.peak_fill_types[first_peak]
            self.parent.peak_hatch_patterns[second_peak] = self.parent.peak_hatch_patterns[first_peak]
            self.hatch_density = 2

            # Set constraints for the second peak
            row1 = first_peak * 2
            row2 = second_peak * 2

            # L/G ratio constraint
            lg_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 5, lg_constraint)

            # FWHM constraint
            if any(element in sheet_name for element in ['Ti2p', 'V2p']) and any(
                    x in self.parent.selected_fitting_method for x in ["LA", "GL", "SGL"]):
                fwhm_constraint = "0.3:3.5"  # Independent FWHM for Ti2p and V2p
            else:
                if "Voigt" in self.parent.selected_fitting_method:
                    fwhm_constraint = "0.3:3.5"
                else:
                    fwhm_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 4, fwhm_constraint)

            # Height constraint
            height_factor = {'p': 0.5, 'd': 0.667, 'f': 0.75}
            height_constraint = f"{chr(65 + first_peak)}*{height_factor[orbital[1]]}#0.05"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 3, height_constraint)

            # Area constraint
            Area_factor = {'p': 0.5, 'd': 0.667, 'f': 0.75}
            area_constraint = f"{chr(65 + first_peak)}*{height_factor[orbital[-1]]}#0.05"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 6, area_constraint)

            # Position constraint
            # splitting = self.doublet_splittings.get(first_word, 0)
            element = re.match(r'([A-Z][a-z]*)', first_word).group(1)
            splitting = self.get_doublet_splitting(element, orbital, self.parent.current_instrument)

            position_constraint = f"{chr(65 + first_peak)}+{splitting}#0.2"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 2, position_constraint)

            # Sigma constraint
            if any(element in sheet_name for element in ['Ti2p', 'V2p']) and any(
                    x in self.parent.selected_fitting_method for x in ["Voigt"]):
                sigma_constraint = "0.01:3"  # Independent FWHM for Ti2p and V2p
            else:
                sigma_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 7, sigma_constraint)

            # Gamma constraint
            gamma_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 8, gamma_constraint)

            # Skew constraint
            skew_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 9, skew_constraint)

            # Calculate peak numbers
            peak_number1 = first_peak + 1
            peak_number2 = second_peak + 1

            # Set peak names
            if orbital[-1] == 'p':
                peak1_name = f"{first_word}3/2 p{peak_number1}"
                peak2_name = f"{first_word}1/2 p{peak_number2}"
            elif orbital[-1] == 'd':
                peak1_name = f"{first_word}5/2 p{peak_number1}"
                peak2_name = f"{first_word}3/2 p{peak_number2}"
            elif orbital[-1] == 'f':
                peak1_name = f"{first_word}7/2 p{peak_number1}"
                peak2_name = f"{first_word}5/2 p{peak_number2}"

            self.parent.peak_params_grid.SetCellValue(row1, 1, peak1_name)
            self.parent.peak_params_grid.SetCellValue(row2, 1, peak2_name)

            # Position the second peak
            first_peak_position = float(self.parent.peak_params_grid.GetCellValue(row1, 2))
            second_peak_position = first_peak_position + splitting
            self.parent.peak_params_grid.SetCellValue(row2, 2, f"{second_peak_position:.2f}")

            # Update window.Data with new constraints and names
            if 'Fitting' in self.parent.Data['Core levels'][sheet_name] and 'Peaks' in \
                    self.parent.Data['Core levels'][sheet_name]['Fitting']:
                peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                new_peaks = {}
                for i, (key, value) in enumerate(peaks.items()):
                    if i == first_peak:
                        new_peaks[peak1_name] = value
                        new_peaks[peak1_name]['Name'] = peak1_name
                    elif i == second_peak:
                        new_peaks[peak2_name] = value
                        new_peaks[peak2_name]['Name'] = peak2_name
                        new_peaks[peak2_name]['Position'] = second_peak_position
                        new_peaks[peak2_name]['Constraints'] = {
                            'Position': position_constraint,
                            'Height': height_constraint,
                            'FWHM': fwhm_constraint,
                            'L/G': lg_constraint,
                            'Area': area_constraint,
                            'Sigma': sigma_constraint,
                            'Gamma': gamma_constraint,
                            'Skew': skew_constraint
                        }
                    else:
                        new_peaks[key] = value
                self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks

        self.parent.clear_and_replot()

    def on_background(self, event):
        if self.parent.bg_min_energy is None or self.parent.bg_max_energy is None:
            sheet_name = self.parent.sheet_combobox.GetValue()
            x_values = self.parent.Data['Core levels'][sheet_name]['B.E.']
            self.parent.bg_min_energy = min(x_values) + 0.2
            self.parent.bg_max_energy = max(x_values) - 0.2

            if 'Background' in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = self.parent.bg_min_energy
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = self.parent.bg_max_energy

        self.parent.plot_manager.plot_background(self.parent)

        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            core_level_data = self.parent.Data['Core levels'][sheet_name]
            if 'Background' in core_level_data:
                self.parent.bg_min_energy = core_level_data['Background']['Bkg Low']
                self.parent.bg_max_energy = core_level_data['Background']['Bkg High']

        save_state(self.parent)

    def on_clear_background(self, event):
        self.parent.plot_manager.clear_background(self.parent)
        self.parent.bg_min_energy = None
        self.parent.bg_max_energy = None
        self.parent.plot_data()
        save_state(self.parent)

    def on_offset_h_change(self, event):
        offset_h_value = self.offset_h_text.GetValue()
        self.parent.set_offset_h(offset_h_value)

    def on_offset_l_change(self, event):
        offset_l_value = self.offset_l_text.GetValue()
        self.parent.set_offset_l(offset_l_value)


    def get_background_description(self, method):
        descriptions = {
            "Multi-Regions Smart": "Same as the Smart background but not restricted to a single region",
            "Smart": "Single region Smart background estimation. It applies a Shirley background "
                     "when the intensity is going up and a linear background when the intensity is "
                     "going down. If the calculated background is above the data then the "
                     "background is set equal to the data.",
            "Shirley": "Iterative background calculation. Reliable for on positive background when "
            "the data contains symmetrical peak.",
            "Linear": "Simple linear background. Usually used on negative background"
            "Tougaard: U4 Tougaard background for Advanced users"
        }
        return descriptions.get(method, "No description available")

    def get_fitting_description(self, model):
        descriptions = {
            "GL (Height)": "Gaussian-Lorentzian product function (height-based).\nEquation: I(x) = H * [exp(-ln(2) * ((x-x₀)/σ)²) * (1 / (1 + ((x-x₀)/γ)²))]",
            "SGL (Height)": "Sum of Gaussian and Lorentzian functions (height-based).\nEquation: I(x) = H * [m * exp(-ln(2) * ((x-x₀)/σ)²) + (1-m) / (1 + ((x-x₀)/γ)²)]",
            "GL (Area)": "Gaussian-Lorentzian product function (area-based).\nEquation: Similar to GL (Height), but normalized for area",
            "SGL (Area)": "Sum of Gaussian and Lorentzian functions (area-based).\nEquation: Similar to SGL (Height), but normalized for area",
            "Pseudo-Voigt (Area)": "Linear combination of Gaussian and Lorentzian.\nEquation: I(x) = A * [η * L(x) + (1-η) * G(x)]",
            "Voigt (Area, L/G, \u03c3, skew)": "Convolution of Gaussian and Lorentzian.\nEquation: I(x) = A * ∫G(x')L(x-x')dx'",
            "Voigt (Area,\u03c3, \u03b3)": "Voigt function with separate Gaussian and Lorentzian widths.\nEquation: "
                                           "I(x) = A * ∫G(x', σ)L(x-x', γ)dx'",
            "Voigt (Area,\u03c3, \u03b3)": "Voigt function with separate Gaussian and Lorentzian widths.\nEquation: "
                                           "I(x) = A * ∫G(x', σ)L(x-x', γ)dx'",
            "Voigt (Area, L/G, \u03c3, skew)": "Voigt function with separate Gaussian and Lorentzian widths.\nEquation: "
                                           "I(x) = A * ∫G(x', σ)L(x-x', γ)dx'",
            "ExpGauss.(Area, \u03c3, \u03b3)": "Gaussian shape model with asymmetric side. The asymmetry is modelled"
                                               "using an exponential decay as per equation: ",
            "LA (Area, \u03c3, \u03b3)": "Asymmetrical lorentzian similar to the model used in casa XPS",
            "LA (Area, \u03c3/\u03b3, \u03b3)": "Asymmetrical lorentzian with asymmetru controlled by the ratio "
                                                "between \u03c3 and \u03b3 similar to the model used in casa XPS",
            "LA*G (Area, \u03c3/\u03b3, \u03b3)": "Asymmetrical lorentzian convoluted with a gaussian peak of a certain width "
                                           "similar to the model used in casa XPS"

        }
        return descriptions.get(model, "No description available")


    def create_info_button(self, parent, tooltip):
        info_button = wx.Button(parent, label="?", size=(20, 20))
        info_button.SetToolTip(tooltip)
        # info_button.Bind(wx.EVT_ENTER_WINDOW, self.on_info_hover)
        info_button.Bind(wx.EVT_LEAVE_WINDOW, self.on_info_leave)
        return info_button

    def on_info_hover(self, event):
        button = event.GetEventObject()
        pos = button.ClientToScreen(button.GetPosition())
        self.show_info_popup(pos, button.GetToolTip().GetTip())

    def on_info_leave(self, event):
        if hasattr(self, 'info_popup'):
            self.info_popup.Destroy()

    def show_info_popup(self, pos, text):
        self.info_popup = wx.PopupWindow(self, wx.SIMPLE_BORDER)
        panel = wx.Panel(self.info_popup)
        st = wx.StaticText(panel, label=text, pos=(10, 10))
        sz = st.GetBestSize()
        self.info_popup.SetSize((sz.width + 20, sz.height + 20))
        panel.SetSize(self.info_popup.GetSize())
        self.info_popup.Position(pos, (0, 0))
        self.info_popup.Show()

    def update_background_info_button(self):
        for child in self.background_panel.GetChildren():
            if isinstance(child, wx.Button) and child.GetLabel() == "?":
                child.SetToolTip(self.get_background_description(self.parent.background_method))
                break

    def update_fitting_info_button(self):
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button) and child.GetLabel() == "?":
                child.SetToolTip(self.get_fitting_description(self.parent.selected_fitting_method))
                break

    def on_tab_change(self, event):
        selected_page = event.GetSelection()
        self.parent.background_tab_selected = (selected_page == 0)
        self.parent.peak_fitting_tab_selected = (selected_page == 1)
        self.parent.show_hide_vlines()

        if selected_page == 0:
            self.parent.enable_background_interaction()
        else:
            self.parent.disable_background_interaction()

        if not self.parent.peak_fitting_tab_selected:
            self.parent.deselect_all_peaks()

        event.Skip()

    def on_close(self, event):
        self.parent.background_tab_selected = False
        self.parent.peak_fitting_tab_selected = False
        self.parent.show_hide_vlines()
        self.parent.deselect_all_peaks()
        self.Destroy()

    def parse_cross_section(self, text):
        try:
            values = [float(x.strip()) for x in text.split(',')]
            if len(values) == 4:
                sheet_name = self.parent.sheet_combobox.GetValue()
                self.parent.Data['Core levels'][sheet_name]['Background'].update({
                    'Tougaard_B': values[0],
                    'Tougaard_C': values[1],
                    'Tougaard_D': values[2],
                    'Tougaard_T0': values[3]
                })
                return values
            raise ValueError
        except ValueError:
            return [2866, 1643, 1, 0]

    def on_cross_section_change(self, event):
        self.parse_cross_section(self.cross_section.GetValue())
        # self.parent.plot_manager.plot_background(self.parent)

    def on_cross_section2_change(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        values = self.cross_section2.GetValue().split(',')
        try:
            self.parent.Data['Core levels'][sheet_name]['Background'].update({
                'Tougaard_B2': float(values[0]),
                'Tougaard_C2': float(values[1]),
                'Tougaard_D2': float(values[2]),
                'Tougaard_T02': float(values[3])
            })
        except (ValueError, IndexError):
            pass

    def on_cross_section3_change(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        values = self.cross_section3.GetValue().split(',')
        try:
            self.parent.Data['Core levels'][sheet_name]['Background'].update({
                'Tougaard_B3': float(values[0]),
                'Tougaard_C3': float(values[1]),
                'Tougaard_D3': float(values[2]),
                'Tougaard_T03': float(values[3])
            })
        except (ValueError, IndexError):
            pass

    def on_clear_between_vlines(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            core_level_data = self.parent.Data['Core levels'][sheet_name]
            if 'Background' in core_level_data:
                bg_low = core_level_data['Background']['Bkg Low']
                bg_high = core_level_data['Background']['Bkg High']
                if bg_low is not None and bg_high is not None:
                    x_values = np.array(self.parent.x_values)
                    raw_data = np.array(core_level_data['Raw Data'])
                    background = np.array(core_level_data['Background']['Bkg Y'])

                    mask = (x_values >= min(bg_low, bg_high)) & (x_values <= max(bg_low, bg_high))
                    background[mask] = raw_data[mask]

                    core_level_data['Background']['Bkg Y'] = background.tolist()
                    self.parent.background = background
                    self.parent.clear_and_replot()


class TougaardFitWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="Tougaard Fit", size=(1050, 600))

        self.parent = parent
        self.panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        main_pos = parent.parent.GetPosition()
        main_size = parent.parent.GetSize()
        tougaard_size = self.GetSize()

        x = main_pos.x + (main_size.width - tougaard_size.width) // 2
        y = main_pos.y + (main_size.height - tougaard_size.height) // 2

        self.SetPosition((x, y))

        # Get data from parent window
        sheet_name = self.parent.parent.sheet_combobox.GetValue()
        self.x_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['B.E.'])
        self.y_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['Raw Data'])

        min_x = min(self.x_values)
        max_x = max(self.x_values)

        # Left panel for controls
        control_panel = wx.Panel(self.panel)
        control_sizer = wx.BoxSizer(wx.VERTICAL)

        # Number of Tougaard backgrounds control
        num_box = wx.StaticBox(control_panel, label="Number of Tougaard Backgrounds")
        num_sizer = wx.StaticBoxSizer(num_box, wx.HORIZONTAL)
        self.num_tougaard = wx.SpinCtrl(control_panel, min=1, max=10, initial=1)
        num_sizer.Add(self.num_tougaard, 1, wx.ALL, 5)

        # Range controls
        range_box = wx.StaticBox(control_panel, label="Fit Range")
        range_sizer = wx.StaticBoxSizer(range_box, wx.VERTICAL)
        range_grid = wx.GridSizer(1, 4, 5, 5)

        self.min_range = wx.SpinCtrlDouble(control_panel, min=0, max=2000, inc=0.1, value=str(max_x - 15))
        self.max_range = wx.SpinCtrlDouble(control_panel, min=0, max=2000, inc=0.1, value=str(max_x - 1))

        range_grid.Add(wx.StaticText(control_panel, label="Min:"), 0)
        range_grid.Add(self.min_range, 0)
        range_grid.Add(wx.StaticText(control_panel, label="Max:"), 0)
        range_grid.Add(self.max_range, 0)
        range_sizer.Add(range_grid, 0, wx.ALL | wx.EXPAND, 5)

        # Background start control
        bg_box = wx.StaticBox(control_panel, label="Background Start")
        bg_sizer = wx.StaticBoxSizer(bg_box, wx.HORIZONTAL)
        self.bg_start = wx.SpinCtrlDouble(control_panel, min=0, max=2000, inc=0.1, value=str(min_x + 1))
        bg_sizer.Add(self.bg_start, 1, wx.ALL, 5)

        # Scrolled window for Tougaard parameters
        self.param_scroll = wx.ScrolledWindow(control_panel)
        self.param_scroll.SetScrollRate(0, 20)
        self.param_sizer = wx.BoxSizer(wx.VERTICAL)

        # Initialize empty list for parameter controls
        self.tougaard_params = []

        # Create initial parameter set
        self.create_tougaard_params(1)

        # Fit button
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.fit_button = wx.Button(control_panel, label="Fit")
        self.fit_button.SetMinSize((110, 40))
        self.fit_button.Bind(wx.EVT_BUTTON, self.on_fit)
        self.copy_button = wx.Button(control_panel, label="Copy to Peak Fitting Screen")
        self.copy_button.SetMinSize((110, 40))
        self.copy_button.Bind(wx.EVT_BUTTON, self.on_copy_values)
        button_sizer.Add(self.fit_button, 1, wx.ALL, 5)
        button_sizer.Add(self.copy_button, 1, wx.ALL, 5)


        # Initialize vertical lines as None
        self.vline_min = None
        self.vline_max = None

        self.y_max = max(self.y_values)
        self.y_min = 0.99*min(self.y_values)

        # Plot setup
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.panel, -1, self.figure)
        self.ax = self.figure.add_subplot(111)

        # Initial plot setup
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')
        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')
        self.ax.legend()
        self.ax.set_xlim(max(self.x_values), min(self.x_values))

        # Add vertical lines for fit range
        min_e = self.min_range.GetValue()
        max_e = self.max_range.GetValue()
        self.vline_min = self.ax.axvline(x=min_e, color='red', linestyle='--', alpha=0.5)
        self.vline_max = self.ax.axvline(x=max_e, color='red', linestyle='--', alpha=0.5)

        # Layout
        control_sizer.Add(num_sizer, 0, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(range_sizer, 0, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(bg_sizer, 0, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(self.param_scroll, 1, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(button_sizer, 0, wx.EXPAND|wx.ALL, 5)
        control_panel.SetSizer(control_sizer)

        main_sizer.Add(control_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

        self.panel.SetSizer(main_sizer)

        # Bind events
        self.min_range.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        self.max_range.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        self.num_tougaard.Bind(wx.EVT_SPINCTRL, self.on_num_tougaard_change)
        self.panel.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)

    def on_key_press(self, event):
        if event.ControlDown():
            if event.GetKeyCode() in [wx.WXK_UP, wx.WXK_DOWN]:
                # y_min =
                intensity_factor = 0.05

                if event.GetKeyCode() == wx.WXK_DOWN:
                    self.y_max = max(self.y_max - intensity_factor * max(self.y_values), self.y_min)
                else:
                    self.y_max += intensity_factor * max(self.y_values)

                self.ax.set_ylim(self.y_min, self.y_max)
                self.canvas.draw_idle()
                return

        event.Skip()

    def create_tougaard_params(self, num_tougaard):
        # Store current values if they exist
        old_values = []
        old_fixed = []
        for params in self.tougaard_params:
            old_values.append({
                'B': params['B']['value'].GetValue(),
                'C': params['C']['value'].GetValue(),
                'D': params['D']['value'].GetValue()
            })
            old_fixed.append({
                'B': params['B']['fixed'].GetValue(),
                'C': params['C']['fixed'].GetValue(),
                'D': params['D']['fixed'].GetValue()
            })

        self.param_sizer.Clear(True)
        self.tougaard_params = []

        for i in range(num_tougaard):
            param_box = wx.StaticBox(self.param_scroll, label=f"Tougaard {i + 1} B [Scaling], C [Position], "
                                                              f"D [Width], T0 is 0")
            box_sizer = wx.StaticBoxSizer(param_box, wx.VERTICAL)

            params = {}
            param_grid = wx.GridBagSizer(5, 5)

            for row, param in enumerate(['B', 'C', 'D']):
                if i < len(old_values):
                    value = old_values[i][param]
                    is_fixed = old_fixed[i][param]
                else:
                    value = 2866 if param == 'B' else 1643 if param == 'C' else 1
                    is_fixed = False

                params[param] = {
                    'value': wx.SpinCtrlDouble(param_box, min=0, max=20000, inc=0.1, value=str(value)),
                    'min': wx.SpinCtrlDouble(param_box, min=0, max=6000, inc=0.1),
                    'max': wx.SpinCtrlDouble(param_box, min=0, max=20000, inc=0.1, value='6000'),
                    'fixed': wx.CheckBox(param_box, label="Fix")
                }
                params[param]['fixed'].SetValue(is_fixed)

                param_grid.Add(wx.StaticText(param_box, label=f"{param}:"), pos=(row, 0), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['value'], pos=(row, 1), flag=wx.EXPAND)
                param_grid.Add(params[param]['fixed'], pos=(row, 2), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(wx.StaticText(param_box, label="Min:"), pos=(row, 3), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['min'], pos=(row, 4), flag=wx.EXPAND)
                param_grid.Add(wx.StaticText(param_box, label="Max:"), pos=(row, 5), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['max'], pos=(row, 6), flag=wx.EXPAND)

            box_sizer.Add(param_grid, 0, wx.EXPAND | wx.ALL, 5)
            self.tougaard_params.append(params)
            self.param_sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)

        self.param_scroll.SetSizer(self.param_sizer)
        self.param_scroll.Layout()



    def on_num_tougaard_change(self, event):
        num = self.num_tougaard.GetValue()
        self.create_tougaard_params(num)
        self.param_scroll.FitInside()

    def on_fit(self, event):
        min_e = self.min_range.GetValue()
        max_e = self.max_range.GetValue()
        bg_start = self.bg_start.GetValue()

        if self.x_values is None or self.y_values is None:
            print("Error: No data available for fitting")
            return

        params = lmfit.Parameters()
        for i, tougaard in enumerate(self.tougaard_params):
            for param_name in ['B', 'C', 'D']:
                name = f'{param_name}{i + 1}'
                value = float(tougaard[param_name]['value'].GetValue())
                is_fixed = tougaard[param_name]['fixed'].GetValue()

                if is_fixed:
                    params.add(name, value=value, vary=False)
                else:
                    min_val = float(tougaard[param_name]['min'].GetValue())
                    max_val = float(tougaard[param_name]['max'].GetValue())
                    params.add(name, value=value, min=min_val, max=max_val, vary=True)

        bg_mask = self.x_values >= bg_start
        x_full = self.x_values[bg_mask]
        y_full = self.y_values[bg_mask]

        if len(x_full) == 0 or len(y_full) == 0:
            print("Error: No data points in selected range")
            return

        fit_mask = (x_full >= min_e) & (x_full <= max_e)
        x_fit = x_full[fit_mask]
        y_fit = y_full[fit_mask]

        if len(x_fit) == 0 or len(y_fit) == 0:
            print("Error: No data points in fitting range")
            return

        def tougaard_model(params, x_full, y_full, x_fit, y_fit):
            baseline = y_full[-1]
            y_shifted = y_full - baseline
            dx = np.mean(np.diff(x_full))
            background_full = np.zeros_like(y_full)

            for i in range(len(x_full)):
                E = x_full[i:] - x_full[i]
                K_total = 0
                for j in range(len(self.tougaard_params)):
                    B = params[f'B{j + 1}'].value
                    C = params[f'C{j + 1}'].value
                    D = params[f'D{j + 1}'].value
                    K_total += B * E / ((C - E ** 2) ** 2 + D * E ** 2)
                background_full[i] = np.trapz(K_total * y_shifted[i:], dx=dx)

            background_full += baseline
            fit_indices = np.where(fit_mask)[0]
            background_fit = background_full[fit_indices]
            return y_fit - background_fit

        result = lmfit.minimize(tougaard_model,
                                params,
                                args=(x_full, y_full, x_fit, y_fit),
                                method='least_squares',
                                ftol=1e-10,
                                xtol=1e-10,
                                max_nfev=100,
                                scale_covar=True,
                                verbose=True)

        print("\nFit Results:")
        print(result.message)
        print(f"Success: {result.success}")
        print(f"Number of function evaluations: {result.nfev}")
        print("\nFitted Parameters:")
        for i in range(len(self.tougaard_params)):
            for param in ['B', 'C', 'D']:
                name = f"{param}{i + 1}"
                value = result.params[name].value
                stderr = result.params[name].stderr if result.params[name].stderr is not None else 0
                print(f"{name}: {value:.2f} ± {stderr:.2f}")

        if result.success:
            self.plot_results(x_fit, y_fit, result.params, result)
        else:
            self.plot_results(x_fit, y_fit, result.params, result)
            print("Warning: Fit did not converge")

    def plot_results(self, x, y, fitted_params, result):
        self.ax.clear()
        bg_start = self.bg_start.GetValue()

        # Update control values with fitted parameters
        for i, tougaard in enumerate(self.tougaard_params):
            tougaard['B']['value'].SetValue(fitted_params[f'B{i + 1}'].value)
            tougaard['C']['value'].SetValue(fitted_params[f'C{i + 1}'].value)
            tougaard['D']['value'].SetValue(fitted_params[f'D{i + 1}'].value)

        # Calculate each background separately
        mask = self.x_values >= bg_start
        x_bg = self.x_values[mask]
        y_bg = self.y_values[mask]
        baseline = y_bg[-1]

        total_background = np.zeros_like(y_bg) + baseline

        # Plot data
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')

        # Plot each Tougaard background
        colors = plt.cm.tab10(np.linspace(0, 1, len(self.tougaard_params)))
        for j, color in enumerate(colors):
            background = np.zeros_like(y_bg) + baseline
            y_shifted = y_bg - baseline
            dx = np.mean(np.diff(x_bg))

            for i in range(len(x_bg)):
                E = x_bg[i:] - x_bg[i]
                B = fitted_params[f'B{j + 1}'].value
                C = fitted_params[f'C{j + 1}'].value
                D = fitted_params[f'D{j + 1}'].value
                K = B * E / ((C - E ** 2) ** 2 + D * E ** 2)
                background[i] += np.trapz(K * y_shifted[i:], dx=dx)

            # background += baseline
            total_background += (background - baseline)
            self.ax.plot(x_bg, background, '--', color=color, alpha=0.5,
                         label=f'Tougaard {j + 1}')

        # Plot total background
        self.ax.plot(x_bg, total_background, 'b-', label='Total Background')

        # Add vertical lines
        min_e = self.min_range.GetValue()
        max_e = self.max_range.GetValue()
        self.vline_min = self.ax.axvline(x=min_e, color='red', linestyle='--', alpha=0.5)
        self.vline_max = self.ax.axvline(x=max_e, color='red', linestyle='--', alpha=0.5)

        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')
        self.ax.legend()
        self.ax.set_xlim(max(self.x_values), min(self.x_values))
        self.ax.set_ylim(self.y_min, self.y_max)  # Use stored y_max
        self.canvas.draw()

    def plot_initial_data(self):
        self.ax.clear()
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')

        # Add vertical lines for fit range
        min_e = self.min_range.GetValue()
        max_e = self.max_range.GetValue()

        self.vline_min = self.ax.axvline(x=min_e, color='red', linestyle='--', alpha=0.5)
        self.vline_max = self.ax.axvline(x=max_e, color='red', linestyle='--', alpha=0.5)

        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')
        self.ax.legend()
        self.ax.set_xlim(max(self.x_values), min(self.x_values))

        # Add vertical lines for fit range
        min_e = self.min_range.GetValue()
        max_e = self.max_range.GetValue()
        self.vline_min = self.ax.axvline(x=min_e, color='red', linestyle='--', alpha=0.5)
        self.vline_max = self.ax.axvline(x=max_e, color='red', linestyle='--', alpha=0.5)

        self.canvas.draw()

    def on_range_change(self, event):
        if self.vline_min is not None:
            self.vline_min.set_xdata([self.min_range.GetValue(), self.min_range.GetValue()])
        if self.vline_max is not None:
            self.vline_max.set_xdata([self.max_range.GetValue(), self.max_range.GetValue()])
        self.canvas.draw()

    def on_copy_values(self, event):
        for i, tougaard in enumerate(self.tougaard_params):
            B = tougaard['B']['value'].GetValue()
            C = tougaard['C']['value'].GetValue()
            D = tougaard['D']['value'].GetValue()

            if i == 0:
                self.parent.cross_section.SetValue(f"{B},{C},{D},0")
            elif i == 1:
                self.parent.cross_section2.SetValue(f"{B},{C},{D},0")
            elif i == 2:
                self.parent.cross_section3.SetValue(f"{B},{C},{D},0")


class CustomComboBox2(wx.ComboBox):
    def __init__(self, parent, style=wx.CB_READONLY):
        style |= wx.CB_OWNERDRAW
        super().__init__(parent, style=style)
        self.green_items = []
        self.Bind(wx.EVT_DRAW_ITEM, self.OnDrawItem)

    def SetGreenItems(self, items):
        self.green_items = items
        self.Refresh()

    def OnDrawItem(self, evt):
        dc = wx.PaintDC(self)
        dc.SetFont(self.GetFont())
        text = self.GetString(evt.GetIndex())
        rect = evt.GetRect()

        if evt.IsSelected():
            bg_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_HIGHLIGHT)
            text_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_HIGHLIGHTTEXT)
        else:
            bg_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW)
            text_color = wx.GREEN if text in self.green_items else wx.BLACK

        dc.SetBrush(wx.Brush(bg_color))
        dc.SetTextForeground(text_color)
        dc.DrawRectangle(rect)
        dc.DrawText(text, rect.x + 3, rect.y + 2)

class CustomComboBox(wx.ComboBox):
    def __init__(self, parent, style=wx.CB_READONLY):
        super().__init__(parent, style=style)
        self.green_items = []

    def SetGreenItems(self, items):
        self.green_items = items
        for i, item in enumerate(self.GetItems()):
            self.SetForegroundColour(wx.GREEN if item in self.green_items else wx.BLACK)
            self.Refresh()

    def GetGreenItems(self):
        return self.green_items

    def OnMeasureItem(self, item):
        return 24

    def OnSelect(self, event):
        selection = self.GetStringSelection()
        if selection in self.green_items:
            wx.CallAfter(self.SetSelection, self.GetSelection())
        else:
            event.Skip()