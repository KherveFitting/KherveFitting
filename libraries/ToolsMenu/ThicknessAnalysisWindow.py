# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

"""
Thickness Analysis Window

Combines two XPS thickness/depth analysis methods in a tabbed interface:
1. Tougaard Method - Based on inelastic background analysis
2. Thickogram Method - Based on overlayer/substrate intensity ratios
"""

import wx
import wx.lib.dialogs


class ThicknessAnalysisWindow(wx.Frame):
    """Combined Thickness Analysis Window with Tougaard and Thickogram tabs."""

    def __init__(self, parent):
        # if 'wxMac' in wx.PlatformInfo:
        #     window_size = (1050, 720)
        # else:
        #     window_size = (1100, 720)
        if 'wxMac' in wx.PlatformInfo:
            window_size = (1000, 800)  # Smaller for macOS
        elif 'wxGTK' in wx.PlatformInfo:  # Linux
            window_size = (1020, 800)
        else:  # Windows
            window_size = (1020, 800)

        super().__init__(parent, title="XPS Thickness Analysis",
                         size=window_size, style=wx.DEFAULT_FRAME_STYLE)

        self.parent = parent
        self.InitUI()
        self.Centre()

        try:
            from libraries.ConfigFile import set_consistent_fonts
            set_consistent_fonts(self)
        except ImportError:
            pass

    def InitUI(self):
        # Create menu bar
        menubar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()
        export_png_item = file_menu.Append(wx.ID_ANY, "Export as PNG\tCtrl+P", "Export plot as PNG image")
        export_svg_item = file_menu.Append(wx.ID_ANY, "Export as SVG\tCtrl+S", "Export plot as SVG image")
        menubar.Append(file_menu, "&File")

        # Help menu
        help_menu = wx.Menu()
        about_tougaard_item = help_menu.Append(wx.ID_ANY, "About Tougaard Method", "About Tougaard depth analysis")
        about_thickogram_item = help_menu.Append(wx.ID_ANY, "About Thickogram Method", "About Thickogram calculator")
        menubar.Append(help_menu, "&Help")

        self.SetMenuBar(menubar)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_export_png, export_png_item)
        self.Bind(wx.EVT_MENU, self.on_export_svg, export_svg_item)
        self.Bind(wx.EVT_MENU, self.on_about_tougaard, about_tougaard_item)
        self.Bind(wx.EVT_MENU, self.on_about_thickogram, about_thickogram_item)

        # Create notebook with tabs
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        self.notebook = wx.Notebook(panel)

        # Import the existing classes
        from libraries.ToolsMenu.TougaardAnalysisWindow import TougaardAnalysisWindow
        from libraries.ToolsMenu.ThickogramWindow import ThickogramWindow

        # Tab 1: Tougaard Analysis - embed as panel
        self.tougaard_container = wx.Panel(self.notebook)
        tougaard_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create TougaardAnalysisWindow but reparent its content
        self.tougaard_window = TougaardAnalysisWindow(self.parent)
        self.tougaard_window.Hide()

        # Get the main panel from TougaardAnalysisWindow and reparent it
        for child in self.tougaard_window.GetChildren():
            child.Reparent(self.tougaard_container)
            tougaard_sizer.Add(child, 0, wx.EXPAND,0)
            break

        self.tougaard_container.SetSizer(tougaard_sizer)
        self.notebook.AddPage(self.tougaard_container, "Tougaard Depth Analysis")

        # Store reference to figure for export
        self.tougaard_figure = self.tougaard_window.figure

        # Tab 2: Thickogram - embed as panel
        self.thickogram_container = wx.Panel(self.notebook)
        thickogram_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create ThickogramWindow but reparent its content
        self.thickogram_window = ThickogramWindow(self.parent)
        self.thickogram_window.Hide()

        # Get the main panel from ThickogramWindow and reparent it
        for child in self.thickogram_window.GetChildren():
            child.Reparent(self.thickogram_container)
            thickogram_sizer.Add(child, 1, wx.EXPAND,0)
            break

        self.thickogram_container.SetSizer(thickogram_sizer)
        self.notebook.AddPage(self.thickogram_container, "Thickogram")

        # Store reference to figure for export
        self.thickogram_figure = self.thickogram_window.figure

        main_sizer.Add(self.notebook, 0, wx.EXPAND | wx.ALL, 0)
        panel.SetSizer(main_sizer)

    def on_export_png(self, event):
        """Export current tab's plot as PNG"""
        current_tab = self.notebook.GetSelection()
        if current_tab == 0:
            figure = self.tougaard_figure
        else:
            figure = self.thickogram_figure

        with wx.FileDialog(self, "Save PNG file", wildcard="PNG files (*.png)|*.png",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            pathname = fileDialog.GetPath()
            try:
                figure.savefig(pathname, format='png', dpi=300, bbox_inches='tight',
                               facecolor='white', edgecolor='none')
                wx.MessageBox(f"Plot saved as PNG: {pathname}", "Export Successful",
                              wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"Error saving PNG: {str(e)}", "Export Error", wx.OK | wx.ICON_ERROR)

    def on_export_svg(self, event):
        """Export current tab's plot as SVG"""
        current_tab = self.notebook.GetSelection()
        if current_tab == 0:
            figure = self.tougaard_figure
        else:
            figure = self.thickogram_figure

        with wx.FileDialog(self, "Save SVG file", wildcard="SVG files (*.svg)|*.svg",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            pathname = fileDialog.GetPath()
            try:
                figure.savefig(pathname, format='svg', bbox_inches='tight',
                               facecolor='white', edgecolor='none')
                wx.MessageBox(f"Plot saved as SVG: {pathname}", "Export Successful",
                              wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"Error saving SVG: {str(e)}", "Export Error", wx.OK | wx.ICON_ERROR)

    def on_about_tougaard(self, event):
        """Show about dialog for Tougaard method"""
        about_text = """TOUGAARD QUANTITATIVE XPS DEPTH ANALYSIS
═══════════════════════════════════════════════════════════════

THEORETICAL FOUNDATION:
This method analyzes the inelastic background shape to determine 
the depth distribution of atoms. Developed by Prof. Sven Tougaard 
at the University of Southern Denmark.

KEY PARAMETERS:
• B (fitted): Tougaard cross-section parameter from fitting
• B₀ = 2866 eV²: Reference value for homogeneous distribution
• C = 1643 eV²: Universal cross-section parameter (most materials)
• λ (IMFP): Inelastic Mean Free Path from TPP-2M formula

DECAY LENGTH FORMULA:
┌─────────────────────────────────────┐
│     L = B₁/(B₀ - B₁) × λ × cos(θ)  │
└─────────────────────────────────────┘

Where:
  B₁ = fitted B parameter
  B₀ = 2866 eV² (homogeneous reference)
  λ  = IMFP (Inelastic Mean Free Path)
  θ  = emission angle from surface normal

INTERPRETATION OF L:
• L > 0: Surface enriched (atoms decay into bulk)
• L < 0: Subsurface/buried (atoms increase with depth)
• |L| > 6λ: Nearly uniform distribution

Ap/B RATIO METHOD:
• Ap/B ≈ 23 eV: Homogeneous distribution
• Ap/B > 30 eV: Surface localized
• Ap/B < 20 eV: Subsurface/buried

PRIMARY REFERENCE:
S. Tougaard, "Practical guide to the use of backgrounds in 
quantitative XPS", J. Vac. Sci. Technol. A 39, 011201 (2021)"""

        dlg = wx.lib.dialogs.ScrolledMessageDialog(self, about_text,
                                                    "About Tougaard Method", size=(550, 500))
        dlg.ShowModal()
        dlg.Destroy()

    def on_about_thickogram(self, event):
        """Show about dialog for Thickogram method - same as ThickogramWindow"""
        about_text = """THICKOGRAM XPS THICKNESS CALCULATOR
        
        Updated from DaveXPS v0.1

    ═══════════════════════════════════════════════════════════════

    THEORETICAL FOUNDATION:
    This calculator implements a powerful XPS analysis technique pioneered by Peter J. Cumpson 
    at Britain's National Physical Laboratory. The method enables precise measurement of thin 
    film thickness when overlayers possess distinct chemical compositions from their underlying 
    substrates.

    METHODOLOGY OVERVIEW:
    The approach utilizes intensity ratios between overlayer and substrate photoelectron peaks, 
    combined with thickogram visualization developed by Prof. Jong Wan Lee at Hallym University, 
    Korea. This dual approach provides both numerical precision and intuitive graphical analysis.

    TECHNICAL ADVANTAGES:
    ✓ Surface contaminants like adventitious carbon do not compromise measurements
    ✓ Instrumental response factors automatically cancel during ratio calculations  
    ✓ Mathematical simplicity without sacrificing analytical rigor
    ✓ Consistent performance across varied thickness ranges

    EXPERIMENTAL REQUIREMENTS:
    Analysis Geometry: Photoelectron collection angles from 30° to 90° (optimal ~45°)
    Energy Range: Photoelectron kinetic energies exceeding 500 eV recommended
    Measurement Precision: Typically ±10% uncertainty from attenuation length estimations

    INPUT SPECIFICATIONS:
    Peak Intensities: Io (overlayer), Is (substrate) - areas or heights
    Sensitivity Factors: So (overlayer RSF), Ss (substrate RSF)  
    Photoelectron Energies: Eo, Es (kinetic energies of respective peaks)
    Geometric Parameters: θ (emission angle), λo (electron attenuation length)

    SCIENTIFIC RECOGNITION:
    This work builds upon foundational research by Peter J. Cumpson (NPL) and visualization 
    innovations by Jong Wan Lee (Hallym University).

    PRIMARY REFERENCE:
    P.J. Cumpson, "Angle-resolved XPS and AES: depth-resolution limits and a general 
    comparison of properties of depth-profile reconstruction methods"
    Surface and Interface Analysis, Volume 29, Pages 403-406 (2000)

    Secondary Reference:
    J.W. Lee, "Thickogram: A web-based tool for XPS overlayer thickness calculation"
    New Physics: Sae Mulli, Volume 65, Pages 1165-1174 (2015)"""

        dlg = wx.lib.dialogs.ScrolledMessageDialog(self, about_text,
                                                    "About Thickogram Method", size=(650, 550))
        dlg.ShowModal()
        dlg.Destroy()