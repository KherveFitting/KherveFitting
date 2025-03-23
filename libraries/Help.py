import wx
import wx.html
import os
import sys
import webbrowser
import urllib.parse
import pygame


def show_libraries_used(window):
    text = """-------------------
Libraries Used:
-------------------

wxPython
Copyright (c) 2018-2023 wxPython Team
License: LGPL 2.1

NumPy
Copyright (c) 2005-2023, NumPy Developers 
License: BSD 3-Clause

SciPy
Copyright (c) 2001-2023 SciPy Developers
License: BSD 3-Clause

Matplotlib
Copyright (c) 2012-2023 Matplotlib Development Team
License: BSD Compatible

LMFIT
Copyright (c) 2012-2023 Matthew Newville, The University of Chicago
License: BSD-3

Pandas
Copyright (c) 2008-2023, AQR Capital Management, LLC, Lambda Foundry, Inc. and PyData Development Team
License: BSD 3-Clause

python-docx
Copyright (c) 2013 Steve Canny
License: MIT

openpyxl
Copyright (c) 2010 openpyxl
License: MIT

Pillow
Copyright (c) 2010-2023 Alex Clark and contributors
License: HPND"""

    dlg = wx.MessageDialog(window, text, "Libraries Used", wx.OK | wx.ICON_INFORMATION)
    dlg.ShowModal()
    dlg.Destroy()

def on_about_OLD(self, event):
    about_dialog = wx.Dialog(None, title="About KherveFitting", size=(400, 460))
    panel = wx.Panel(about_dialog)
    sizer = wx.BoxSizer(wx.VERTICAL)

    # Add text information
    name = wx.StaticText(panel, label="KherveFitting")
    name.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
    version = wx.StaticText(panel, label="Version 1.42 Apr 25")
    description = wx.StaticText(panel, label="XPS peak fitting software\ndeveloped by Dr. Gwilherm Kerherve\nat Imperial College London")
    website = wx.adv.HyperlinkCtrl(panel, -1, "Imperial College Profile", "https://www.imperial.ac.uk/people/g.kerherve")
    developers = wx.StaticText(panel, label="Developers:\nDr. Gwilherm Kerherve / g.kerherve@imperial.ac.uk\nWilliam Skinner")
    copyright = wx.StaticText(panel, label="(C) 2024 Gwilherm Kerherve")


    libraries_button = wx.Button(panel, label="Libraries Used")
    libraries_button.Bind(wx.EVT_BUTTON, lambda evt: show_libraries_used(self))

    # Add all elements to sizer
    for item in [name, version, description, website, developers, copyright, libraries_button]:
        sizer.Add(item, 0, wx.ALIGN_CENTER | wx.ALL, 5)

    panel.SetSizer(sizer)

def on_about(self, event):
    about_dialog = wx.Dialog(None, title="About KherveFitting", size=(400, 470))
    panel = wx.Panel(about_dialog)
    sizer = wx.BoxSizer(wx.VERTICAL)

    name = wx.StaticText(panel, label="KherveFitting")
    name.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
    version = wx.StaticText(panel, label="Version 1.42 March 25")

    button_grid = wx.GridBagSizer(2, 2)
    libraries_button = wx.Button(panel, label="Libraries Used")
    libraries_button.Bind(wx.EVT_BUTTON, lambda evt: show_libraries_used(self))
    version_log_button = wx.Button(panel, label="Version Log")
    version_log_button.Bind(wx.EVT_BUTTON, lambda evt: show_version_log(self))

    button_grid.Add(libraries_button, pos=(0, 0), flag=wx.ALL, border=5)
    button_grid.Add(version_log_button, pos=(0, 1), flag=wx.ALL, border=5)

    description = wx.StaticText(panel,
                                label="An open-source XPS peak fitting software\ndeveloped by Dr. Gwilherm Kerherve\nat Imperial College London")
    website = wx.adv.HyperlinkCtrl(panel, -1, "Imperial College Profile",
                                   "https://www.imperial.ac.uk/people/g.kerherve")
    developers = wx.StaticText(panel,
                               label="Developers:\nDr. Gwilherm Kerherve / g.kerherve@imperial.ac.uk\nWilliam Skinner")
    copyright = wx.StaticText(panel, label="(C) 2024 Gwilherm Kerherve")



    for item in [name, version]:
        sizer.Add(item, 0, wx.ALIGN_CENTER | wx.ALL, 5)

    sizer.Add(button_grid, 0, wx.ALIGN_CENTER | wx.ALL, 5)
    panel.SetSizer(sizer)

    for item in [description, website, developers, copyright]:
        sizer.Add(item, 0, wx.ALIGN_CENTER | wx.ALL, 5)

    current_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(current_dir, "Images")
    qr_path = os.path.join(image_path, "buymeacoffee_qr.png")
    if os.path.exists(qr_path):
        qr_image = wx.Image(qr_path, wx.BITMAP_TYPE_PNG)
        qr_image = qr_image.Scale(150, 150, wx.IMAGE_QUALITY_HIGH)
        qr_bitmap = wx.StaticBitmap(panel, -1, wx.Bitmap(qr_image))
        sizer.Add(qr_bitmap, 0, wx.ALIGN_CENTER | wx.ALL, 5)
    about_dialog.ShowModal()
    about_dialog.Destroy()



    # Load and display QR code
    # qr_path = os.path.join("libraries", "Images", "buymeacoffee_qr.png")
    current_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(current_dir, "Images")
    qr_path = os.path.join(image_path, "buymeacoffee_qr.png")
    if os.path.exists(qr_path):
        qr_image = wx.Image(qr_path, wx.BITMAP_TYPE_PNG)
        qr_image = qr_image.Scale(150, 150, wx.IMAGE_QUALITY_HIGH)
        qr_bitmap = wx.StaticBitmap(panel, -1, wx.Bitmap(qr_image))
        sizer.Add(qr_bitmap, 0, wx.ALIGN_CENTER | wx.ALL, 5)
    about_dialog.ShowModal()
    about_dialog.Destroy()

def show_quick_help(parent):
    # if getattr(sys, 'frozen', False):
    #     # If the application is run as a bundle, use the bundle's directory
    #     application_path = sys._MEIPASS
    # else:
    #     # If the application is run as a script, use the script's directory
    #     application_path = os.path.dirname(os.path.abspath(__file__))
    # image_path = os.path.join(application_path, "Images")

    current_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(current_dir, "Images")

    help_text = (
        "<body bgcolor='#FFFFE0'>"
        "<h2><font color='#66CC66'>KherveFitting Help</font></h2>"
        
        "<p>KherveFitting is an open-source software developed in Python, using wxPython for the graphical user "
        "interface, MatplotLib for data visualization, NumPy and lmfit for numerical computations and curve fitting "
        "algorithms, Panda for manipulating Excel files. KherveFitting is distributed under the BSD-3 License, "
        "allowing for broad use, modification, and distribution. When using KherveFitting in academic or research "
        "contexts, appropriate citation would be appreciated to acknowledge the software's contribution.</p>"


        "<h3><font color='#006400'>Opening Files</font></h3>"        
        
        "<p>KherveFitting can open Excel files (.xlsx) and import/convert VAMAS files (.vms), AVG files and Avantage "
        "files  into Excel format.  For best results: "
        "<ul>"
        "<li>Place raw data (X,Y) in Columns A and B, starting at row 0</li>"
        "<li>Use the row offset control in the horizontal toolbar if needed</li>"
        "<li>Save each core level in a separate sheet named after the core level (e.g., Si2p, Al2p, C1s, O1s)</li>"
        "</ul>"
        "</p>"
        "<p>When reopening a saved fitting, KherveFitting also looks for the corresponding .json file containing  all the"
        "peak properties.</p>"

        "<h3><font color='#006400'>Saving Files</font></h3>"
        
        f"<img src='{os.path.join(image_path, 'file_saving.png')}' alt='File Saving'>"
        "<p>KherveFitting offers three saving options:</p>"
        "<ol>"
        "<li>Save corrected binding energy, background, envelope, residuals, and fitted peak data of the active "
        "core level to columns D onwards in the corresponding Excel sheet. The picture of the plot is also saved in "
        "cell D6. Peak fitting properties for all core levels are saved in a JSON file.</li>"
        
        "<li>Save the figure of the active core level to the corresponding Excel sheet and as a PNG file. The "
        "resolution (DPI) is 300 DPI</li>" # can be adjusted in the preference window.</li>"
        "<li>Save all fitted core level data, including figures, to the Excel file. Peak fitting properties for all core levels are saved in a JSON file.</li>"
        "</ol>"

        "<h3><font color='#006400'>Controlling Peaks</font></h3>"
        f"<img src='{os.path.join(image_path, 'peak_control.png')}' alt='Peak Control'>"
        "<ul>"
        "<li>Ensure the Peak Fitting tab in the Peak Fitting window is selected to move peaks</li>"
        "<li>Use TAB or 'Q' to select a peak.</li>"
        "<li>Use the left-click and drag the cross to move the position of a peak</li>"
        "<li>Shift + Left-click and move the mouse to adjust the width of the peak</li>"        
        "</ul>"

        "<h3><font color='#006400'>Peak Fitting Window</font></h3>"
        f"<img src='{os.path.join(image_path, 'peak_fitting_window.png')}' alt='Peak Fitting Window'>"
        "<h4>Background Tab</h4>"
        "<p>Five background types available: Linear, Shirley, Smart, Multi-Regions Smart, and Tougaard. Drag the red lines on the plot to set the background range.</p>"
        "<ul>"
        "<li><b>Linear Background:</b><br>"
        "<pre>Y = mx + b</pre>"
        "where m is the slope and b is the y-intercept</li>"
        "<li><b>Shirley Background:</b><br>"
        "<pre>B(E) = k × ∫<sub>E</sub><sup>E<sub>max</sub></sup> I(E') dE'</pre>"
        "where k is a constant and I(E') is the spectrum intensity</li>"
        "<li><b>Tougaard Background:</b><br>"
        "<pre>B(E) = λ(E) × ∫<sub>E</sub><sup>∞</sup> K(E' - E) × [f(E') - B(E')] dE'</pre>"
        "where λ(E) is the inelastic mean free path and K(E) is the loss function</li>"
        "</ul>"
        "<p>Use high BE and low BE controls to apply offsets at range boundaries.</p>"

        "<h4>Peak Fitting Tab</h4>"
        "<p>Fit single peaks or doublets. Doublet splitting values are stored in 'split.txt'. Intensity ratios for doublets: 0.5 for p-shell, 0.67 for d-shell, 0.75 for f-shell.</p>"
        "<p>Available fitting models:</p>"
        "<ul>"
        "<li><b>GL (Gaussian-Lorentzian product):</b><br>"
        "<pre>I(x) = H × [exp(-ln(2) × ((x-x<sub>0</sub>)/σ)<sup>2</sup>) × (1 / (1 + ((x-x<sub>0</sub>)/γ)<sup>2</sup>))]</pre></li>"
        "<li><b>SGL (Gaussian-Lorentzian sum):</b><br>"
        "<pre>I(x) = H × [m × exp(-ln(2) × ((x-x<sub>0</sub>)/σ)<sup>2</sup>) + (1-m) / (1 + ((x-x<sub>0</sub>)/γ)<sup>2</sup>)]</pre></li>"
        "<li><b>Pseudo-Voigt:</b><br>"
        "<pre>I(x) = H × [η × L(x) + (1-η) × G(x)]</pre>"
        "where L(x) is Lorentzian and G(x) is Gaussian</li>"
        "<li><b>Voigt:</b><br>"
        "<pre>I(x) = H × ∫<sub>-∞</sub><sup>∞</sup> G(x') × L(x-x') dx'</pre>"
        "convolution of Gaussian and Lorentzian</li>"
        "</ul>"
        "<p>Where:<br>"
        "H is peak height<br>"
        "x<sub>0</sub> is peak center<br>"
        "σ is Gaussian width<br>"
        "γ is Lorentzian width<br>"
        "m and η are mixing parameters</p>"
    
        "<h3><font color='#006400'>Peak Fitting Parameter Grid</font></h3>"
        f"<img src='{os.path.join(image_path, 'parameter_grid.png')}' alt='Parameter Grid'>"
        "<p>Each peak uses two rows: values in the first row, constraints in the second.</p>"
        "<p>Constraint shortcuts:</p>"
        "<ul>"
        "<li>'a', 'b', 'c' → 'A*1', 'B*1', 'C*1' (follow peak A, B, or C)</li>"
        "<li>'fi' → 'Fixed' (fix the value)</li>"
        "<li>'#0.5' → Constrain to ±0.5 eV of the peak position</li>"
        "</ul>"

        "<h3><font color='#006400'>Binding Energy Correction</font></h3>"
        f"<img src='{os.path.join(image_path, 'BEcorrections.png')}' alt='BE Correction icon'>"
        "<p>The BE correction button looks for a peak labeled 'C1s C-C' and calculates the difference from 284.8 eV. This correction is applied to all core levels. Fit all data before applying the BE correction.</p>"

        "<h3><font color='#006400'>Plot Customization</font></h3>"
        f"<img src='{os.path.join(image_path, 'plot_customization.png')}' alt='Plot Customization'>"
        "<p>Use the Preferences window to customize plot appearance, including:</p>"
        "<ul>"
        "<li>Colors for raw data, background, fitted peaks, and residuals</li>"
        "<li>Line styles (solid, dashed, dotted)</li>"
        "<li>Marker types for data points</li>"
        "<li>Font sizes and styles</li>"
        "<li>Axis labels and titles</li>"
        "</ul>"

        "<h3><font color='#006400'>Toggling Display Elements</font></h3>"
        f"<img src='{os.path.join(image_path, 'toggle_elements.png')}' alt='Toggle Elements'>"
        "<p>Use toggle buttons to show or hide various plot elements:</p>"
        "<ul>"
        "<li>Raw data points</li>"
        "<li>Background line</li>"
        "<li>Individual fitted peaks</li>"
        "<li>Overall envelope</li>"
        "<li>Residuals</li>"
        "<li>Legend</li>"
        "</ul>"

        "<h3><font color='#006400'>Zooming and Navigation</font></h3>"
        f"<img src='{os.path.join(image_path, 'zoom_navigation.png')}' alt='Zoom and Navigation'>"
        "<p>Several tools are available for zooming and navigating the plot:</p>"
        "<ul>"
        "<li>Use the zoom in/out buttons or keyboard shortcuts</li>"
        "<li>Click and drag to create a zoom box</li>"
        "<li>Double-click to reset the view</li>"
        "<li>Use the pan tool to move around when zoomed in</li>"
        "<li>Adjust x and y axis limits manually in the plot settings</li>"
        "</ul>"

        "<h3><font color='#006400'>Exporting Results</font></h3>"
        f"<img src='{os.path.join(image_path, 'export_results.png')}' alt='Export Results'>"
        "<p>Export fitted peak parameters, areas, and atomic percentages to a summary table for further analysis. The export function provides:"
        "<ul>"
        "<li>Peak positions, heights, and widths</li>"
        "<li>Integrated areas for each peak</li>"
        "<li>Relative sensitivity factors (RSF) used</li>"
        "<li>Calculated atomic percentages</li>"
        "<li>Options to export as CSV, Excel, or copy to clipboard</li>"
        "</ul>"
        "</p>"

        "<h3><font color='#006400'>Noise Analysis</font></h3>"
        f"<img src='{os.path.join(image_path, 'noise_analysis.png')}' alt='Noise Analysis'>"
        "<p>Use the Noise Analysis tool to assess data quality and determine the signal-to-noise ratio of your spectra. Features include:"
        "<ul>"
        "<li>Automatic noise level detection</li>"
        "<li>Signal-to-noise ratio calculation</li>"
        "<li>Noise histogram display</li>"
        "<li>Options for different noise reduction methods</li>"
        "</ul>"
        "</p>"

        "</body>"
    )


    help_dialog = wx.Dialog(parent, title="Quick Help", size=(650, 650),
                            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
    # help_dialog.SetBackgroundColour(wx.Colour(255, 255, 255))  # Light yellow
    html_window = wx.html.HtmlWindow(help_dialog)
    html_window.SetBackgroundColour(wx.Colour(255, 255, 255))  # Light yellow
    html_window.SetPage(help_text)

    help_dialog.Show()
    help_dialog.Bind(wx.EVT_CLOSE, lambda evt: help_dialog.Destroy())

    # '''
    # image_path = os.path.join(application_path, "Images")
    #
    # '''

def show_shortcuts(parent):
    shortcuts_html = """
    <html>
    <head>
    <style>
        body {font-size: 10px; line-height: 1.2;}
        ul {list-style-type: none; padding: 0; margin: 0;}
        li {margin: 2px 0;}
    </style>
    </head>
    <body>
    <h3 style="margin-bottom: 5px;">Keyboard Shortcuts</h3>
    <ul>
    <li><b>Tab:</b> Select next peak</li>
    <li><b>Q:</b> Select previous peak</li>
    <li><b>Ctrl+Minus (-):</b> Zoom out</li>
    <li><b>Ctrl+Equal (=):</b> Zoom in</li>
    <li><b>Ctrl+Left bracket [:</b> Select previous core level</li>
    <li><b>Ctrl+Right bracket ]:</b> Select next core level</li>
    <li><b>Ctrl+Up:</b> Increase plot intensity</li>
    <li><b>Ctrl+Down:</b> Decrease plot intensity</li>
    <li><b>Ctrl+Left:</b> Move plot to High BE</li>
    <li><b>Ctrl+Right:</b> Move plot to Low BE</li>
    <li><b>SHIFT+Left:</b> Decrease High BE</li>
    <li><b>SHIFT+Right:</b> Increase High BE</li>
    <li><b>Ctrl+Z:</b> Undo up to 30 events</li>
    <li><b>Ctrl+Y:</b> Redo</li>
    <li><b>Ctrl+S:</b> Save. Only works on the grid and not on the figure canvas</li>
    <li><b>Ctrl+P:</b> Open peak fitting window</li>
    <li><b>Ctrl+A:</b> Open Area window</li>
    <li><b>Ctrl+K:</b> Show Keyboard shortcut</li>
    <li><b>Alt+Up:</b> Increase peak intensity</li>
    <li><b>Alt+Down:</b> Decrease peak intensity</li>
    <li><b>Alt+Left:</b> Move peak to High BE</li>
    <li><b>Alt+Right:</b> Move peak to Low BE</li>
    <li><b>Alt+SHIFT+Left:</b> increase FWHM</li>
    <li><b>Alt+SHIFT+Right:</b> Decrease FWHM</li>
    <li><b>SHIFT+Mouse Left button:</b> increase/decrease FWHM in Peak Fitting Tab</li>
    <li><b>SHIFT+Mouse Left button:</b> increase/decrease Offset Low or High in Background Tab</li>
    </ul>
    </body>
    </html>
    """

    dlg = wx.Dialog(parent, title="List of Shortcuts", size=(380, 650), style=wx.DEFAULT_DIALOG_STYLE |
                                                                              wx.RESIZE_BORDER)
    html_win = wx.html.HtmlWindow(dlg)
    html_win.SetPage(shortcuts_html)

    # btn = wx.Button(dlg, wx.ID_OK, "Close")
    # btn.Bind(wx.EVT_BUTTON, lambda event: dlg.EndModal(wx.ID_OK))

    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(html_win, 1, wx.EXPAND | wx.ALL, 1)
    # sizer.Add(btn, 0, wx.ALIGN_CENTER | wx.ALL, 5)

    dlg.SetSizer(sizer)
    # dlg.ShowModal()
    dlg.Show()


def show_mini_game(parent):
    from libraries.MiniGame import ParticleSimulation
    pygame.init()
    sim = ParticleSimulation()
    sim.run()
    pygame.quit()

def report_bug_OLD(window):
    """Displays a bug report dialog with options for describing the issue and sending a bug report."""

    dlg = wx.Dialog(window, title="Report Bug", size=(400, 300))
    vbox = wx.BoxSizer(wx.VERTICAL)

    # Widgets
    description_label = wx.StaticText(dlg, label="Bug Description:")
    description = wx.TextCtrl(dlg, style=wx.TE_MULTILINE, size=(380, 100))

    def send_report(evt):
        """Send the bug report via the default email client."""
        subject = "KherveFitting Bug Report"
        body = (f"{description.GetValue()}\n\n"
                "Please attach an image of the bug if applicable.")
        body_encoded = urllib.parse.quote(body)  # URL-encode the body for proper formatting

        if not description.GetValue().strip():
            wx.MessageBox("Please provide a bug description before sending.", "Warning", wx.ICON_WARNING)
            return

        try:
            webbrowser.open(
                f'mailto:g.kerherve@imperial.ac.uk?subject={urllib.parse.quote(subject)}&body={body_encoded}')
            wx.MessageBox("Bug report opened in your email client. Please attach any relevant images before sending.",
                          "Success", wx.ICON_INFORMATION)
            dlg.Close()
        except Exception as e:
            wx.MessageBox(f"Failed to open email client: {str(e)}", "Error", wx.ICON_ERROR)

    # Buttons
    send_btn = wx.Button(dlg, label="Send Report")
    send_btn.Bind(wx.EVT_BUTTON, send_report)

    # Layout
    vbox.Add(description_label, 0, wx.ALL, 5)
    vbox.Add(description, 1, wx.EXPAND | wx.ALL, 5)
    vbox.Add(send_btn, 0, wx.ALIGN_CENTER | wx.ALL, 10)

    dlg.SetSizer(vbox)
    dlg.ShowModal()
    dlg.Destroy()

def report_bug(window):
    """Opens the bug report form in the user's default web browser."""
    dlg = wx.MessageDialog(
        parent=window,
        message=(
            "You will be redirected to the bug report form. "
            "Please provide a description of the issue and attach any relevant files."
        ),
        caption="Report Bug",
        style=wx.OK | wx.CANCEL | wx.ICON_INFORMATION
    )

    if dlg.ShowModal() == wx.ID_OK:
        # Open the web form in the default browser
        webbrowser.open("https://imperial.eu.qualtrics.com/jfe/form/SV_8IZDByjiBQrWdHU")
        # https: // imperial.eu.qualtrics.com / jfe / form / SV_8IZDByjiBQrWdHU
        wx.MessageBox(
            "The bug report form has been opened in your browser.",
            "Success",
            wx.ICON_INFORMATION
        )
    dlg.Destroy()


def show_version_log(window):
    dlg = wx.Dialog(window, title="Version Log", size=(600, 400))
    text = wx.TextCtrl(dlg, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)

    version_log = """Version 1.41-1.5 (March-Apr 2025)
- Release of the Mac Version
- Added Samples/Experiment Manager for easy view of multiple files
- Added Copy and paste menu for core levels and peak fitting grid
- Added A propagation of the peak fitting parameters (menu and =) just like casa
- Added information in the Manual for installation instruction
- Added Sort core levels for multiple samples
- Added backup of the excel file and json file
- Added 2D Sheet Manager similar to CasaXPS
- Change Horizontal toolbar so that it is compatible with Mac
- Preparation to the release of the Linux version for Synchrotron
- Added File > New File in the menu open a new instance of KherveFitting
- Added File > Save as... in the menu
- Added File > Optimise File menu
- Added Save Table in the main Save Sheet and remove it from the toolbar
- Improve the refresh function to exclude the Results Table sheet
- Remove the Noise Analysis from the toolbar. Only available in Tools
- Move the Skewed Voigt in the non default as it was confusing for new users
- Fix bugs that stop Fitting to continue
- Fix bugs with the vamas file from kratos
- Added mini-game back to the help menu


Version 1.4 (February 2025)
- Mac version coming soon
- Added Skewed Voigt model. Default to leastsq fitting method
- Added Automatic download of new KherveFitting version
- Added KratosC1s and KratosF1s to the library
- Created values for all spin orbit coupling for the KratosC1s and KratosF1s libraries
- Added a Toggle menu icon on the left side in line with the MAC version
- Added Settings Icon in the toolbar
- Added Bugs report
- Added Open File Location
- Added Constant Multiplication in Mod window 
- Added Version Log and Libraries Used in the About Menu
- Added Photon source energy in Instrument settings
- Added a View & and edit library in the preference window/instruments.
- Added possibility of changing the reference peak for binding energy correction in Instrument settings
- Added Linewidth settings in the preference window for residual, envelope, background
- Added peak name above the cross when moving a peak
- Added right-click menu to zoom in and out
- Added Delete icons to delete rows in the peak fitting grid
- Added First time window for new users
- Homogeneous font type and size throughout all platforms Mac and Windows
- Reduced Size of Peak fitting window
- Added .xls compatibility for the import of the data 
- Removed the fitting of the data when saving data. it should only be done during fitting
- Improved size of Preference Window (smaller)
- Improve Toggle by adding a third state. Removal of Raw Data, envelope
- Improved Legend filtering 
- Improved Check of the use of constraints in the peak fitting grid
- Improved Measure Area window - more tools added
- Improved Undo / Redo States 
- Improved File checking before opening onto KherveFitting
- Fixed KE plot for all toggles

Version 1.3 (December 2024)
- Added Tougaard background with fitting option
- Added D-parameter calculation for auger spectra
- Export word report improved with automatic table of contents 
- Added support for extracting transmission data from VAMAS files
- Added support for Kratos, PHI SPE and MRS files
- Raw data can now be shown either as scatter plot or lines
- Added plot customisation options for Excel, Word, PNG... in preferences window
- Added peak count and peak area ratio
- Added survey scan customisation
- Added labels option for spectra
- Added EAL IMFP calculation for Thin Film
- Added angular correction factor
- Added axis title and label customization
- Added draggable text annotations
- Added labels manager window
- Added keyboard shortcuts for text size
- Fixed core level name display for doublet peaks

Version 1.0 to 1.2 (November 2024)
- Added LA, LA*G peak models from CasaXPS
- Added new options for Voigt peaks
- Added Gaussian peak with exponential tail
- Added relative sensitivity factors from Scofield, Wagner
- Added inelastic mean free path correction
- Added energy scale toggle between BE and KE
- Added Undo/Redo functionality
- Added peak naming for doublet"""

    text.SetValue(version_log)

    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(text, 1, wx.EXPAND | wx.ALL, 5)

    btn = wx.Button(dlg, wx.ID_OK, "OK")
    sizer.Add(btn, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

    dlg.SetSizer(sizer)
    dlg.ShowModal()
    dlg.Destroy()
