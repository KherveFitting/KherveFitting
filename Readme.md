# KherveFitting

## Introduction

KherveFitting is an open-source software implemented in Python, using wxPython for the graphical user interface,
MatplotLib for data visualization, NumPy and lmfit for numerical computations and curve fitting algorithms, Panda 
and openpyxl for manipulating Excel file. 

When using KherveFitting in academic or research contexts, appropriate citation is requested  to acknowledge the 
software's contribution to your work.

To download Windows or Mac installer please  go to https://sourceforge.net/projects/khervefitting/



For python source, make sure to use:  pip install -r requirements.txt

## Keyboard Shortcuts

- **Tab:** Select next peak
- **Q:** Select previous peak
- **Ctrl+Minus (-):** Zoom out
- **Equal (=):** Zoom in
- **Ctrl+Left bracket [:** Select previous core level
- **Ctrl+Right bracket ]:** Select next core level
- **Ctrl+Up:** Increase intensity
- **Ctrl+Down:** Decrease intensity
- **Ctrl+Left:** Move plot to High BE
- **Ctrl+Right:** Move plot to Low BE
- **SHIFT+Left:** Decrease High BE
- **SHIFT+Right:** Increase High BE
- **Ctrl+Z:** Undo up to 30 events
- **Ctrl+Y:** Redo
- **Ctrl+S:** Save. Only works on the grid and not on the figure canvas.
- **Ctrl+P:** Open peak fitting window
- **Ctrl+A:** Open Area window
- **Ctrl+K:** Show Keyboard shortcut

## Opening Files

KherveFitting can open Excel files (.xlsx) and import/convert VAMAS files (.vms) into Excel format. For best results:

- Place raw data (X,Y) in Columns A and B, starting at row 0
- Use the row offset control in the horizontal toolbar if needed
- Save each core level in a separate sheet named after the core level (e.g., Si2p, Al2p, C1s, O1s)

When reopening a saved fitting, KherveFitting also looks for the corresponding .json file containing peak properties.

## Saving Files

KherveFitting offers three saving options:

1. Save corrected binding energy, background, envelope, residuals, and fitted peak data of the active core level to columns D onwards in the corresponding Excel sheet. Peak fitting properties for all core levels are saved in a JSON file.
2. Save the figure of the active core level to the corresponding Excel sheet and as a PNG file. The resolution (DPI) can be adjusted in the preference window.
3. Save all fitted core level data, including figures, to the Excel file. Peak fitting properties for all core levels are saved in a JSON file.

## Controlling Peaks

- Ensure the Peak Fitting tab in the Peak Fitting window is selected to move peaks
- Left-click and drag the cross to move a peak
- Shift + Left-click to adjust peak width
- Use the middle mouse scroll to change sheets or core levels

## Peak Fitting Window

### Background Tab

Four background types available: Linear, Shirley, Smart, Multi-Regions Smart. Drag the red lines on the plot to set the 
background range.

- **Linear Background:** Y = mx + b
- **Shirley Background:** B(E) = k × ∫<sub>E</sub><sup>E<sub>max</sub></sup> I(E') dE'
- **Smart Background:** Using linear and Shirley background depending on decreasing and increasing background
- **Adaptive Shirley**


Use high BE and low BE controls to apply offsets at range boundaries.

### Peak Fitting Tab

Fit single peaks or doublets. Doublet splitting values are stored in 'DS.txt'. Intensity ratios for doublets: 0.5 for 
p-shell, 0.67 for d-shell, 0.75 for f-shell.

Available fitting models:

- GL Gaussian-Lorentzian product (constraint with area or height)
- SGL Gaussian-Lorentzian sum (constraint with area or height)
- Pseudo-Voigt from Lmfit library (constraint with area only)
- Voigt from lmfit library (constraint with area only)
- Asymmetric Exponential Gaussian line shape (Under Test)
- Asymmetric Lorentzian LA  with various constraint(Under Test)


## Peak Fitting Parameter Grid

Each peak uses two rows: values in the first row, constraints in the second.

Constraint shortcuts:
- 'a', 'b', 'c' → 'A * 1', 'B * 1', 'C * 1' (follow peak A, B, or C)
- 'fi' → 'Fixed' (fix the value)
- '#0.5' → Constrain to ±0.5 eV of the peak position

## Binding Energy Correction

The BE correction button looks for a peak labeled 'C1s C-C' and calculates the difference from 284.8 eV. This 
correction is applied to all core levels. Fit all data before applying the BE correction.

## Plot Customization

Use the Preferences window to customize plot appearance, including:
- Colors for raw data, background, fitted peaks, and residuals
- Line styles (solid, dashed, dotted)
- Marker types for data points
- Font sizes and styles
- Axis labels and titles

## Toggling Display Elements

Use toggle buttons to show or hide various plot elements:
- Raw data points
- Background line
- Individual fitted peaks
- Overall envelope
- Residuals
- Legend

## Zooming and Navigation

Several tools are available for zooming and navigating the plot:
- Use the zoom in/out buttons or keyboard shortcuts
- Click and drag to create a zoom box
- Use the pan tool to move around when zoomed in

## Exporting Results

Export fitted peak parameters, areas, and atomic percentages to a summary table for further analysis. The export function provides:
- Peak positions, heights, and widths
- Integrated areas for each peak
- Relative sensitivity factors (RSF) used
- Calculated atomic percentages
- Options to export as Excel, PNG, PDF, SVG

## Noise Analysis

Use the Noise Analysis tool to assess data quality and determine the signal-to-noise ratio of your spectra. Features include:
- Automatic noise level detection
- Signal-to-noise ratio calculation
- Noise histogram display
- Options for different noise reduction methods

## License

KherveFitting is distributed under the BSD-3 License allowing for broad use, modification, and distribu-
tion.

## Contact

Gwilherm Kerherve: g.kerherve@imperial.ac.uk

## Acknowledgements

This work was supported by Imperial College London and the Group of Prof. David J. Payne