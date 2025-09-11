# KherveFitting

## Preface

KherveFitting represents the third major iteration of LG4X for X-ray photoelectron spectroscopy 
(XPS) curve fitting analysis.

- **LG4X-V1**: Originally developed by Hideki Nakajima, providing the foundational graphical user interface for 
  XPS curve fitting based on the Python lmfit package.

- **LG4X-V2**: Substantially enhanced and extended by Julian Andreas Hochhaus (2022-2024), introducing significant 
  improvements to the fitting procedures, user interface, and validation capabilities. LG4X-V2 is actively 
  maintained and available through GitHub and Flathub.

- **KherveFitting or LG4X-V3**: Developed by Gwilherm Kerherve at Imperial College London, representing an advanced 
  iteration that introduces wxPython-based interface, Excel integration, and enhanced fitting algorithms while 
  maintaining the core philosophy of accessible XPS data analysis.

We acknowledge and thank both Hideki Nakajima for creating the original LG4X software and Julian Andreas Hochhaus 
for his substantial contributions to LG4X-V2, which continue to serve the XPS community. KherveFitting builds upon 
this legacy while introducing new capabilities for modern XPS analysis workflows.

## Introduction

KherveFitting is a full-featured XPS & Raman fitting software implemented in Python, using wxPython for the graphical user interface, MatplotLib for data visualization, NumPy and lmfit for numerical computations and curve fitting algorithms, Pandas and openpyxl for manipulating Excel files. The software integrates advanced XPS analysis tools including lmfitxps for specialized XPS fitting models and background calculations.

When using KherveFitting in academic or research contexts, appropriate citation is requested to acknowledge the software's contribution to your work.

## Download and Installation

### Windows and macOS Installers
Download from SourceForge: https://sourceforge.net/projects/khervefitting/

**Installation Notes:**
- **macOS**: Launch from Applications folder (right-click and select "Open" first time due to security restrictions)
- **Windows**: Choose file location but DO NOT install in Program Files. Do not run as administrator.

### Python Source
For source installation: `pip install -r requirements.txt`

## File Support

KherveFitting can open and convert multiple file formats:
- **Excel files** (.xlsx)
- **VAMAS files** (.vms) - automatic conversion to Excel
- **Thermo files** (.avg)
- **Kratos files** (.kal)
- **Phi files** (.spe)
- **Specs, Scienta, Omicron** formats

## Keyboard Shortcuts

- **Tab:** Select next peak
- **Q:** Select previous peak
- **Ctrl+Minus (-):** Zoom out
- **Ctrl+Equal (=):** Zoom in
- **Ctrl+Left bracket [:** Select previous core level
- **Ctrl+Right bracket ]:** Select next core level
- **Ctrl+Up:** Increase plot intensity
- **Ctrl+Down:** Decrease plot intensity
- **Ctrl+Left:** Move plot to High BE
- **Ctrl+Right:** Move plot to Low BE
- **SHIFT+Left:** Decrease High BE
- **SHIFT+Right:** Increase High BE
- **Ctrl+Z:** Undo up to 30 events
- **Ctrl+Y:** Redo
- **Ctrl+S:** Save (only works on grid, not figure canvas)
- **Ctrl+P:** Open peak fitting window
- **Ctrl+A:** Open Area window
- **Ctrl+K:** Show keyboard shortcuts
- **Alt+Up/Down:** Increase/Decrease peak intensity
- **Alt+Left/Right:** Move peak to High/Low BE

## File Operations

### Opening Files

**Best Practices:**
- Place raw data (X,Y) in Columns A and B, starting at row 0
- Use the row offset control in horizontal toolbar if needed
- Save each core level in separate sheet named after core level (e.g., Si2p, Al2p, C1s, O1s)
- When reopening saved fittings, ensure corresponding .json file is in same directory

### Saving Files

Three saving options available:

1. **Save Active Core Level:** Corrected binding energy, background, envelope, residuals, and fitted peak data saved to columns D onwards in Excel sheet. Peak properties saved in JSON format.

2. **Save Figure:** Current core level figure saved to Excel sheet and as PNG file. Resolution adjustable in preferences.

3. **Save All:** All fitted core level data and figures saved to Excel file with complete JSON metadata.

## Background Analysis

### Background Types Available

- **Linear Background:** Y = mx + b
- **Shirley Background:** B(E) = k × ∫<sub>E</sub><sup>E<sub>max</sub></sup> I(E') dE'
- **Smart Background:** Automatically chooses between linear and Shirley based on spectral features
- **Multi-Regions Smart:** Different background types for different energy ranges
- **Adaptive Smart:** Advanced region-specific background calculation
- **U2-Tougaard:** Two-parameter Tougaard background with auto-calculated B parameter
- **U4-Tougaard:** Four-parameter Tougaard background for complex inelastic backgrounds

**Controls:**
- Drag red vertical lines on plot to set background range
- Use High BE and Low BE offset controls for boundary adjustments
- Default iteration limit: 100 for all iterative methods (Shirley, Smart, Adaptive Smart)

## Peak Fitting

### Available Fitting Models

- **GL:** Gaussian-Lorentzian product (area or height constraints)
- **SGL:** Gaussian-Lorentzian sum (area or height constraints)  
- **Pseudo-Voigt:** From lmfit library (area constraints only)
- **Voigt:** From lmfit library (area constraints only)
- **Asymmetric Exponential Gaussian:** Advanced lineshape (Under Test)
- **Asymmetric Lorentzian LA:** Various constraint options (Under Test)
- **Doniac-Sunjic (G*DS):** For metallic systems

### Peak Parameters

**Parameter Grid (all data in .2f format):**
- Each peak uses two rows: values in first row, constraints in second
- Standard intensity ratios: 0.5 (p-shell), 0.67 (d-shell), 0.75 (f-shell)
- Doublet splitting values stored in library file

**Constraint Shortcuts:**
- `'a'`, `'b'`, `'c'` → `'A * 1'`, `'B * 1'`, `'C * 1'` (follow peak A, B, or C)
- `'fi'` → `'Fixed'` (fix the value)
- `'#0.5'` → Constrain to ±0.5 eV of peak position

### Peak Control

- Ensure Peak Fitting tab selected for peak manipulation
- Left-click and drag cross marker to move peaks
- Shift + Left-click to adjust peak width
- Middle mouse scroll to change sheets or core levels

## Advanced Features

### Survey Scan Analysis

**Auto ID Functionality:**
- Automatic peak identification for Survey and Wide scan sheets
- Element library with binding energy references
- Core Level List window for element identification
- Compatible with various instrument data formats

### Binding Energy Correction

The BE correction function:
- Searches for peak labeled 'C1s C-C'
- Calculates difference from reference value (284.8 eV)
- Applies correction to all core levels
- **Note:** Fit all data before applying BE correction

### Valence Band Measurements

Specialized VB analysis window provides:
- Fermi edge fitting using lmfitxps FermiEdgeModel
- Thermal broadening analysis
- VBM (Valence Band Maximum) determination
- Cut-off energy calculation
- Integration with main fitting workflow

### Quantitative Analysis

**Atomic Concentration Calculation:**
- Multiple intensity calibration methods: TPP-2M, Scofield, Wagner
- Atomic concentration accurately calculated and matching Thermo Avantage software
- RSF (Relative Sensitivity Factor) library included
- Export options for quantitative results

## Plot Customization

**Preferences Window Options:**
- Colors for raw data, background, fitted peaks, residuals
- Line styles (solid, dashed, dotted)
- Marker types and sizes
- Font sizes and styles
- Axis labels and titles
- Figure resolution (DPI) settings

**Display Toggle Controls:**
- Raw data points
- Background line
- Individual fitted peaks
- Overall envelope
- Residuals
- Legend

## Navigation and Analysis Tools

**Zoom and Navigation:**
- Zoom in/out buttons and keyboard shortcuts
- Click and drag zoom box creation
- Pan tool for navigation when zoomed
- Range selection with vertical lines

**Quality Assessment:**
- Real-time fit quality indicators: R², Reduced Chi-squared
- Noise analysis with automatic detection
- Signal-to-noise ratio calculation
- Residuals analysis for fit quality

## Data Export

**Export Options:**
- Peak positions, heights, widths (all in .2f format)
- Integrated areas for each peak
- Relative sensitivity factors (RSF)
- Calculated atomic percentages
- Multiple formats: Excel, PNG, PDF, SVG

## Technical Specifications

**Dependencies:**
- Python with wxPython, matplotlib, NumPy, pandas, openpyxl
- lmfit for curve fitting algorithms
- lmfitxps for specialized XPS models and backgrounds
- scipy for mathematical operations

**Data Format:**
- All numerical grid data formatted to .2f precision
- Excel integration with automatic sheet detection
- JSON metadata for fit parameters and constraints

**Optimization:**
- Multiple optimization methods available
- Convergence control (20-500 iterations)
- Various weighting schemes: uniform, intensity-based, statistical-XPS, hybrid-XPS
- Undo/Redo support (up to 30 operations)

## License

KherveFitting is distributed under the BSD-3 License, allowing for broad use, modification, and distribution.

## Contact

**Gwilherm Kerherve**  
Email: g.kerherve@imperial.ac.uk  
YouTube Channel: https://www.youtube.com/@xpsexamples-imperialcolleg6571

## Acknowledgements

This work was supported by Imperial College London and the Group of Prof. David J. Payne.