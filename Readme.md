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

### Input Formats

KherveFitting supports comprehensive file format compatibility with drag-and-drop functionality:

**Native Formats:**
- **Excel** (.xlsx, .xls) - KherveFitting native, Thermo Avantage export, generic
- **JSON** - Metadata and peak parameters

**Manufacturer Formats:**
- **VAMAS** (.vms) - ISO 14976 standard, automatic conversion to Excel
- **Thermo Scientific** (.avg) - Avantage format
- **Kratos Analytical** (.kal)
- **Physical Electronics (PHI)** (.spe)
- **MRS format** (.mrs) - Single and batch import
- **VG-Microtech** (.1) - Legacy format

**Generic/Universal:**
- **ASCII** (.asc) - XPS data
- **CSV** (.csv) - Comma-separated values
- **Text** (.txt) - Raman and XAS data
- **Data** (.dat) - Diamond B07 XAS format

**Synchrotron Data:**
- **Diamond B07 XAS** (.txt, .dat) - X-ray Absorption Spectroscopy

### Output Formats

- **Excel** (.xlsx) with embedded plots and JSON metadata
- **VAMAS** (.vms) export
- **Images**: PNG, PDF, SVG
- **Data**: TXT, CSV, DAT
- **Reports**: DOCX (Word) with formatted results
- **Python scripts** for plot recreation

## Keyboard Shortcuts

### Navigation
- **Tab:** Select next peak
- **Q:** Select previous peak
- **Ctrl+[ (Left bracket):** Select previous core level
- **Ctrl+] (Right bracket):** Select next core level

### File Operations
- **Ctrl+O:** Open file
- **Ctrl+S:** Quick save (JSON only)
- **Ctrl+N:** New window instance
- **Ctrl+Q:** Exit application

### View and Zoom
- **Ctrl+- (Minus):** Zoom out
- **Ctrl+= (Equal):** Zoom in
- **Ctrl+Up:** Increase plot intensity
- **Ctrl+Down:** Decrease plot intensity
- **Ctrl+Left:** Move plot to High BE
- **Ctrl+Right:** Move plot to Low BE

### Background Adjustment
- **Shift+Left:** Decrease High BE boundary
- **Shift+Right:** Increase High BE boundary

### Peak Manipulation
- **Alt+Up:** Increase peak intensity (5% of current height)
- **Alt+Down:** Decrease peak intensity
- **Alt+Left:** Move peak to High BE (0.1 eV steps)
- **Alt+Right:** Move peak to Low BE (0.1 eV steps)
- **Alt+Shift+Left/Right:** Adjust peak width (0.05 eV for FWHM or σ)

### Operations and Tools
- **Ctrl+Z:** Undo (up to 30 events)
- **Ctrl+Y:** Redo
- **Ctrl+P:** Open peak fitting window
- **Ctrl+A:** Open Area window
- **Ctrl+D:** Open D-parameter calculator
- **Ctrl+K:** Show keyboard shortcuts
- **Ctrl+M:** Open manual
- **Ctrl+I:** Manual Peak ID / Labels
- **Ctrl+B:** Toggle Kinetic Energy view (Beta)

### Mouse Interactions
- **Left-click + Drag:** Move peak (when peak fitting tab selected)
- **Shift + Left-click:** Adjust peak width
- **Middle scroll:** Change sheets or core levels
- **Right-click:** Context menu with copy/paste, export, data editing options
- **Click and drag:** Create zoom box
- **Drag mode tool:** Pan when zoomed

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

**Standard Models:**
- **GL (Gaussian-Lorentzian Product):** Area or Height constraints
- **SGL (Gaussian-Lorentzian Sum):** Area or Height constraints
- **Pseudo-Voigt:** From lmfit library (Area constraints only)

**Voigt Models:**
- **Voigt (Area, L/G, σ):** Classic Voigt profile
- **Voigt (Area, σ, γ):** Alternative parameterization
- **Voigt (Area, L/G, σ, S):** With skew parameter

**Asymmetric Models (Under Test):**
- **Asymmetric Exponential Gaussian:** Advanced lineshape for asymmetric peaks
- **Asymmetric Lorentzian LA:** Various constraint options

**Metallic Systems:**
- **Doniach-Sunjic (DS):** DS (A, σ, γ) for metallic systems
- **DS*G:** Doniach-Sunjic convoluted with Gaussian, DS*G (A, σ, γ, S)

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

### Analysis Tools Suite

KherveFitting includes comprehensive analysis tools accessible via the Tools menu:

**1. Area Under Curve Calculator:**
- Multiple background methods for integration
- Quantification without full peak fitting
- Batch mode for processing multiple spectra

**2. Peak Fitting Window:**
- Normal and Mini modes for workspace efficiency
- Background, Fitting, and Batch Operations tabs
- Access to all fitting models and advanced constraints

**3. Automatic Peak Identification (Auto ID):**
- Automatic element identification for Survey and Wide scans
- Comprehensive element library with XPS/NIST binding energy references
- Core Level List window with priority-based detection (P1, P2, P3 elements)
- Minimum intensity threshold (2% of max)
- Auger peak recognition
- Compatible with all instrument data formats

**4. Valence Band Measurements (VBM):**
- Specialized VB analysis window
- Fermi edge fitting using lmfitxps FermiEdgeModel
- Thermal broadening analysis
- Valence Band Maximum (VBM) determination
- Cut-off energy calculation
- Integration with main fitting workflow

**5. D-Parameter Calculator:**
- Differentiation analysis for peak shape characterization
- Asymmetry quantification

**6. Thickogram Calculator (Beta):**
- Overlayer thickness calculations
- IMFP (Inelastic Mean Free Path) integration

**7. PCA Analysis:**
- Principal Component Analysis for spectral datasets
- Multivariate analysis tools for complex data

**8. Plot Modifications Window:**
- Advanced plot customization beyond preferences
- Draggable text annotations with positioning controls
- Custom labels and markers

**9. Tougaard/Raman Background:**
- Specialized fitting window for advanced backgrounds
- Cross-section parameter optimization

**10. Profile Analysis:**
- Profile Data Creator for depth profiling experiments
- Profile Editor for analyzing depth profile results
- Multi-layer analysis capabilities

**11. Noise Analysis:**
- Automatic noise level detection
- Signal-to-noise ratio calculation
- Data quality assessment

**12. EDX/SEM Analysis (Under Test):**
- Energy Dispersive X-ray analysis window
- Element mapping capabilities
- Separate menubar for EDX-specific functions

**13. XAS Background (Beta):**
- X-ray Absorption Spectroscopy background removal
- Pre-edge and post-edge normalization
- XANES/EXAFS data processing

### Binding Energy Correction

Automatic BE correction function:
- Searches for peak labeled 'C1s C-C' (customizable reference)
- Calculates offset from reference value (default: 284.8 eV)
- Applies correction to all core levels in dataset
- Configurable reference peak name and energy in Preferences
- **Important:** Fit all data before applying BE correction

### Quantitative Analysis

**Atomic Concentration Calculation:**

**RSF Libraries:**
- **TPP-2M (default):** Tanuma-Powell-Penn formula for IMFP calculation
- **Scofield:** Classical photoionization cross-sections
- **Wagner:** Empirical sensitivity factors

**Features:**
- Accurate atomic percentage calculation matching Thermo Avantage software
- Weight percentage calculation using atomic masses
- Element-specific sensitivity factors for full periodic table
- Angular correction support (0-90°, default: 54.7°)
- Transmission function correction (NPL-based)
- Multiple core levels per element support

**Export Options:**
- Results Grid with checkbox selection
- Export to Excel with formatting
- Export to Word reports
- Peak positions, heights, widths, areas, RSF values, atomic %

## User Interface Features

### Main Interface Components

**Layout Options:**
- **Grid Layout:** Side-by-side or Tabbed view for parameter grids
- **Themes:** None, Simple, Simple Dark, Simple Darker, Simple Very Dark, Raised, Sunken
- **Splitter Design:** Matplotlib canvas on left, data grids on right
- **Default Size:** 1440x700 optimized for modern displays

**Toolbars:**
- **Horizontal Toolbar:** File operations, fitting tools, analysis functions, library management
- **Vertical Toolbar:** Zoom controls, plot adjustments (BE/intensity), text tools, labels
- **Toggle Toolbar:** Plot visibility, Peak fill, Y-axis, Legend, Fit results, Residuals
- **Delete Toolbar:** Remove results grid entries (all/first/last)

**Data Management:**
- **Sample Manager:** Metadata and sample information tracking
- **Labels Manager:** Annotation library for consistent labeling
- **File Manager:** Recent files list (up to 20 files)
- **Sheet Navigation:** Combobox for quick core level switching

**Interactive Features:**
- Drag-and-drop file opening
- Right-click context menus throughout interface
- Tooltips on all toolbar buttons
- Green vertical line tool (moveable marker)
- Draggable text annotations on plots

### Workflow Support

**Undo/Redo System:**
- Up to 30-50 operations stored
- Works across peak fitting, background changes, and parameter edits
- Keyboard shortcuts: Ctrl+Z (Undo), Ctrl+Y (Redo)

**Auto-Backup:**
- Configurable automatic backup intervals (default: 30 minutes)
- Enable/disable in Preferences
- Prevents data loss during long fitting sessions

**Copy/Paste:**
- Copy/paste peak parameters between sheets
- Copy/paste entire core levels
- Sheet operations: Copy, Join, Delete, Sort

**Data Editing:**
- Crop functionality for data trimming
- Direct grid editing with validation
- Row offset control for Excel data reading

## Plot Customization

### Preferences Window

**Raw Data Appearance:**
- Style: Scatter or line plots
- Scatter size (1-10), marker types (o, s, ^, v, <, >, D, etc.)
- Line width, alpha (transparency 0-1), color
- Line styles: solid, dashed, dotted, dash-dot

**Background, Envelope, and Residuals:**
- Individual color selection
- Alpha transparency control
- Line style and thickness
- Residuals display: off, on main plot, or separate subplot

**Peak Appearance:**
- 15 customizable peak colors
- Fill types: Solid fill or hatch patterns (/, \, |, -, +, x, o, O, ., *)
- Peak line styles: None, Black, Same color as fill, Grey
- Alpha, thickness, and pattern controls
- Toggle peak fill on/off

**Text and Labels:**
- Plot font family (default: DejaVu Sans)
- Axis title size (default: 12)
- Axis number size (default: 10)
- Legend font size (default: 8)
- Label font size (default: 8)
- Core level text size (default: 15)
- Custom X and Y axis labels
- Minor tick divisions (sublines)

**Export Settings:**
- Excel plots: 5.2" × 5.2" @ 100 DPI
- Survey plots: 10" × 5" @ 100 DPI
- Word reports: 5" × 5" @ 300 DPI
- Generic export: 8" × 6" @ 300 DPI
- Customizable dimensions and resolution for all formats

### Display Toggle Controls

Quick visibility toggles for:
- Raw data points
- Background line
- Individual fitted peaks with optional fill
- Overall envelope
- Residuals (main plot or separate)
- Legend with customizable position
- Y-axis
- Fit results summary

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

### System Requirements

**Platform Support:**
- **Windows:** Windows 7 and later (Windows 10/11 recommended)
- **macOS:** macOS 10.13 (High Sierra) and later
- **Linux:** Most distributions with wxGTK support

**Installation Notes:**
- **Windows:** Install in user directory (NOT Program Files), do not run as administrator
- **macOS:** Launch from Applications folder, right-click and select "Open" on first launch
- **Python Source:** Requires Python 3.8 or later

### Dependencies

**Core Dependencies:**

**GUI and Visualization:**
- wxPython 4.2.2 - Cross-platform GUI framework
- matplotlib 3.8.4 - Interactive plotting and visualization
- Pillow 10.3.0 - Image processing

**Scientific Computing:**
- NumPy 1.26.4 - Numerical array operations
- SciPy 1.13.0 - Scientific algorithms and optimization
- lmfit 1.3.2 - Non-linear least-squares minimization
- lmfitxps - Specialized XPS fitting models and backgrounds
- pandas 2.2.1 - Data manipulation and analysis
- uncertainties 3.2.2 - Error propagation

**File Handling:**
- openpyxl 3.1.2 - Excel file read/write
- xlrd 2.0.1 - Legacy Excel reading
- XlsxWriter 3.2.0 - Excel writing with formatting
- vamas 0.1.1 - VAMAS ISO 14976 format support
- spe2py 2.0.0 - PHI SPE file format
- yadg 6.0 - Multi-format data conversion
- pyarrow / fastparquet - Parquet file support

**Document Generation:**
- python-pptx 0.6.23 - PowerPoint reports
- odfpy 1.4.1 - OpenDocument formats

**Utilities:**
- requests 2.32.3 - HTTP requests for updates
- psutil - System resource monitoring
- pyperclip - Clipboard operations
- beautifulsoup4 - Web content parsing
- tqdm 4.66.4 - Progress bars for batch operations

### Data Format and Precision

**Data Handling:**
- All numerical grid data formatted to .2f precision
- Excel integration with automatic sheet detection
- JSON metadata for fit parameters and constraints
- Automatic backup system with configurable intervals

**File Structure:**
- X, Y data in columns A and B starting at row 0
- Fitted data saved to columns D onwards (15 columns managed by default)
- Each core level in separate sheet (e.g., Si2p, C1s, O1s)
- JSON sidecar files for complete parameter preservation

### Optimization and Performance

**Fitting Algorithms:**
- Multiple optimization methods from scipy/lmfit
- Convergence control (20-500 iterations, default: 100 for background, 50 for peaks)
- Various weighting schemes:
  - Uniform weighting
  - Intensity-based weighting
  - Statistical-XPS weighting
  - Hybrid-XPS weighting

**Performance Features:**
- Real-time fit quality indicators (R², Reduced χ²)
- Efficient data structures for large datasets
- Optimized background calculations
- Batch processing capabilities for multiple spectra

**Undo/Redo System:**
- Up to 30-50 operations stored
- Low memory footprint using differential storage
- Works across all editing operations

## Documentation and Learning Resources

### Built-in Help

**In-Application:**
- **Ctrl+K:** Comprehensive keyboard shortcuts list
- **Ctrl+M:** Full user manual (PDF, v1.5)
- **Tooltips:** Hover over toolbar buttons for quick descriptions
- **Context Menus:** Right-click anywhere for context-specific options

### Online Resources

**YouTube Channel:**
- Video tutorials and examples: https://www.youtube.com/@xpsexamples-imperialcolleg6571
- Step-by-step fitting demonstrations
- Feature highlights and tips

**Reference Papers:**
- Multiplet splitting in XPS
- Coster-Kronig effect explanation
- Peak shape fitting strategies
- D-parameter measurement techniques
- Element-specific fitting guides (C, Fe, Cu, Ti, V, Sc, Zn, Cr, Mn, Co, Ni)

### Sample Data

**30+ Example Datasets Included:**

**Metals:** Ag, Au, Cu, Ni, Pt

**Oxides:** CeOx, CoOx, Cr₂O₃, CuOx, Fe₂O₃, La₂O₃, MnOx, MoO₃, NiOx, RuO₂, TiO₂, V₂O₅, Y₂O₃, ZrO₂

**Complex Materials:** LaSrTiO₃ (STO), LaSrCoMnO₃, PdCaCO₃

**Polymers:** HDPE, PEEK, PET

**Special Datasets:** Raman spectroscopy, Valence Band measurements, Voigt model examples

### Element Database

**KherveDB:**
- Periodic table integration with XPS/NIST binding energies
- LibraryID for automatic element identification
- Binding energy references for all elements (H to U)
- Auger peak database
- Download statistics tracking

### Peak Libraries

**Cloud Integration:**
- Save/load peak libraries to GitHub
- Share fitting parameters with community
- Access pre-configured fitting models

### Community Support

**Bug Reporting:**
- Built-in bug reporting system
- GitHub issue tracker: https://sourceforge.net/projects/khervefitting/
- Direct feedback to development team

**User Registration:**
- Optional registration for update notifications
- Usage analytics (opt-in) to improve software
- Download statistics viewer

## License

KherveFitting is **dual-licensed**:

- **GNU General Public Licence v3.0** ([LICENSE-GPL.txt](LICENSE-GPL.txt)) — free for
  academic, research, and open-source use. Any distributed software that
  incorporates KherveFitting (in whole or in part) must itself be released
  under the GPL.
- **Commercial Licence** ([LICENSE-COMMERCIAL.txt](LICENSE-COMMERCIAL.txt)) — required
  for use in proprietary or closed-source products. Contact
  [gwilherm.kerherve@gmail.com](mailto:gwilherm.kerherve@gmail.com) for terms.

See [LICENSE.md](LICENSE.md) for the full summary and
[LICENCE-CHANGE-NOTICE.md](LICENCE-CHANGE-NOTICE.md) for the transition from the
previous BSD-3-Clause licence. New contributors are asked to agree to the
[Contributor Licence Agreement](CLA.md).

© 2024–2026 Gwilherm Kerherve. All rights reserved.

## Contact

**Gwilherm Kerherve**  
Email: g.kerherve@imperial.ac.uk  
YouTube Channel: https://www.youtube.com/@xpsexamples-imperialcolleg6571

## Acknowledgements

This work was supported by Imperial College London and the Group of Prof. David J. Payne.