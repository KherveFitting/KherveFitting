import tkinter as tk
from tkinter import ttk, Menu
import pandas as pd
import webbrowser
from tkinter import font as tkfont
import re
from typing import Dict, List, Tuple
import os
import sys
import pyperclip  # For clipboard functionality
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import platform


class PeriodicTableXPS(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KherveDB Library: How I wish NIST would look like")
        self.geometry("620x660")
        # Fix the width but allow height to vary
        self.minsize(620, 660)  # Minimum width set to 740, minimum height can be 0
        self.maxsize(620, 10000)  # Maximum width fixed at 740, height can be very large


        # Set up styles and fonts
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(size=9)
        self.heading_font = tkfont.Font(family="Helvetica", size=10, weight="bold")

        # Configure Mac-specific combobox font
        if self.detect_mac_os():
            self.style.configure("TCombobox", font=("Helvetica", 8))

        # Create main menu
        self.create_menu()

        # Load data
        self.load_data()

        # Create main frames
        self.create_frames()

        # Create widgets
        self.create_periodic_table()
        self.create_search_area()
        self.create_results_table()

        # Initialize variables
        self.selected_element = None
        self.selected_line = None

    def detect_mac_os(self):
        return platform.system() == 'Darwin'

    def create_menu(self):
        """Create the application menu bar"""
        menubar = Menu(self)
        self.config(menu=menubar)

        # File menu
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.destroy)

        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def show_about(self):
        """Display information about the application and database"""
        about_window = tk.Toplevel(self)
        about_window.title("About My KherveNIST Library")
        about_window.geometry("500x300")
        about_window.resizable(False, False)

        # Center the window on screen
        about_window.update_idletasks()
        width = about_window.winfo_width()
        height = about_window.winfo_height()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        # Make the dialog modal
        about_window.transient(self)
        about_window.grab_set()

        # Create a frame with some padding
        frame = tk.Frame(about_window, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Title label
        title_label = tk.Label(frame, text="My KherveNIST Library",
                               font=("Helvetica", 14, "bold"))
        title_label.pack(pady=(0, 10))

        # Information text
        info_text = tk.Text(frame, wrap=tk.WORD, height=10, width=50,
                            font=("Helvetica", 9))
        info_text.pack(fill=tk.BOTH, expand=True)

        # Insert the about text
        about_str = """This application provides access to the NIST X-ray Photoelectron Spectroscopy (XPS) binding energy database recorded in 2019, before it was shut down during the Trump administration.

        The original NIST database is an invaluable resource for scientists and researchers in materials science, 
        chemistry, and physics fields, providing standard reference data for XPS analysis.
    
        This application aims to preserve and provide easy access to this important scientific resource.
    
        Developer: Gwilherm Kerherve
        Version: 1.0
        """
        info_text.insert(tk.END, about_str)
        info_text.config(state=tk.DISABLED)  # Make text read-only

        # Close button
        close_button = tk.Button(frame, text="Close", command=about_window.destroy)
        close_button.pack(pady=(10, 0))

        # Wait for the window to be closed
        self.wait_window(about_window)

    def load_data_OLD(self):
        """Load XPS data from CSV file"""
        # Determine the correct path when running as script or frozen executable
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            # First check if we're in the libraries directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.dirname(current_dir)  # Go up one level to the root

        # Try both locations - root directory and libraries directory
        possible_paths = [
            os.path.join(base_path, "NIST_BE.xlsx"),
            os.path.join(base_path, "libraries", "NIST_BE.xlsx")
        ]

        data_found = False
        for data_path in possible_paths:
            if os.path.exists(data_path):
                try:
                    self.df = pd.read_excel(data_path)
                    # Get unique elements and lines
                    self.elements = sorted(self.df['Element'].unique())
                    self.lines = sorted(self.df['Line'].unique())
                    data_found = True
                    break
                except Exception as e:
                    continue
        print(f'Loaded the .parquet NIST library')
        print(f'Loaded the .xlsx NIST library')
        if not data_found:
            tk.messagebox.showerror("Error", f"Failed to load data: NIST_BE.xlsx not found in expected locations")
            self.destroy()

    def load_data(self):
        """Load XPS data from binary file (fast) or CSV file (fallback)"""
        # Determine the correct path when running as script or frozen executable
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            # First check if we're in the libraries directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.dirname(current_dir)  # Go up one level to the root

        # Try both locations - prioritize parquet files for speed
        possible_paths = [
            os.path.join(base_path, "NIST_BE.parquet"),
            os.path.join(base_path, "libraries", "NIST_BE.parquet"),
            os.path.join(base_path, "NIST_BE.xlsx"),
            os.path.join(base_path, "libraries", "NIST_BE.xlsx")
        ]

        data_found = False
        for data_path in possible_paths:
            if os.path.exists(data_path):
                try:
                    if data_path.endswith('.parquet'):
                        self.df = pd.read_parquet(data_path)
                        print(f'Loaded the .parquet NIST library')
                    else:
                        self.df = pd.read_excel(data_path)
                        print(f'Loaded the .xlsx NIST library')

                    # Get unique elements and lines
                    self.elements = sorted(self.df['Element'].unique())
                    self.lines = sorted(self.df['Line'].unique())
                    data_found = True
                    break
                except Exception as e:
                    continue

        if not data_found:
            tk.messagebox.showerror("Error", f"Failed to load data: NIST_BE file not found in expected locations")
            self.quit()

    def create_frames(self):
        """Create main layout frames"""
        # Top frame for periodic table
        self.periodic_frame = tk.Frame(self, bg="#f0f0f0", padx=10, pady=10)
        self.periodic_frame.pack(fill=tk.BOTH, expand=False)

        # Middle frame for search options
        self.search_frame = tk.Frame(self, bg="#e0e0e0", padx=10, pady=10)
        self.search_frame.pack(fill=tk.X, expand=False)

        # Bottom frame for results table
        self.results_frame = tk.Frame(self, padx=10, pady=10)
        self.results_frame.pack(fill=tk.BOTH, expand=True)

    def create_periodic_table(self):
        """Create the periodic table buttons"""
        # Define element positions in the periodic table
        self.element_positions = self.get_element_positions()

        # Create a frame for the periodic table grid
        pt_grid = tk.Frame(self.periodic_frame)
        pt_grid.pack(fill=tk.BOTH, expand=True)

        # Define color schemes based on element categories
        colors = {
            'alkali_metal': "#FF6666",  # Light red
            'alkaline_earth': "#FFDEAD",  # Navajo white
            'transition_metal': "#FFC0CB",  # Pink
            'post_transition': "#CCCCCC",  # Light gray
            'metalloid': "#97FFFF",  # Light cyan
            'nonmetal': "#A0FFA0",  # Light green
            'halogen': "#FFFF99",  # Light yellow
            'noble_gas': "#C8A2C8",  # Lilac
            'lanthanide': "#FFBFFF",  # Light magenta
            'actinide': "#FF99CC",  # Light pink
            'unknown': "#E8E8E8"  # Very light gray
        }

        # Define element categories
        element_categories = {
            'H': 'nonmetal',
            'He': 'noble_gas',
            'Li': 'alkali_metal', 'Na': 'alkali_metal', 'K': 'alkali_metal',
            'Rb': 'alkali_metal', 'Cs': 'alkali_metal', 'Fr': 'alkali_metal',
            'Be': 'alkaline_earth', 'Mg': 'alkaline_earth', 'Ca': 'alkaline_earth',
            'Sr': 'alkaline_earth', 'Ba': 'alkaline_earth', 'Ra': 'alkaline_earth',
            'Sc': 'transition_metal', 'Ti': 'transition_metal', 'V': 'transition_metal',
            'Cr': 'transition_metal', 'Mn': 'transition_metal', 'Fe': 'transition_metal',
            'Co': 'transition_metal', 'Ni': 'transition_metal', 'Cu': 'transition_metal',
            'Zn': 'transition_metal', 'Y': 'transition_metal', 'Zr': 'transition_metal',
            'Nb': 'transition_metal', 'Mo': 'transition_metal', 'Tc': 'transition_metal',
            'Ru': 'transition_metal', 'Rh': 'transition_metal', 'Pd': 'transition_metal',
            'Ag': 'transition_metal', 'Cd': 'transition_metal', 'Hf': 'transition_metal',
            'Ta': 'transition_metal', 'W': 'transition_metal', 'Re': 'transition_metal',
            'Os': 'transition_metal', 'Ir': 'transition_metal', 'Pt': 'transition_metal',
            'Au': 'transition_metal', 'Hg': 'transition_metal', 'Rf': 'transition_metal',
            'Db': 'transition_metal', 'Sg': 'transition_metal', 'Bh': 'transition_metal',
            'Hs': 'transition_metal', 'Mt': 'transition_metal', 'Ds': 'transition_metal',
            'Rg': 'transition_metal', 'Cn': 'transition_metal',
            'B': 'metalloid', 'Si': 'metalloid', 'Ge': 'metalloid',
            'As': 'metalloid', 'Sb': 'metalloid', 'Te': 'metalloid', 'Po': 'metalloid',
            'C': 'nonmetal', 'N': 'nonmetal', 'O': 'nonmetal', 'P': 'nonmetal',
            'S': 'nonmetal', 'Se': 'nonmetal',
            'F': 'halogen', 'Cl': 'halogen', 'Br': 'halogen', 'I': 'halogen', 'At': 'halogen',
            'Ne': 'noble_gas', 'Ar': 'noble_gas', 'Kr': 'noble_gas',
            'Xe': 'noble_gas', 'Rn': 'noble_gas', 'Og': 'noble_gas',
            'Al': 'post_transition', 'Ga': 'post_transition', 'In': 'post_transition',
            'Sn': 'post_transition', 'Tl': 'post_transition', 'Pb': 'post_transition',
            'Bi': 'post_transition', 'Nh': 'post_transition', 'Fl': 'post_transition',
            'Mc': 'post_transition', 'Lv': 'post_transition', 'Ts': 'post_transition'
        }

        # Add lanthanides and actinides
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']

        for element in lanthanides:
            element_categories[element] = 'lanthanide'
        for element in actinides:
            element_categories[element] = 'actinide'

        # # Create buttons for each element
        # for element, (row, col) in self.element_positions.items():
        #     # Get color based on element category
        #     category = element_categories.get(element, 'unknown')
        #     color = colors.get(category, colors['unknown'])
        #
        #     # Create button
        #     btn = tk.Button(pt_grid, text=element, width=3, height=1,
        #                     bg=color, activebackground=self.darken_color(color),
        #                     command=lambda e=element: self.select_element(e))
        #
        #     # Check if element is in our dataset
        #     if element in self.elements:
        #         btn.config(relief=tk.RAISED)
        #     else:
        #         btn.config(relief=tk.SUNKEN, state=tk.DISABLED)
        #
        #     btn.grid(row=row, column=col, padx=1, pady=1, sticky='nsew')

        # Create buttons for each element
        for element, (row, col) in self.element_positions.items():
            # Get color based on element category
            category = element_categories.get(element, 'unknown')
            color = colors.get(category, colors['unknown'])

            # Create button based on platform
            btn = self.create_mac_button(pt_grid, element, color, row, col)
            btn.grid(row=row, column=col, padx=1, pady=1, sticky='nsew')




        # Add labels for lanthanides and actinides
        tk.Label(pt_grid, text="*", font=("Helvetica", 11)).grid(row=6, column=2)
        tk.Label(pt_grid, text="**", font=("Helvetica", 11)).grid(row=7, column=2)



    def get_element_positions(self) -> Dict[str, Tuple[int, int]]:
        """Define positions for elements in the periodic table grid"""
        positions = {}

        # Period 1
        positions['H'] = (0, 0)
        positions['He'] = (0, 17)

        # Period 2
        positions['Li'] = (1, 0)
        positions['Be'] = (1, 1)
        positions['B'] = (1, 12)
        positions['C'] = (1, 13)
        positions['N'] = (1, 14)
        positions['O'] = (1, 15)
        positions['F'] = (1, 16)
        positions['Ne'] = (1, 17)

        # Period 3
        positions['Na'] = (2, 0)
        positions['Mg'] = (2, 1)
        positions['Al'] = (2, 12)
        positions['Si'] = (2, 13)
        positions['P'] = (2, 14)
        positions['S'] = (2, 15)
        positions['Cl'] = (2, 16)
        positions['Ar'] = (2, 17)

        # Period 4
        positions['K'] = (3, 0)
        positions['Ca'] = (3, 1)
        for i, symbol in enumerate(['Sc', 'Ti', 'V', 'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn']):
            positions[symbol] = (3, i + 2)
        positions['Ga'] = (3, 12)
        positions['Ge'] = (3, 13)
        positions['As'] = (3, 14)
        positions['Se'] = (3, 15)
        positions['Br'] = (3, 16)
        positions['Kr'] = (3, 17)

        # Period 5
        positions['Rb'] = (4, 0)
        positions['Sr'] = (4, 1)
        for i, symbol in enumerate(['Y', 'Zr', 'Nb', 'Mo', 'Tc', 'Ru', 'Rh', 'Pd', 'Ag', 'Cd']):
            positions[symbol] = (4, i + 2)
        positions['In'] = (4, 12)
        positions['Sn'] = (4, 13)
        positions['Sb'] = (4, 14)
        positions['Te'] = (4, 15)
        positions['I'] = (4, 16)
        positions['Xe'] = (4, 17)

        # Period 6
        positions['Cs'] = (5, 0)
        positions['Ba'] = (5, 1)
        positions['La'] = (5, 2)  # Placeholder for lanthanide indicator
        for i, symbol in enumerate(['Hf', 'Ta', 'W', 'Re', 'Os', 'Ir', 'Pt', 'Au', 'Hg']):
            positions[symbol] = (5, i + 3)
        positions['Tl'] = (5, 12)
        positions['Pb'] = (5, 13)
        positions['Bi'] = (5, 14)
        positions['Po'] = (5, 15)
        positions['At'] = (5, 16)
        positions['Rn'] = (5, 17)

        # Period 7
        positions['Fr'] = (6, 0)
        positions['Ra'] = (6, 1)
        positions['Ac'] = (6, 2)  # Placeholder for actinide indicator
        for i, symbol in enumerate(['Rf', 'Db', 'Sg', 'Bh', 'Hs', 'Mt', 'Ds', 'Rg', 'Cn']):
            positions[symbol] = (6, i + 3)
        positions['Nh'] = (6, 12)
        positions['Fl'] = (6, 13)
        positions['Mc'] = (6, 14)
        positions['Lv'] = (6, 15)
        positions['Ts'] = (6, 16)
        positions['Og'] = (6, 17)

        # Lanthanides (Period 8)
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        for i, symbol in enumerate(lanthanides):
            positions[symbol] = (8, i + 2)

        # Actinides (Period 9)
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']
        for i, symbol in enumerate(actinides):
            positions[symbol] = (9, i + 2)

        return positions

    def darken_color(self, hex_color: str) -> str:
        """Darken a hex color for button press effect"""
        # Convert hex to RGB
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

        # Darken by factor
        factor = 0.8
        r = max(int(r * factor), 0)
        g = max(int(g * factor), 0)
        b = max(int(b * factor), 0)

        # Convert back to hex
        return f'#{r:02x}{g:02x}{b:02x}'

    def create_search_area(self):
        """Create the search interface and options"""
        # Set up layout frames
        left_frame = tk.Frame(self.search_frame, bg="#e0e0e0")
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        right_frame = tk.Frame(self.search_frame, bg="#e0e0e0")
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Element selection info
        tk.Label(left_frame, text="Selected Element:", bg="#e0e0e0", font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9)).grid(row=0, column=0, sticky='w', pady=2)
        self.element_label = tk.Label(left_frame, text="None", width=6, relief=tk.SUNKEN, bg="white", font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9))
        self.element_label.grid(row=0, column=1, sticky='w', padx=5, pady=2)

        # Line selection dropdown
        tk.Label(left_frame, text="XPS Line:", bg="#e0e0e0", font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9)).grid(row=1, column=0, sticky='w', pady=2, )
        self.line_var = tk.StringVar()
        combo_width = 7 if self.detect_mac_os() else 15
        self.line_dropdown = ttk.Combobox(left_frame, textvariable=self.line_var, width=combo_width, state="readonly")
        self.line_dropdown['values'] = ['All Lines'] + list(self.lines)
        self.line_dropdown.current(0)
        self.line_dropdown.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        self.line_dropdown.bind('<<ComboboxSelected>>', self.update_results)

        # Create a frame for search boxes
        search_frame = tk.Frame(right_frame, bg="#e0e0e0")
        search_frame.pack(fill=tk.X)

        # Formula search
        tk.Label(search_frame, text="Search Formula:", bg="#e0e0e0",font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9)).grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.on_search_change)
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=20,
                                     font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9))
        self.search_entry.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        # Name search
        tk.Label(search_frame, text="Search Name:", bg="#e0e0e0", font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9)).grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.name_search_var = tk.StringVar()
        self.name_search_var.trace("w", self.on_search_change)
        self.name_search_entry = tk.Entry(search_frame, textvariable=self.name_search_var, width=20,
                                          font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9))
        self.name_search_entry.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # # Reset button
        # self.reset_btn = tk.Button(search_frame, text="Reset All", command=self.reset_all)
        # self.reset_btn.grid(row=0, column=3, padx=10, pady=2, sticky='w')

        # # Set button height based on platform
        # button_height = 2 if self.detect_mac_os() else 1

        # Add element properties button
        self.properties_btn = tk.Button(search_frame, text="Properties", command=self.show_element_properties, font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9))
        self.properties_btn.grid(row=0, column=3, sticky='w', padx=10, pady=2)

        # Plot button - ADD THIS NEW BUTTON
        self.plot_btn = tk.Button(search_frame, text="Plot Results", command=self.plot_results, font=("Helvetica", 12) if self.detect_mac_os() else ("Helvetica", 9))
        self.plot_btn.grid(row=1, column=3, padx=10, pady=2, sticky='w')

    def plot_results(self):
        """Create a matplotlib plot of binding energies with a combobox for resolution control"""
        # Get the currently filtered data
        filtered_df = self.get_filtered_data()

        # Check if there's any data to plot
        if filtered_df.empty:
            messagebox.showinfo("No Data", "There is no data to plot.")
            return

        # Extract binding energies
        binding_energies = filtered_df['BE (eV)'].dropna().values

        if len(binding_energies) == 0:
            messagebox.showinfo("No Data", "No binding energy values to plot.")
            return

        # Create a new window for the plot
        plot_window = tk.Toplevel(self)
        plot_window.title("Binding Energy Distribution")
        plot_window.geometry("800x700")

        # Create a top frame for resolution control
        control_frame = tk.Frame(plot_window, bg="#e0e0e0", pady=10, padx=10)
        control_frame.pack(fill=tk.X, expand=False)

        # Create a clearer label for the resolution control
        tk.Label(control_frame, text="Histogram Resolution:",
                 bg="#e0e0e0", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))

        # Create a StringVar for the resolution value
        resolution_var = tk.StringVar(value="0.1")

        # Define available resolutions
        resolutions = ["0.1", "0.2", "0.3", "0.4", "0.5", "0.6", "0.7", "0.8", "0.9", "1.0"]

        # Create combobox for resolution selection
        resolution_combo = ttk.Combobox(control_frame, textvariable=resolution_var,
                                        values=resolutions, width=5, state="readonly")
        resolution_combo.pack(side=tk.LEFT, padx=5)

        # Add "eV" label after combobox
        tk.Label(control_frame, text="eV",
                 bg="#e0e0e0").pack(side=tk.LEFT, padx=(0, 20))

        # Round the min and max values to create nice-looking bin boundaries
        min_energy = np.floor(np.min(binding_energies) * 10) / 10
        max_energy = np.ceil(np.max(binding_energies) * 10) / 10

        # Create a frame for the plot
        plot_frame = tk.Frame(plot_window)
        plot_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create the figure and canvas
        fig = Figure(figsize=(10, 6), dpi=100)
        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Add navigation toolbar at the bottom with all standard matplotlib settings
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar_frame = tk.Frame(plot_window)
        toolbar_frame.pack(fill=tk.X, padx=10, pady=0)
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()

        # je n sais pas si c'

        # Function to update the plot based on resolution
        def update_plot(event=None):
            # Clear the previous plot
            fig.clear()
            ax = fig.add_subplot(111)

            # Get the selected resolution
            bin_width = float(resolution_var.get())

            # Create the bins with the specified resolution
            bins = np.arange(min_energy, max_energy + bin_width, bin_width)

            # Create histogram
            counts, bins, patches = ax.hist(binding_energies, bins=bins, alpha=0.7,
                                            color='skyblue', edgecolor='black')

            # Add a smooth curve (kernel density estimation)
            from scipy.stats import gaussian_kde
            kde = gaussian_kde(binding_energies, bw_method='silverman')
            x = np.linspace(min_energy, max_energy, 1000)
            y = kde(x) * len(binding_energies) * bin_width  # Scale to match histogram height
            ax.plot(x, y, 'r-', linewidth=2)

            # Add labels and title
            element_str = f" for {self.selected_element}" if self.selected_element else ""
            line_str = f" ({self.line_var.get()})" if self.line_var.get() != "All Lines" else ""

            ax.set_xlabel('Binding Energy (eV)')
            ax.set_ylabel('Number of References')
            ax.set_title(f'Binding Energy Distribution{element_str}{line_str}')

            # Add grid
            ax.grid(True, linestyle='--', alpha=0.7)

            # Show the total number of references
            ax.text(0.98, 0.95, f'Total References: {len(binding_energies)}',
                    transform=ax.transAxes, ha='right', va='top',
                    bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

            # Show the current resolution in the plot
            ax.text(0.98, 0.90, f'Resolution: {bin_width} eV',
                    transform=ax.transAxes, ha='right', va='top',
                    bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

            # Refresh the canvas
            canvas.draw()

        # Bind the update function to the combobox selection
        resolution_combo.bind("<<ComboboxSelected>>", update_plot)

        # # Add a button frame at the bottom
        # button_frame = tk.Frame(plot_window)
        # button_frame.pack(fill=tk.X, padx=10, pady=10)
        #
        # # Add a close button
        # close_button = tk.Button(button_frame, text="Close", command=plot_window.destroy,
        #                          width=10, height=1)
        # close_button.pack(side=tk.RIGHT)

        # Initial plot with default resolution of 0.1 eV
        update_plot()

    def create_results_table(self):
        """Create the results table to display XPS data"""
        # Create a frame for the table
        table_frame = tk.Frame(self.results_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        # Modify the Treeview style to allow smaller columns
        style = ttk.Style()
        if self.detect_mac_os():
            style.configure("Treeview", rowheight=24, font=('Helvetica', 11))
            style.configure("Treeview.Heading", font=('Helvetica', 12))
        else:
            style.configure("Treeview", rowheight=20, font=('Helvetica', 9))
            style.configure("Treeview.Heading", font=('Helvetica', 10))

        # Create treeview widget
        columns = ("", "Line", "BE (eV)", "Formula", "Name", "Journal")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Treeview")

        # Define column widths and headings

        widths = {"": 30, "Line": 40, "BE (eV)": 60, "Formula": 100, "Name": 150, "Journal": 220}
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=widths[col], stretch=False)

        # Add scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout
        self.tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')

        # Configure grid weights
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Status bar for results count
        self.status_bar = tk.Label(self.results_frame, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind click event for URL opening
        self.tree.bind("<Double-1>", self.on_item_double_click)

        # Add right-click context menu - platform specific
        if self.detect_mac_os():
            # Mac uses Button-2 for right-click or Control+Button-1
            self.tree.bind("<Button-2>", self.show_context_menu)
            self.tree.bind("<Control-Button-1>", self.show_context_menu)
        else:
            # Windows/Linux use Button-3
            self.tree.bind("<Button-3>", self.show_context_menu)

        # Create context menu
        self.context_menu = Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="Copy Full Reference", command=self.copy_reference)
        self.context_menu.add_command(label="Copy Journal Only", command=self.copy_journal_only)
        self.context_menu.add_command(label="Search in Google Scholar", command=self.search_google_scholar)
        self.context_menu.add_separator()  # Add a separator
        self.context_menu.add_command(label="Show Full Information", command=self.show_full_info)

        # Create sort direction trackers
        self.sort_direction = {}
        for col in columns:
            self.sort_direction[col] = False  # Start with ascending

    def sort_column(self, col):
        """Sort treeview contents when a column header is clicked"""
        # Toggle sort direction
        self.sort_direction[col] = not self.sort_direction[col]

        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Get filtered data
        filtered_df = self.get_filtered_data()

        # Apply sorting
        if col == "BE (eV)":
            # For numeric column
            filtered_df = filtered_df.sort_values(by=col, ascending=self.sort_direction[col])
        else:
            # For text columns
            filtered_df = filtered_df.sort_values(by=col, ascending=self.sort_direction[col])

        # Populate the tree with sorted data
        self.populate_tree(filtered_df)

        # Update header text to show sort direction
        for column in self.tree["columns"]:
            if column == col:
                direction = "▲" if self.sort_direction[col] else "▼"
                self.tree.heading(column, text=f"{column} {direction}")
            else:
                # Remove sort indicators from other columns
                self.tree.heading(column, text=column.replace(" ▲", "").replace(" ▼", ""))

    def copy_journal_only(self):
        """Copy just the journal information to clipboard"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, "values")

        # Journal is the 6th column (index 5)
        journal = values[5]

        # Copy to clipboard
        try:
            pyperclip.copy(journal)
            self.status_bar.config(text="Journal copied to clipboard")
        except:
            self.status_bar.config(text="Failed to copy journal")

    def get_filtered_data_OLD(self):
        """Get the filtered dataframe based on current selections"""
        filtered_df = self.df.copy()

        # Filter by element if selected
        if self.selected_element:
            filtered_df = filtered_df[filtered_df['Element'] == self.selected_element]

        # Filter by line if not 'All Lines'
        selected_line = self.line_var.get()
        if selected_line != 'All Lines':
            filtered_df = filtered_df[filtered_df['Line'] == selected_line]

        # Filter by Formula search text
        formula_search = self.search_var.get().strip().lower()
        if formula_search:
            filtered_df = filtered_df[filtered_df['Formula'].str.lower().str.contains(formula_search, na=False)]

        # Filter by Name search text
        name_search = self.name_search_var.get().strip().lower()
        if name_search:
            filtered_df = filtered_df[filtered_df['Name'].str.lower().str.contains(name_search, na=False)]

        return filtered_df

    def get_filtered_data(self):
        """Get the filtered dataframe based on current selections"""
        import re  # Add this import at the top of your file if not already there

        filtered_df = self.df.copy()

        # Filter by element if selected
        if self.selected_element:
            filtered_df = filtered_df[filtered_df['Element'] == self.selected_element]

        # Filter by line if not 'All Lines'
        selected_line = self.line_var.get()
        if selected_line != 'All Lines':
            filtered_df = filtered_df[filtered_df['Line'] == selected_line]

        # Filter by Formula search text
        formula_search = self.search_var.get().strip().lower()
        if formula_search:
            # Escape special regex characters
            formula_search = re.escape(formula_search)
            filtered_df = filtered_df[filtered_df['Formula'].str.lower().str.contains(formula_search, na=False)]

        # Filter by Name search text
        name_search = self.name_search_var.get().strip().lower()
        if name_search:
            # Escape special regex characters
            name_search = re.escape(name_search)
            filtered_df = filtered_df[filtered_df['Name'].str.lower().str.contains(name_search, na=False)]

        return filtered_df

    def populate_tree(self, df):
        """Populate the treeview with data from the dataframe"""
        count = 0
        for _, row in df.iterrows():
            values = (
                row['Element'],
                row['Line'],
                f"{row['BE (eV)']:.2f}" if pd.notnull(row['BE (eV)']) else "",
                row['Formula'] if pd.notnull(row['Formula']) else "",
                row['Name'] if pd.notnull(row['Name']) else "",
                row['Journal'] if pd.notnull(row['Journal']) else ""
            )

            # Store URL in the item's values (not shown in the table)
            item_id = self.tree.insert("", "end", values=values)
            # self.tree.item(item_id, tags=(row['URL'] if pd.notnull(row['URL']) else "",))
            count += 1

        # Update status bar
        self.status_bar.config(text=f"{count} results found")

    def show_context_menu(self, event):
        """Show context menu on right-click"""
        # Get the item that was clicked
        item = self.tree.identify_row(event.y)
        if not item:
            return

        # Select the item
        self.tree.selection_set(item)

        # Show the context menu regardless of which column was clicked
        self.context_menu.post(event.x_root, event.y_root)

        # # Only show context menu if clicked on Journal column
        # col = self.tree.identify_column(event.x)
        # col_index = int(col.replace('#', '')) - 1
        #
        # # Journal is the 6th column (index 5)
        # if col_index == 5:
        #     self.context_menu.post(event.x_root, event.y_root)

    def search_google_scholar(self):
        """Open Google Scholar with the journal text as search query"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, "values")

        # Journal is the 6th column (index 5)
        journal = values[5]

        if journal and journal.strip():
            # Encode the journal text for a URL
            import urllib.parse
            query = urllib.parse.quote(journal)

            # Construct the Google Scholar URL
            scholar_url = f"https://scholar.google.com/scholar?q={query}"

            # Open in default browser
            webbrowser.open(scholar_url)
            self.status_bar.config(text="Opened in Google Scholar")
        else:
            self.status_bar.config(text="No journal information to search")

    def copy_reference(self):
        """Copy the full reference to clipboard"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, "values")

        # Construct a complete reference from available data
        element = values[0]
        line = values[1]
        be = values[2]
        formula = values[3]
        name = values[4]
        journal = values[5]

        reference = f"{element} {line} - {be} eV - {formula} - {name} - {journal}"

        # Copy to clipboard
        try:
            pyperclip.copy(reference)
            self.status_bar.config(text="Reference copied to clipboard")
        except:
            self.status_bar.config(text="Failed to copy reference")



    def select_element(self, element: str):
        """Handle element selection from periodic table"""
        self.selected_element = element
        self.element_label.config(text=element)

        # Update line dropdown to only show lines available for selected element
        element_lines = ['All Lines'] + sorted(self.df[self.df['Element'] == element]['Line'].unique().tolist())
        self.line_dropdown['values'] = element_lines
        self.line_dropdown.current(0)

        self.update_results()

    def update_results_OLD(self, event=None):
        """Update the results table based on current selections"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Apply filters
        filtered_df = self.df.copy()

        # Filter by element if selected
        if self.selected_element:
            filtered_df = filtered_df[filtered_df['Element'] == self.selected_element]

        # Filter by line if not 'All Lines'
        selected_line = self.line_var.get()
        if selected_line != 'All Lines':
            filtered_df = filtered_df[filtered_df['Line'] == selected_line]

        # Filter by search text
        search_text = self.search_var.get().strip().lower()
        if search_text:
            filtered_df = filtered_df[filtered_df['Formula'].str.lower().str.contains(search_text, na=False)]

        # Sort by BE
        filtered_df = filtered_df.sort_values(by='BE (eV)')

        # Populate the tree
        count = 0
        for _, row in filtered_df.iterrows():
            values = (
                row['Element'],
                row['Line'],
                f"{row['BE (eV)']:.2f}" if pd.notnull(row['BE (eV)']) else "",
                row['Formula'] if pd.notnull(row['Formula']) else "",
                row['Name'] if pd.notnull(row['Name']) else "",
                row['Journal'] if pd.notnull(row['Journal']) else ""
            )

            # Store URL in the item's values (not shown in the table)
            item_id = self.tree.insert("", "end", values=values)
            self.tree.item(item_id, tags=(row['URL'] if pd.notnull(row['URL']) else "",))
            count += 1

        # Update status bar
        self.status_bar.config(text=f"{count} results found")

    def update_results(self, event=None):
        """Update the results table based on current selections"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Get filtered data
        filtered_df = self.get_filtered_data()

        # Default sort by BE (eV)
        filtered_df = filtered_df.sort_values(by='BE (eV)')

        # Populate the tree
        self.populate_tree(filtered_df)

    def on_search_change(self, *args):
        """Handle search text changes"""
        self.update_results()

    def on_item_double_click(self, event):
        """Handle double-click on table row to show full info"""
        item = self.tree.selection()
        if not item:
            return

        # Show full information instead of opening URL
        self.show_full_info()

    def show_full_info(self):
        """Open a new window with full information about the selected item"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.tree.item(item, "values")

        # Get element, line, and BE to find the exact row in the original dataframe
        element = values[0]
        line = values[1]
        be = float(values[2]) if values[2] else None
        formula = values[3]

        # Find matching rows in the dataframe
        if be is not None:
            matches = self.df[(self.df['Element'] == element) &
                              (self.df['Line'] == line) &
                              (abs(self.df['BE (eV)'] - be) < 0.01) &  # Using approximate match for floating point
                              (self.df['Formula'] == formula)]
        else:
            matches = self.df[(self.df['Element'] == element) &
                              (self.df['Line'] == line) &
                              (self.df['Formula'] == formula)]

        if matches.empty:
            self.status_bar.config(text="No detailed information found")
            return

        # Get the first match (should be only one)
        row_data = matches.iloc[0]

        # Create a new window
        detail_window = tk.Toplevel(self)
        detail_window.title(f"Full Information: {element} {line} - {be} eV")
        detail_window.geometry("400x500")

        # Create a frame for the details table
        frame = tk.Frame(detail_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create a two-column Treeview widget
        detail_tree = ttk.Treeview(frame, columns=("Property", "Value"), show="headings")
        detail_tree.heading("Property", text="Property")
        detail_tree.heading("Value", text="Value")

        # Set column widths
        detail_tree.column("Property", width=200)
        detail_tree.column("Value", width=200)

        # Set up scrollbars
        vsb = ttk.Scrollbar(frame, orient="vertical", command=detail_tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=detail_tree.xview)
        detail_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout
        detail_tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')

        # Configure grid weights
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        # Insert data into the tree
        for col in self.df.columns:
            value = row_data[col]
            if pd.notnull(value):
                detail_tree.insert("", tk.END, values=(col, value))

        # Add a button to close the window
        tk.Button(detail_window, text="Close", command=detail_window.destroy).pack(pady=10)

    def reset_all(self):
        """Reset all selections"""
        self.selected_element = None
        self.element_label.config(text="None")
        self.line_dropdown.current(0)
        self.search_var.set("")
        self.name_search_var.set("")
        self.update_results()

    def create_mac_button(self, parent, element, color, row, col):
        """Create macOS canvas-based button"""
        # Create canvas
        canvas = tk.Canvas(parent, width=28, height=28, bg=color,
                           highlightthickness=1, highlightbackground="#666666")

        # Add text
        text_id = canvas.create_text(14, 14, text=element, font=("Helvetica", 11),
                                     fill="black" if element in self.elements else "#888888")

        # Store element and enabled state
        canvas.element = element
        canvas.enabled = element in self.elements
        canvas.default_color = color
        canvas.text_id = text_id

        # Bind events only if enabled
        if canvas.enabled:
            canvas.bind("<Button-1>", lambda e: self.select_element(element))
            canvas.bind("<Double-1>", lambda e: self.on_element_double_click(element))
            canvas.bind("<Enter>", lambda e: self.on_canvas_enter(canvas))
            canvas.bind("<Leave>", lambda e: self.on_canvas_leave(canvas))
        else:
            canvas.config(bg=self.lighten_color(color))

        return canvas

    def lighten_color(self, hex_color: str) -> str:
        """Lighten a hex color for disabled state"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

        # Lighten by factor
        factor = 1.3
        r = min(int(r * factor), 255)
        g = min(int(g * factor), 255)
        b = min(int(b * factor), 255)

        return f'#{r:02x}{g:02x}{b:02x}'

    def on_canvas_enter(self, canvas):
        """Handle mouse enter on canvas button"""
        if canvas.enabled:
            darker_color = self.darken_color(canvas.default_color)
            canvas.config(bg=darker_color)

    def on_canvas_leave(self, canvas):
        """Handle mouse leave on canvas button"""
        if canvas.enabled:
            canvas.config(bg=canvas.default_color)

    def create_windows_button(self, parent, element, color, row, col):
        """Create Windows-optimized button (original logic)"""
        btn = tk.Button(parent, text=element, width=3, height=1,
                        bg=color, activebackground=self.darken_color(color),
                        command=lambda e=element: self.select_element(e))

        # Add double-click binding
        btn.bind("<Double-1>", lambda e: self.on_element_double_click(element))

        # Check if element is in our dataset
        if element in self.elements:
            btn.config(relief=tk.RAISED)
        else:
            btn.config(relief=tk.SUNKEN, state=tk.DISABLED)

        return btn

    def on_element_double_click(self, element):
        """Handle double-click on element to open properties"""
        if element in self.elements:
            self.select_element(element)
            self.show_element_properties()



    def show_element_properties(self):
        """Show properties for the selected element"""
        if not self.selected_element:
            messagebox.showinfo("No Element Selected", "Please select an element from the periodic table first.")
            return

        # Element properties data
        element_data = self.get_element_properties(self.selected_element)

        # Create a new window
        properties_window = tk.Toplevel(self)
        properties_window.title(f"Properties for {self.selected_element}")
        properties_window.geometry("550x500")
        properties_window.resizable(True, True)

        # Add a title
        title_frame = tk.Frame(properties_window, bg="#f0f0f0", padx=10, pady=10)
        title_frame.pack(fill=tk.X, expand=False)

        # Element symbol in large font
        symbol_label = tk.Label(title_frame, text=self.selected_element,
                                font=("Helvetica", 40, "bold"), bg="#f0f0f0")
        symbol_label.pack(side=tk.LEFT, padx=20)

        # Element name and basic info
        info_frame = tk.Frame(title_frame, bg="#f0f0f0")
        info_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20)

        name_label = tk.Label(info_frame, text=element_data["Name"],
                              font=("Helvetica", 20), bg="#f0f0f0")
        name_label.pack(anchor="w")

        atomic_number_label = tk.Label(info_frame, text=f"Atomic Number: {element_data['Atomic Number']}",
                                       font=("Helvetica", 12), bg="#f0f0f0")
        atomic_number_label.pack(anchor="w")

        # Main properties grid
        grid_frame = tk.Frame(properties_window, padx=20, pady=10)
        grid_frame.pack(fill=tk.BOTH, expand=True)

        # Create a canvas with scrollbar for the properties
        canvas = tk.Canvas(grid_frame)
        scrollbar = ttk.Scrollbar(grid_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Define property groups and layout
        property_groups = {
            "Physical Properties": ["Atomic Mass", "Density", "Melting Point", "Boiling Point", "State at 20°C"],
            "Atomic Properties": ["Electron Configuration", "Electronegativity", "Atomic Radius", "Ionization Energy"],
            "General Information": ["Group", "Period", "Category", "Discovered By", "Year of Discovery"],
            "XPS Information": ["Common Core Levels", "Most Intense Line", "Typical FWHM", "Chemical Shift Range"]
        }

        # Add property groups to the grid
        row = 0
        for group_name, properties in property_groups.items():
            # Group header
            group_label = tk.Label(scrollable_frame, text=group_name,
                                   font=("Helvetica", 12, "bold"), bg="#e0e0e0",
                                   padx=5, pady=5, relief=tk.RIDGE)
            group_label.grid(row=row, column=0, columnspan=2, sticky="ew", padx=5, pady=10)
            row += 1

            # Properties in this group
            for prop in properties:
                prop_label = tk.Label(scrollable_frame, text=prop + ":",
                                      font=("Helvetica", 10), anchor="e", padx=5, pady=3)
                prop_label.grid(row=row, column=0, sticky="e")

                value = element_data.get(prop, "N/A")
                value_label = tk.Label(scrollable_frame, text=str(value),
                                       font=("Helvetica", 10), anchor="w", padx=5, pady=3,
                                       relief=tk.SUNKEN, bg="white", width=40)
                value_label.grid(row=row, column=1, sticky="w")

                row += 1

        # XPS Binding Energies Summary
        xps_label = tk.Label(scrollable_frame, text="XPS Binding Energies",
                             font=("Helvetica", 12, "bold"), bg="#e0e0e0",
                             padx=5, pady=5, relief=tk.RIDGE)
        xps_label.grid(row=row, column=0, columnspan=2, sticky="ew", padx=5, pady=10)
        row += 1

        # Get XPS data for this element
        element_xps_data = self.df[self.df['Element'] == self.selected_element]
        if not element_xps_data.empty:
            # Group by Line and calculate average BE
            lines_summary = element_xps_data.groupby('Line')['BE (eV)'].agg(
                ['mean', 'count', 'min', 'max']).reset_index()

            # Create a frame to contain the table
            table_container = tk.Frame(scrollable_frame)
            table_container.grid(row=row, column=0, columnspan=2, sticky="ew", padx=5)

            # Headers with fixed widths
            headers = ["Line", "Avg BE (eV)", "Min BE (eV)", "Max BE (eV)", "Count"]

            # Create table headers
            for col, header in enumerate(headers):
                header_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                header_frame.grid(row=0, column=col, sticky="ew")

                label = tk.Label(header_frame, text=header, font=("Helvetica", 10, "bold"),
                                 bg="#e8e8e8", padx=5, pady=2)
                label.pack(fill="both", expand=True)

            # Configure column widths
            table_container.columnconfigure(0, minsize=80)  # Line
            table_container.columnconfigure(1, minsize=80)  # Avg BE
            table_container.columnconfigure(2, minsize=80)  # Min BE
            table_container.columnconfigure(3, minsize=80)  # Max BE
            table_container.columnconfigure(4, minsize=60)  # Count

            # Add data rows
            for i, (_, line_data) in enumerate(lines_summary.iterrows()):
                row_idx = i + 1

                # Line column
                cell_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                cell_frame.grid(row=row_idx, column=0, sticky="ew")
                tk.Label(cell_frame, text=line_data['Line'], font=("Helvetica", 10),
                         padx=5, pady=2).pack(fill="both")

                # Avg BE column
                cell_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                cell_frame.grid(row=row_idx, column=1, sticky="ew")
                tk.Label(cell_frame, text=f"{line_data['mean']:.2f}", font=("Helvetica", 10),
                         padx=5, pady=2).pack(fill="both")

                # Min BE column
                cell_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                cell_frame.grid(row=row_idx, column=2, sticky="ew")
                tk.Label(cell_frame, text=f"{line_data['min']:.2f}", font=("Helvetica", 10),
                         padx=5, pady=2).pack(fill="both")

                # Max BE column
                cell_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                cell_frame.grid(row=row_idx, column=3, sticky="ew")
                tk.Label(cell_frame, text=f"{line_data['max']:.2f}", font=("Helvetica", 10),
                         padx=5, pady=2).pack(fill="both")

                # Count column
                cell_frame = tk.Frame(table_container, relief=tk.RIDGE, borderwidth=1)
                cell_frame.grid(row=row_idx, column=4, sticky="ew")
                tk.Label(cell_frame, text=str(line_data['count']), font=("Helvetica", 10),
                         padx=5, pady=2).pack(fill="both")

            row += 1
        else:
            no_data_label = tk.Label(scrollable_frame, text="No XPS data available for this element",
                                     font=("Helvetica", 10, "italic"), padx=5, pady=5)
            no_data_label.grid(row=row, column=0, columnspan=2, sticky="ew")
            row += 1

    def get_element_properties(self, element_symbol):
        """Get properties for a specific element"""
        # Comprehensive database of element properties based on reference tables
        element_properties = {
            "H": {
                "Name": "Hydrogen",
                "Atomic Number": 1,
                "Atomic Mass": "1.00794 u",
                "Density": "0.0708 g/cm³",
                "Melting Point": "-259.34 °C",
                "Boiling Point": "-252.87 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 2.20,
                "Ionization Energy": "13.598 eV",
                "Specific Heat": "14.304 J/(g·K)",
                "Group": 1,
                "Period": 1,
                "Category": "Nonmetal",
                "Common Core Levels": "1s",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.2 eV",
                "Chemical Shift Range": "0-3 eV"
            },
            "He": {
                "Name": "Helium",
                "Atomic Number": 2,
                "Atomic Mass": "4.002602 u",
                "Density": "0.122 g/cm³",
                "Melting Point": "— °C",
                "Boiling Point": "-268.93 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "24.587 eV",
                "Specific Heat": "5.193 J/(g·K)",
                "Group": 18,
                "Period": 1,
                "Category": "Noble Gas"
            },
            "Li": {
                "Name": "Lithium",
                "Atomic Number": 3,
                "Atomic Mass": "6.941 u",
                "Density": "0.534 g/cm³",
                "Melting Point": "180.50 °C",
                "Boiling Point": "1342 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "5.392 eV",
                "Specific Heat": "3.582 J/(g·K)",
                "Group": 1,
                "Period": 2,
                "Category": "Alkali Metal"
            },
            "Be": {
                "Name": "Beryllium",
                "Atomic Number": 4,
                "Atomic Mass": "9.012182 u",
                "Density": "1.848 g/cm³",
                "Melting Point": "1287 °C",
                "Boiling Point": "2471 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "9.323 eV",
                "Specific Heat": "1.825 J/(g·K)",
                "Group": 2,
                "Period": 2,
                "Category": "Alkaline Earth Metal"
            },
            "B": {
                "Name": "Boron",
                "Atomic Number": 5,
                "Atomic Mass": "10.811 u",
                "Density": "2.34 g/cm³",
                "Melting Point": "2075 °C",
                "Boiling Point": "4000 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s² 2p¹",
                "Ground State": "²P°₁/₂",
                "Ionization Energy": "8.298 eV",
                "Specific Heat": "1.026 J/(g·K)",
                "Group": 13,
                "Period": 2,
                "Category": "Metalloid"
            },
            "C": {
                "Name": "Carbon",
                "Atomic Number": 6,
                "Atomic Mass": "12.0107 u",
                "Density": "1.9-2.3 (graphite) g/cm³",
                "Melting Point": "4492 (10.3 MPa) °C",
                "Boiling Point": "3825ᵇ °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s² 2p²",
                "Ground State": "³P₀",
                "Ionization Energy": "11.260 eV",
                "Specific Heat": "0.709 J/(g·K)",
                "Group": 14,
                "Period": 2,
                "Category": "Nonmetal",
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.8-1.2 eV",
                "Chemical Shift Range": "0-10 eV"
            },
            "N": {
                "Name": "Nitrogen",
                "Atomic Number": 7,
                "Atomic Mass": "14.00674 u",
                "Density": "0.808 g/cm³",
                "Melting Point": "-210.00 °C",
                "Boiling Point": "-195.79 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p³",
                "Ground State": "⁴S°₃/₂",
                "Ionization Energy": "14.534 eV",
                "Specific Heat": "1.040 J/(g·K)",
                "Group": 15,
                "Period": 2,
                "Category": "Nonmetal"
            },
            "O": {
                "Name": "Oxygen",
                "Atomic Number": 8,
                "Atomic Mass": "15.9994 u",
                "Density": "1.14 g/cm³",
                "Melting Point": "-218.79 °C",
                "Boiling Point": "-182.95 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁴",
                "Ground State": "³P₂",
                "Ionization Energy": "13.618 eV",
                "Specific Heat": "0.918 J/(g·K)",
                "Group": 16,
                "Period": 2,
                "Category": "Nonmetal",
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.3 eV",
                "Chemical Shift Range": "0-8 eV"
            },
            "F": {
                "Name": "Fluorine",
                "Atomic Number": 9,
                "Atomic Mass": "18.9984032 u",
                "Density": "1.50 g/cm³",
                "Melting Point": "-219.62 °C",
                "Boiling Point": "-188.12 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁵",
                "Ground State": "²P°₃/₂",
                "Ionization Energy": "17.423 eV",
                "Specific Heat": "0.824 J/(g·K)",
                "Group": 17,
                "Period": 2,
                "Category": "Halogen"
            },
            "Ne": {
                "Name": "Neon",
                "Atomic Number": 10,
                "Atomic Mass": "20.1797 u",
                "Density": "1.207 g/cm³",
                "Melting Point": "-248.59 °C",
                "Boiling Point": "-246.08 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁶",
                "Ground State": "¹S₀",
                "Ionization Energy": "21.565 eV",
                "Specific Heat": "1.030 J/(g·K)",
                "Group": 18,
                "Period": 2,
                "Category": "Noble Gas"
            },
            "Na": {
                "Name": "Sodium",
                "Atomic Number": 11,
                "Atomic Mass": "22.989770 u",
                "Density": "0.971 g/cm³",
                "Melting Point": "97.80 °C",
                "Boiling Point": "883 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "5.139 eV",
                "Specific Heat": "1.228 J/(g·K)",
                "Group": 1,
                "Period": 3,
                "Category": "Alkali Metal"
            },
            "Mg": {
                "Name": "Magnesium",
                "Atomic Number": 12,
                "Atomic Mass": "24.3050 u",
                "Density": "1.738 g/cm³",
                "Melting Point": "650 °C",
                "Boiling Point": "1090 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "7.646 eV",
                "Specific Heat": "1.023 J/(g·K)",
                "Group": 2,
                "Period": 3,
                "Category": "Alkaline Earth Metal"
            },
            "Al": {
                "Name": "Aluminum",
                "Atomic Number": 13,
                "Atomic Mass": "26.981538 u",
                "Density": "2.6989 g/cm³",
                "Melting Point": "660.32 °C",
                "Boiling Point": "2519 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p¹",
                "Ground State": "²P°₁/₂",
                "Ionization Energy": "5.986 eV",
                "Specific Heat": "0.897 J/(g·K)",
                "Group": 13,
                "Period": 3,
                "Category": "Post-Transition Metal"
            },
            "Si": {
                "Name": "Silicon",
                "Atomic Number": 14,
                "Atomic Mass": "28.0855 u",
                "Density": "2.3325 g/cm³",
                "Melting Point": "1414 °C",
                "Boiling Point": "3265 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p²",
                "Ground State": "³P₀",
                "Ionization Energy": "8.152 eV",
                "Specific Heat": "0.705 J/(g·K)",
                "Group": 14,
                "Period": 3,
                "Category": "Metalloid"
            },
            "P": {
                "Name": "Phosphorus",
                "Atomic Number": 15,
                "Atomic Mass": "30.973761 u",
                "Density": "1.82 g/cm³",
                "Melting Point": "44.15 °C",
                "Boiling Point": "280.5 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p³",
                "Ground State": "⁴S°₃/₂",
                "Ionization Energy": "10.487 eV",
                "Specific Heat": "0.769 J/(g·K)",
                "Group": 15,
                "Period": 3,
                "Category": "Nonmetal"
            },
            "S": {
                "Name": "Sulfur",
                "Atomic Number": 16,
                "Atomic Mass": "32.066 u",
                "Density": "2.07 g/cm³",
                "Melting Point": "119.6 °C",
                "Boiling Point": "444.60 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p⁴",
                "Ground State": "³P₂",
                "Ionization Energy": "10.360 eV",
                "Specific Heat": "0.710 J/(g·K)",
                "Group": 16,
                "Period": 3,
                "Category": "Nonmetal"
            },
            "Cl": {
                "Name": "Chlorine",
                "Atomic Number": 17,
                "Atomic Mass": "35.4527 u",
                "Density": "1.56 (3.3, 6) g/cm³",
                "Melting Point": "-101.5 °C",
                "Boiling Point": "-34.04 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ne] 3s² 3p⁵",
                "Ground State": "²P°₃/₂",
                "Ionization Energy": "12.968 eV",
                "Specific Heat": "0.479 J/(g·K)",
                "Group": 17,
                "Period": 3,
                "Category": "Halogen"
            },
            "Ar": {
                "Name": "Argon",
                "Atomic Number": 18,
                "Atomic Mass": "39.948 u",
                "Density": "1.40 g/cm³",
                "Melting Point": "-189.35 °C",
                "Boiling Point": "-185.85 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ne] 3s² 3p⁶",
                "Ground State": "¹S₀",
                "Ionization Energy": "15.760 eV",
                "Specific Heat": "0.520 J/(g·K)",
                "Group": 18,
                "Period": 3,
                "Category": "Noble Gas"
            },
            "K": {
                "Name": "Potassium",
                "Atomic Number": 19,
                "Atomic Mass": "39.0983 u",
                "Density": "0.862 g/cm³",
                "Melting Point": "63.5 °C",
                "Boiling Point": "759 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 4s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "4.341 eV",
                "Specific Heat": "0.757 J/(g·K)",
                "Group": 1,
                "Period": 4,
                "Category": "Alkali Metal"
            },
            "Ca": {
                "Name": "Calcium",
                "Atomic Number": 20,
                "Atomic Mass": "40.078 u",
                "Density": "1.55 g/cm³",
                "Melting Point": "842 °C",
                "Boiling Point": "1484 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 4s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "6.113 eV",
                "Specific Heat": "0.647 J/(g·K)",
                "Group": 2,
                "Period": 4,
                "Category": "Alkaline Earth Metal"
            },
            "Sc": {
                "Name": "Scandium",
                "Atomic Number": 21,
                "Atomic Mass": "44.955912 u",
                "Density": "2.989 g/cm³",
                "Melting Point": "1541 °C",
                "Boiling Point": "2836 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹ 4s²",
                "Ground State": "²D₃/₂",
                "Ionization Energy": "6.561 eV",
                "Specific Heat": "0.568 J/(g·K)",
                "Group": 3,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Lars Fredrik Nilson",
                "Year of Discovery": 1879
            },
            "Ti": {
                "Name": "Titanium",
                "Atomic Number": 22,
                "Atomic Mass": "47.867 u",
                "Density": "4.54 g/cm³",
                "Melting Point": "1668 °C",
                "Boiling Point": "3287 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d² 4s²",
                "Ground State": "³F₂",
                "Ionization Energy": "6.828 eV",
                "Specific Heat": "0.523 J/(g·K)",
                "Group": 4,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "William Gregor",
                "Year of Discovery": 1791
            },
            "V": {
                "Name": "Vanadium",
                "Atomic Number": 23,
                "Atomic Mass": "50.9415 u",
                "Density": "6.11 g/cm³",
                "Melting Point": "1910 °C",
                "Boiling Point": "3407 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d³ 4s²",
                "Ground State": "⁴F₃/₂",
                "Ionization Energy": "6.746 eV",
                "Specific Heat": "0.489 J/(g·K)",
                "Group": 5,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Andrés Manuel del Río",
                "Year of Discovery": 1801
            },
            "Cr": {
                "Name": "Chromium",
                "Atomic Number": 24,
                "Atomic Mass": "51.9961 u",
                "Density": "7.15 g/cm³",
                "Melting Point": "1907 °C",
                "Boiling Point": "2671 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁵ 4s¹",
                "Ground State": "⁷S₃",
                "Ionization Energy": "6.767 eV",
                "Specific Heat": "0.449 J/(g·K)",
                "Group": 6,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Louis Nicolas Vauquelin",
                "Year of Discovery": 1797
            },
            "Mn": {
                "Name": "Manganese",
                "Atomic Number": 25,
                "Atomic Mass": "54.938045 u",
                "Density": "7.44 g/cm³",
                "Melting Point": "1246 °C",
                "Boiling Point": "2061 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁵ 4s²",
                "Ground State": "⁶S₅/₂",
                "Ionization Energy": "7.434 eV",
                "Specific Heat": "0.479 J/(g·K)",
                "Group": 7,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Johan Gottlieb Gahn",
                "Year of Discovery": 1774
            },
            "Fe": {
                "Name": "Iron",
                "Atomic Number": 26,
                "Atomic Mass": "55.845 u",
                "Density": "7.874 g/cm³",
                "Melting Point": "1538 °C",
                "Boiling Point": "2861 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁶ 4s²",
                "Ground State": "⁵D₄",
                "Ionization Energy": "7.902 eV",
                "Specific Heat": "0.449 J/(g·K)",
                "Group": 8,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Co": {
                "Name": "Cobalt",
                "Atomic Number": 27,
                "Atomic Mass": "58.933200 u",
                "Density": "8.9 g/cm³",
                "Melting Point": "1495 °C",
                "Boiling Point": "2927 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁷ 4s²",
                "Ground State": "⁴F₉/₂",
                "Ionization Energy": "7.881 eV",
                "Specific Heat": "0.421 J/(g·K)",
                "Group": 9,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Georg Brandt",
                "Year of Discovery": 1735
            },
            "Ni": {
                "Name": "Nickel",
                "Atomic Number": 28,
                "Atomic Mass": "58.6934 u",
                "Density": "8.90225 g/cm³",
                "Melting Point": "1455 °C",
                "Boiling Point": "2913 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁸ 4s²",
                "Ground State": "³F₄",
                "Ionization Energy": "7.640 eV",
                "Specific Heat": "0.444 J/(g·K)",
                "Group": 10,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Axel Fredrik Cronstedt",
                "Year of Discovery": 1751
            },
            "Cu": {
                "Name": "Copper",
                "Atomic Number": 29,
                "Atomic Mass": "63.546 u",
                "Density": "8.96 g/cm³",
                "Melting Point": "1084.62 °C",
                "Boiling Point": "2562 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "7.726 eV",
                "Specific Heat": "0.385 J/(g·K)",
                "Group": 11,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Zn": {
                "Name": "Zinc",
                "Atomic Number": 30,
                "Atomic Mass": "65.39 u",
                "Density": "7.13325 g/cm³",
                "Melting Point": "419.53 °C",
                "Boiling Point": "907 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "9.394 eV",
                "Specific Heat": "0.388 J/(g·K)",
                "Group": 12,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Andreas Sigismund Marggraf",
                "Year of Discovery": 1746
            },
            "Ga": {
                "Name": "Gallium",
                "Atomic Number": 31,
                "Atomic Mass": "69.723 u",
                "Density": "5.90429.6 g/cm³",
                "Melting Point": "29.76 °C",
                "Boiling Point": "2204 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p¹",
                "Ground State": "²P₁/₂",
                "Ionization Energy": "5.999 eV",
                "Specific Heat": "0.371 J/(g·K)",
                "Group": 13,
                "Period": 4,
                "Category": "Post-transition Metal",
                "Discovered By": "Lecoq de Boisbaudran",
                "Year of Discovery": 1875
            },
            "Ge": {
                "Name": "Germanium",
                "Atomic Number": 32,
                "Atomic Mass": "72.61 u",
                "Density": "5.32325 g/cm³",
                "Melting Point": "938.25 °C",
                "Boiling Point": "2833 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p²",
                "Ground State": "³P₀",
                "Ionization Energy": "7.899 eV",
                "Specific Heat": "0.320 J/(g·K)",
                "Group": 14,
                "Period": 4,
                "Category": "Metalloid",
                "Discovered By": "Clemens Winkler",
                "Year of Discovery": 1886
            },
            "As": {
                "Name": "Arsenic",
                "Atomic Number": 33,
                "Atomic Mass": "74.92160 u",
                "Density": "5.73 g/cm³",
                "Melting Point": "817 °C (at 28 atm)",
                "Boiling Point": "603 °C (sublimes)",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p³",
                "Ground State": "⁴S³/₂",
                "Ionization Energy": "9.789 eV",
                "Specific Heat": "0.329 J/(g·K)",
                "Group": 15,
                "Period": 4,
                "Category": "Metalloid",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Se": {
                "Name": "Selenium",
                "Atomic Number": 34,
                "Atomic Mass": "78.96 u",
                "Density": "4.79 g/cm³",
                "Melting Point": "220.5 °C",
                "Boiling Point": "685 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁴",
                "Ground State": "³P₂",
                "Ionization Energy": "9.752 eV",
                "Specific Heat": "0.321 J/(g·K)",
                "Group": 16,
                "Period": 4,
                "Category": "Nonmetal",
                "Discovered By": "Jöns Jakob Berzelius",
                "Year of Discovery": 1817
            },
            "Br": {
                "Name": "Bromine",
                "Atomic Number": 35,
                "Atomic Mass": "79.904 u",
                "Density": "3.12 g/cm³",
                "Melting Point": "-7.2 °C",
                "Boiling Point": "58.8 °C",
                "State at 20°C": "Liquid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁵",
                "Ground State": "²P₃/₂",
                "Ionization Energy": "11.814 eV",
                "Specific Heat": "0.226 J/(g·K)",
                "Group": 17,
                "Period": 4,
                "Category": "Halogen",
                "Discovered By": "Antoine Jérôme Balard",
                "Year of Discovery": 1826
            },
            "Kr": {
                "Name": "Krypton",
                "Atomic Number": 36,
                "Atomic Mass": "83.80 u",
                "Density": "2.16 g/cm³",
                "Melting Point": "-157.37 °C (at 73.2 kPa)",
                "Boiling Point": "-153.22 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁶",
                "Ground State": "¹S₀",
                "Ionization Energy": "14.000 eV",
                "Specific Heat": "0.248 J/(g·K)",
                "Group": 18,
                "Period": 4,
                "Category": "Noble Gas",
                "Discovered By": "William Ramsay and Morris Travers",
                "Year of Discovery": 1898
            },
            "Rb": {
                "Name": "Rubidium",
                "Atomic Number": 37,
                "Atomic Mass": "85.4678 u",
                "Density": "1.532 g/cm³",
                "Melting Point": "39.30 °C",
                "Boiling Point": "688 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 5s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "4.177 eV",
                "Specific Heat": "0.363 J/(g·K)",
                "Group": 1,
                "Period": 5,
                "Category": "Alkali Metal",
                "Discovered By": "Robert Bunsen and Gustav Kirchhoff",
                "Year of Discovery": 1861
            },
            "Sr": {
                "Name": "Strontium",
                "Atomic Number": 38,
                "Atomic Mass": "87.62 u",
                "Density": "2.54 g/cm³",
                "Melting Point": "777 °C",
                "Boiling Point": "1382 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 5s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "5.695 eV",
                "Specific Heat": "0.301 J/(g·K)",
                "Group": 2,
                "Period": 5,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Adair Crawford",
                "Year of Discovery": 1790
            },
            "Y": {
                "Name": "Yttrium",
                "Atomic Number": 39,
                "Atomic Mass": "88.90585 u",
                "Density": "4.46925 g/cm³",
                "Melting Point": "1522 °C",
                "Boiling Point": "3345 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹ 5s²",
                "Ground State": "²D₃/₂",
                "Ionization Energy": "6.217 eV",
                "Specific Heat": "0.298 J/(g·K)",
                "Group": 3,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Johan Gadolin",
                "Year of Discovery": 1794
            },
            "Zr": {
                "Name": "Zirconium",
                "Atomic Number": 40,
                "Atomic Mass": "91.224 u",
                "Density": "6.506 g/cm³",
                "Melting Point": "1855 °C",
                "Boiling Point": "4409 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d² 5s²",
                "Ground State": "³F₂",
                "Ionization Energy": "6.634 eV",
                "Specific Heat": "0.278 J/(g·K)",
                "Group": 4,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Martin Heinrich Klaproth",
                "Year of Discovery": 1789
            },
            "Nb": {
                "Name": "Niobium",
                "Atomic Number": 41,
                "Atomic Mass": "92.90638 u",
                "Density": "8.57 g/cm³",
                "Melting Point": "2477 °C",
                "Boiling Point": "4744 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁴ 5s¹",
                "Ground State": "⁶D₁/₂",
                "Ionization Energy": "6.759 eV",
                "Specific Heat": "0.265 J/(g·K)",
                "Group": 5,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Charles Hatchett",
                "Year of Discovery": 1801
            },
            "Mo": {
                "Name": "Molybdenum",
                "Atomic Number": 42,
                "Atomic Mass": "95.94 u",
                "Density": "10.22 g/cm³",
                "Melting Point": "2623 °C",
                "Boiling Point": "4639 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁵ 5s¹",
                "Ground State": "⁷S₃",
                "Ionization Energy": "7.092 eV",
                "Specific Heat": "0.251 J/(g·K)",
                "Group": 6,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Carl Wilhelm Scheele",
                "Year of Discovery": 1778
            },
            "Tc": {
                "Name": "Technetium",
                "Atomic Number": 43,
                "Atomic Mass": "(98) u",
                "Density": "11.50 g/cm³",
                "Melting Point": "2157 °C",
                "Boiling Point": "4265 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁵ 5s²",
                "Ground State": "⁶S₅/₂",
                "Ionization Energy": "7.28 eV",
                "Specific Heat": "—",
                "Group": 7,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Carlo Perrier and Emilio Segrè",
                "Year of Discovery": 1937
            },
            "Ru": {
                "Name": "Ruthenium",
                "Atomic Number": 44,
                "Atomic Mass": "101.07 u",
                "Density": "12.41 g/cm³",
                "Melting Point": "2334 °C",
                "Boiling Point": "4150 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁷ 5s¹",
                "Ground State": "⁵F₅",
                "Ionization Energy": "7.360 eV",
                "Specific Heat": "0.238 J/(g·K)",
                "Group": 8,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Karl Ernst Claus",
                "Year of Discovery": 1844
            },
            "Rh": {
                "Name": "Rhodium",
                "Atomic Number": 45,
                "Atomic Mass": "102.90550 u",
                "Density": "12.41 g/cm³",
                "Melting Point": "1964 °C",
                "Boiling Point": "3695 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁸ 5s¹",
                "Ground State": "⁴F₉/₂",
                "Ionization Energy": "7.459 eV",
                "Specific Heat": "0.243 J/(g·K)",
                "Group": 9,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "William Hyde Wollaston",
                "Year of Discovery": 1804
            },
            "Pd": {
                "Name": "Palladium",
                "Atomic Number": 46,
                "Atomic Mass": "106.42 u",
                "Density": "12.02 g/cm³",
                "Melting Point": "1554.9 °C",
                "Boiling Point": "2963 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰",
                "Ground State": "¹S₀",
                "Ionization Energy": "8.337 eV",
                "Specific Heat": "0.246 J/(g·K)",
                "Group": 10,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "William Hyde Wollaston",
                "Year of Discovery": 1803
            },
            "Ag": {
                "Name": "Silver",
                "Atomic Number": 47,
                "Atomic Mass": "107.8682 u",
                "Density": "10.50 g/cm³",
                "Melting Point": "961.78 °C",
                "Boiling Point": "2162 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "7.576 eV",
                "Specific Heat": "0.235 J/(g·K)",
                "Group": 11,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Cd": {
                "Name": "Cadmium",
                "Atomic Number": 48,
                "Atomic Mass": "112.411 u",
                "Density": "8.65 g/cm³",
                "Melting Point": "321.07 °C",
                "Boiling Point": "767 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "8.994 eV",
                "Specific Heat": "0.232 J/(g·K)",
                "Group": 12,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Friedrich Stromeyer",
                "Year of Discovery": 1817
            },
            "In": {
                "Name": "Indium",
                "Atomic Number": 49,
                "Atomic Mass": "114.818 u",
                "Density": "7.31 g/cm³",
                "Melting Point": "156.60 °C",
                "Boiling Point": "2072 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p¹",
                "Ground State": "²P₁/₂",
                "Ionization Energy": "5.786 eV",
                "Specific Heat": "0.233 J/(g·K)",
                "Group": 13,
                "Period": 5,
                "Category": "Post-transition Metal",
                "Discovered By": "Ferdinand Reich and Hieronymous Theodor Richter",
                "Year of Discovery": 1863
            },
            "Sn": {
                "Name": "Tin",
                "Atomic Number": 50,
                "Atomic Mass": "118.710 u",
                "Density": "7.31 g/cm³",
                "Melting Point": "231.93 °C",
                "Boiling Point": "2602 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p²",
                "Ground State": "³P₀",
                "Ionization Energy": "7.344 eV",
                "Specific Heat": "0.228 J/(g·K)",
                "Group": 14,
                "Period": 5,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Sb": {
                "Name": "Antimony",
                "Atomic Number": 51,
                "Atomic Mass": "121.760 u",
                "Density": "6.691 g/cm³",
                "Melting Point": "630.73 °C",
                "Boiling Point": "1587 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p³",
                "Ground State": "⁴S₃/₂",
                "Ionization Energy": "8.608 eV",
                "Specific Heat": "0.207 J/(g·K)",
                "Group": 15,
                "Period": 5,
                "Category": "Metalloid",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Te": {
                "Name": "Tellurium",
                "Atomic Number": 52,
                "Atomic Mass": "127.60 u",
                "Density": "6.24 g/cm³",
                "Melting Point": "449.51 °C",
                "Boiling Point": "988 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁴",
                "Ground State": "³P₂",
                "Ionization Energy": "9.010 eV",
                "Specific Heat": "0.202 J/(g·K)",
                "Group": 16,
                "Period": 5,
                "Category": "Metalloid",
                "Discovered By": "Franz-Joseph Müller von Reichenstein",
                "Year of Discovery": 1782
            },
            "I": {
                "Name": "Iodine",
                "Atomic Number": 53,
                "Atomic Mass": "126.90447 u",
                "Density": "4.93 g/cm³",
                "Melting Point": "113.7 °C",
                "Boiling Point": "184.4 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁵",
                "Ground State": "²P₃/₂",
                "Ionization Energy": "10.451 eV",
                "Specific Heat": "0.145 J/(g·K)",
                "Group": 17,
                "Period": 5,
                "Category": "Halogen",
                "Discovered By": "Bernard Courtois",
                "Year of Discovery": 1811
            },
            "Xe": {
                "Name": "Xenon",
                "Atomic Number": 54,
                "Atomic Mass": "131.29 u",
                "Density": "3.52 g/cm³",
                "Melting Point": "-111.79 °C (at 81.6 kPa)",
                "Boiling Point": "-108.12 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁶",
                "Ground State": "¹S₀",
                "Ionization Energy": "12.130 eV",
                "Specific Heat": "0.158 J/(g·K)",
                "Group": 18,
                "Period": 5,
                "Category": "Noble Gas",
                "Discovered By": "William Ramsay and Morris Travers",
                "Year of Discovery": 1898
            },
            "Cs": {
                "Name": "Cesium",
                "Atomic Number": 55,
                "Atomic Mass": "132.90545 u",
                "Density": "1.873 g/cm³",
                "Melting Point": "28.5 °C",
                "Boiling Point": "671 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 6s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "3.894 eV",
                "Specific Heat": "0.242 J/(g·K)",
                "Group": 1,
                "Period": 6,
                "Category": "Alkali Metal",
                "Discovered By": "Robert Bunsen and Gustav Kirchhoff",
                "Year of Discovery": 1860
            },
            "Ba": {
                "Name": "Barium",
                "Atomic Number": 56,
                "Atomic Mass": "137.327 u",
                "Density": "3.5 g/cm³",
                "Melting Point": "727 °C",
                "Boiling Point": "1897 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 6s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "5.212 eV",
                "Specific Heat": "0.204 J/(g·K)",
                "Group": 2,
                "Period": 6,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Humphry Davy",
                "Year of Discovery": 1808
            },
            "La": {
                "Name": "Lanthanum",
                "Atomic Number": 57,
                "Atomic Mass": "138.9055 u",
                "Density": "6.14525 g/cm³",
                "Melting Point": "918 °C",
                "Boiling Point": "3464 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 5d¹ 6s²",
                "Ground State": "²D₃/₂",
                "Ionization Energy": "5.577 eV",
                "Specific Heat": "0.195 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1839
            },
            "Ce": {
                "Name": "Cerium",
                "Atomic Number": 58,
                "Atomic Mass": "140.116 u",
                "Density": "6.77025 g/cm³",
                "Melting Point": "798 °C",
                "Boiling Point": "3443 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹ 5d¹ 6s²",
                "Ground State": "¹G₄",
                "Ionization Energy": "5.539 eV",
                "Specific Heat": "0.192 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Martin Heinrich Klaproth, Jöns Jakob Berzelius, Wilhelm Hisinger",
                "Year of Discovery": 1803
            },
            "Pr": {
                "Name": "Praseodymium",
                "Atomic Number": 59,
                "Atomic Mass": "140.90765 u",
                "Density": "6.773 g/cm³",
                "Melting Point": "931 °C",
                "Boiling Point": "3520 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f³ 6s²",
                "Ground State": "⁴I₉/₂",
                "Ionization Energy": "5.473 eV",
                "Specific Heat": "0.193 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Auer von Welsbach",
                "Year of Discovery": 1885
            },
            "Nd": {
                "Name": "Neodymium",
                "Atomic Number": 60,
                "Atomic Mass": "144.24 u",
                "Density": "7.00825 g/cm³",
                "Melting Point": "1021 °C",
                "Boiling Point": "3074 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁴ 6s²",
                "Ground State": "⁵I₄",
                "Ionization Energy": "5.525 eV",
                "Specific Heat": "0.190 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Auer von Welsbach",
                "Year of Discovery": 1885
            },
            "Pm": {
                "Name": "Promethium",
                "Atomic Number": 61,
                "Atomic Mass": "(145) u",
                "Density": "7.26425 g/cm³",
                "Melting Point": "1042 °C",
                "Boiling Point": "3000 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁵ 6s²",
                "Ground State": "⁶H₅/₂",
                "Ionization Energy": "5.582 eV",
                "Specific Heat": "—",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jacob A. Marinsky, Lawrence E. Glendenin, Charles D. Coryell",
                "Year of Discovery": 1945
            },
            "Sm": {
                "Name": "Samarium",
                "Atomic Number": 62,
                "Atomic Mass": "150.36 u",
                "Density": "7.52025 g/cm³",
                "Melting Point": "1074 °C",
                "Boiling Point": "1794 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁶ 6s²",
                "Ground State": "⁷F₀",
                "Ionization Energy": "5.644 eV",
                "Specific Heat": "0.197 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Paul Émile Lecoq de Boisbaudran",
                "Year of Discovery": 1879
            },
            "Eu": {
                "Name": "Europium",
                "Atomic Number": 63,
                "Atomic Mass": "151.964 u",
                "Density": "5.24425 g/cm³",
                "Melting Point": "822 °C",
                "Boiling Point": "1529 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁷ 6s²",
                "Ground State": "⁸S₇/₂",
                "Ionization Energy": "5.670 eV",
                "Specific Heat": "0.182 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Eugène-Anatole Demarçay",
                "Year of Discovery": 1901
            },
            "Gd": {
                "Name": "Gadolinium",
                "Atomic Number": 64,
                "Atomic Mass": "157.25 u",
                "Density": "7.90125 g/cm³",
                "Melting Point": "1313 °C",
                "Boiling Point": "3273 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁷ 5d¹ 6s²",
                "Ground State": "⁹D₂",
                "Ionization Energy": "6.150 eV",
                "Specific Heat": "0.236 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jean Charles Galissard de Marignac",
                "Year of Discovery": 1880
            },
            "Tb": {
                "Name": "Terbium",
                "Atomic Number": 65,
                "Atomic Mass": "158.92534 u",
                "Density": "8.230 g/cm³",
                "Melting Point": "1356 °C",
                "Boiling Point": "3230 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁹ 6s²",
                "Ground State": "⁶H₁₅/₂",
                "Ionization Energy": "5.864 eV",
                "Specific Heat": "0.182 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1843
            },
            "Dy": {
                "Name": "Dysprosium",
                "Atomic Number": 66,
                "Atomic Mass": "162.50 u",
                "Density": "8.55125 g/cm³",
                "Melting Point": "1412 °C",
                "Boiling Point": "2567 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁰ 6s²",
                "Ground State": "⁵I₈",
                "Ionization Energy": "5.939 eV",
                "Specific Heat": "0.170 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Paul Émile Lecoq de Boisbaudran",
                "Year of Discovery": 1886
            },
            "Ho": {
                "Name": "Holmium",
                "Atomic Number": 67,
                "Atomic Mass": "164.93032 u",
                "Density": "8.79525 g/cm³",
                "Melting Point": "1474 °C",
                "Boiling Point": "2700 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹¹ 6s²",
                "Ground State": "⁴I₁₅/₂",
                "Ionization Energy": "6.022 eV",
                "Specific Heat": "0.165 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Marc Delafontaine, Jacques-Louis Soret",
                "Year of Discovery": 1878
            },
            "Er": {
                "Name": "Erbium",
                "Atomic Number": 68,
                "Atomic Mass": "167.26 u",
                "Density": "9.06625 g/cm³",
                "Melting Point": "1529 °C",
                "Boiling Point": "2868 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹² 6s²",
                "Ground State": "³H₆",
                "Ionization Energy": "6.108 eV",
                "Specific Heat": "0.168 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1843
            },
            "Tm": {
                "Name": "Thulium",
                "Atomic Number": 69,
                "Atomic Mass": "168.93421 u",
                "Density": "9.32125 g/cm³",
                "Melting Point": "1545 °C",
                "Boiling Point": "1950 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹³ 6s²",
                "Ground State": "²F₇/₂",
                "Ionization Energy": "6.184 eV",
                "Specific Heat": "0.160 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Per Teodor Cleve",
                "Year of Discovery": 1879
            },
            "Yb": {
                "Name": "Ytterbium",
                "Atomic Number": 70,
                "Atomic Mass": "173.04 u",
                "Density": "6.966 g/cm³",
                "Melting Point": "819 °C",
                "Boiling Point": "1196 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 6s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "6.254 eV",
                "Specific Heat": "0.155 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jean Charles Galissard de Marignac",
                "Year of Discovery": 1878
            },
            "Lu": {
                "Name": "Lutetium",
                "Atomic Number": 71,
                "Atomic Mass": "174.967 u",
                "Density": "9.84125 g/cm³",
                "Melting Point": "1663 °C",
                "Boiling Point": "3402 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹ 6s²",
                "Ground State": "²D₃/₂",
                "Ionization Energy": "5.426 eV",
                "Specific Heat": "0.154 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Georges Urbain",
                "Year of Discovery": 1907
            },
            "Hf": {
                "Name": "Hafnium",
                "Atomic Number": 72,
                "Atomic Mass": "178.49 u",
                "Density": "13.31 g/cm³",
                "Melting Point": "2233 °C",
                "Boiling Point": "4603 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d² 6s²",
                "Ground State": "³F₂",
                "Ionization Energy": "6.825 eV",
                "Specific Heat": "0.144 J/(g·K)",
                "Group": 4,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Dirk Coster and George de Hevesy",
                "Year of Discovery": 1923
            },
            "Ta": {
                "Name": "Tantalum",
                "Atomic Number": 73,
                "Atomic Mass": "180.9479 u",
                "Density": "16.654 g/cm³",
                "Melting Point": "3017 °C",
                "Boiling Point": "5458 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d³ 6s²",
                "Ground State": "⁴F₃/₂",
                "Ionization Energy": "7.550 eV",
                "Specific Heat": "0.140 J/(g·K)",
                "Group": 5,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Anders Gustaf Ekeberg",
                "Year of Discovery": 1802
            },
            "W": {
                "Name": "Tungsten",
                "Atomic Number": 74,
                "Atomic Mass": "183.84 u",
                "Density": "19.3 g/cm³",
                "Melting Point": "3422 °C",
                "Boiling Point": "5555 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁴ 6s²",
                "Ground State": "⁵D₀",
                "Ionization Energy": "7.864 eV",
                "Specific Heat": "0.132 J/(g·K)",
                "Group": 6,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Fausto and Juan José de Elhuyar",
                "Year of Discovery": 1783
            },
            "Re": {
                "Name": "Rhenium",
                "Atomic Number": 75,
                "Atomic Mass": "186.207 u",
                "Density": "21.02 g/cm³",
                "Melting Point": "3186 °C",
                "Boiling Point": "5596 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁵ 6s²",
                "Ground State": "⁶S₅/₂",
                "Ionization Energy": "7.834 eV",
                "Specific Heat": "0.137 J/(g·K)",
                "Group": 7,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Masataka Ogawa, Walter Noddack, Ida Tacke, Otto Berg",
                "Year of Discovery": 1925
            },
            "Os": {
                "Name": "Osmium",
                "Atomic Number": 76,
                "Atomic Mass": "190.23 u",
                "Density": "22.57 g/cm³",
                "Melting Point": "3033 °C",
                "Boiling Point": "5012 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁶ 6s²",
                "Ground State": "⁵D₄",
                "Ionization Energy": "8.438 eV",
                "Specific Heat": "0.130 J/(g·K)",
                "Group": 8,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Smithson Tennant",
                "Year of Discovery": 1803
            },
            "Ir": {
                "Name": "Iridium",
                "Atomic Number": 77,
                "Atomic Mass": "192.217 u",
                "Density": "22.4217 g/cm³",
                "Melting Point": "2446 °C",
                "Boiling Point": "4428 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁷ 6s²",
                "Ground State": "⁴F₉/₂",
                "Ionization Energy": "8.967 eV",
                "Specific Heat": "0.131 J/(g·K)",
                "Group": 9,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Smithson Tennant",
                "Year of Discovery": 1803
            },
            "Pt": {
                "Name": "Platinum",
                "Atomic Number": 78,
                "Atomic Mass": "195.078 u",
                "Density": "21.45 g/cm³",
                "Melting Point": "1768.4 °C",
                "Boiling Point": "3825 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁹ 6s¹",
                "Ground State": "³D₃",
                "Ionization Energy": "8.959 eV",
                "Specific Heat": "0.133 J/(g·K)",
                "Group": 10,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Antonio de Ulloa",
                "Year of Discovery": 1735
            },
            "Au": {
                "Name": "Gold",
                "Atomic Number": 79,
                "Atomic Mass": "196.96655 u",
                "Density": "19.3 g/cm³",
                "Melting Point": "1064.18 °C",
                "Boiling Point": "2856 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "9.226 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 11,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Hg": {
                "Name": "Mercury",
                "Atomic Number": 80,
                "Atomic Mass": "200.59 u",
                "Density": "13.546 g/cm³",
                "Melting Point": "-38.83 °C",
                "Boiling Point": "356.73 °C",
                "State at 20°C": "Liquid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "10.438 eV",
                "Specific Heat": "0.140 J/(g·K)",
                "Group": 12,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Tl": {
                "Name": "Thallium",
                "Atomic Number": 81,
                "Atomic Mass": "204.3833 u",
                "Density": "11.85 g/cm³",
                "Melting Point": "304 °C",
                "Boiling Point": "1473 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p¹",
                "Ground State": "²P₁/₂",
                "Ionization Energy": "6.108 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 13,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "William Crookes",
                "Year of Discovery": 1861
            },
            "Pb": {
                "Name": "Lead",
                "Atomic Number": 82,
                "Atomic Mass": "207.2 u",
                "Density": "11.35 g/cm³",
                "Melting Point": "327.46 °C",
                "Boiling Point": "1749 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p²",
                "Ground State": "³P₀",
                "Ionization Energy": "7.417 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 14,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Bi": {
                "Name": "Bismuth",
                "Atomic Number": 83,
                "Atomic Mass": "208.98038 u",
                "Density": "9.747 g/cm³",
                "Melting Point": "271.40 °C",
                "Boiling Point": "1564 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p³",
                "Ground State": "⁴S₃/₂",
                "Ionization Energy": "7.286 eV",
                "Specific Heat": "0.122 J/(g·K)",
                "Group": 15,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Po": {
                "Name": "Polonium",
                "Atomic Number": 84,
                "Atomic Mass": "(209) u",
                "Density": "9.32 g/cm³",
                "Melting Point": "254 °C",
                "Boiling Point": "962 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁴",
                "Ground State": "³P₂",
                "Ionization Energy": "8.417 eV",
                "Specific Heat": "—",
                "Group": 16,
                "Period": 6,
                "Category": "Metalloid",
                "Discovered By": "Marie and Pierre Curie",
                "Year of Discovery": 1898
            },
            "At": {
                "Name": "Astatine",
                "Atomic Number": 85,
                "Atomic Mass": "(210) u",
                "Density": "—",
                "Melting Point": "302 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁵",
                "Ground State": "²P₃/₂",
                "Ionization Energy": "—",
                "Specific Heat": "—",
                "Group": 17,
                "Period": 6,
                "Category": "Halogen",
                "Discovered By": "Dale R. Corson, Kenneth Ross MacKenzie, Emilio Segrè",
                "Year of Discovery": 1940
            },
            "Rn": {
                "Name": "Radon",
                "Atomic Number": 86,
                "Atomic Mass": "(222) u",
                "Density": "—",
                "Melting Point": "-71 °C",
                "Boiling Point": "-61.7 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁶",
                "Ground State": "¹S₀",
                "Ionization Energy": "10.748 eV",
                "Specific Heat": "0.094 J/(g·K)",
                "Group": 18,
                "Period": 6,
                "Category": "Noble Gas",
                "Discovered By": "Friedrich Ernst Dorn",
                "Year of Discovery": 1900
            },
            "Fr": {
                "Name": "Francium",
                "Atomic Number": 87,
                "Atomic Mass": "(223) u",
                "Density": "—",
                "Melting Point": "27 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 7s¹",
                "Ground State": "²S₁/₂",
                "Ionization Energy": "4.073 eV",
                "Specific Heat": "—",
                "Group": 1,
                "Period": 7,
                "Category": "Alkali Metal",
                "Discovered By": "Marguerite Perey",
                "Year of Discovery": 1939
            },
            "Ra": {
                "Name": "Radium",
                "Atomic Number": 88,
                "Atomic Mass": "(226) u",
                "Density": "—",
                "Melting Point": "700 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 7s²",
                "Ground State": "¹S₀",
                "Ionization Energy": "5.278 eV",
                "Specific Heat": "—",
                "Group": 2,
                "Period": 7,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Marie and Pierre Curie",
                "Year of Discovery": 1898
            },
            "Ac": {
                "Name": "Actinium",
                "Atomic Number": 89,
                "Atomic Mass": "(227) u",
                "Density": "—",
                "Melting Point": "1051 °C",
                "Boiling Point": "3198 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 6d¹ 7s²",
                "Ground State": "²D₃/₂",
                "Ionization Energy": "5.17 eV",
                "Specific Heat": "0.120 J/(g·K)",
                "Group": 3,
                "Period": 7,
                "Category": "Actinide",
                "Discovered By": "André-Louis Debierne, Friedrich Oskar Giesel",
                "Year of Discovery": 1899
            },
            "Th": {
                "Name": "Thorium",
                "Atomic Number": 90,
                "Atomic Mass": "232.0381 u",
                "Density": "11.72 g/cm³",
                "Melting Point": "1750 °C",
                "Boiling Point": "4788 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 6d² 7s²",
                "Ground State": "³F₂",
                "Ionization Energy": "6.307 eV",
                "Specific Heat": "0.113 J/(g·K)",
                "Group": 3,
                "Period": 7,
                "Category": "Actinide",
                "Discovered By": "Jöns Jakob Berzelius",
                "Year of Discovery": 1829
            }

        }

        # Add more elements as needed...

        # Default properties for elements not explicitly defined
        default_properties = {
            "Name": element_symbol,
            "Atomic Number": "N/A",
            "Atomic Mass": "N/A",
            "Density": "N/A",
            "Melting Point": "N/A",
            "Boiling Point": "N/A",
            "State at 20°C": "N/A",
            "Electron Configuration": "N/A",
            "Ground State": "N/A",
            "Electronegativity": "N/A",
            "Ionization Energy": "N/A",
            "Specific Heat": "N/A",
            "Group": "N/A",
            "Period": "N/A",
            "Category": "N/A",
            "Common Core Levels": "N/A",
            "Most Intense Line": "N/A",
            "Typical FWHM": "N/A",
            "Chemical Shift Range": "N/A"
        }

        # Return known properties or default set
        return element_properties.get(element_symbol, default_properties)




if __name__ == "__main__":
    app = PeriodicTableXPS()
    app.mainloop()