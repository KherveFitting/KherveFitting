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

class PeriodicTableXPS(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("My KherveNIST Library: How I wish NIST would look like")
        self.geometry("620x660")
        # Fix the width but allow height to vary
        self.minsize(620, 660)  # Minimum width set to 740, minimum height can be 0
        self.maxsize(620, 10000)  # Maximum width fixed at 740, height can be very large

        # self.SetMinSize((262, 380))
        # self.SetMaxSize((262, 380))

        # Set up styles and fonts
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(size=9)
        self.heading_font = tkfont.Font(family="Helvetica", size=10, weight="bold")

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

    def create_menu(self):
        """Create the application menu bar"""
        menubar = Menu(self)
        self.config(menu=menubar)

        # File menu
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.quit)

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
                            font=("Helvetica", 10))
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

    def load_data(self):
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

        if not data_found:
            tk.messagebox.showerror("Error", f"Failed to load data: NIST_BE.xlsx not found in expected locations")
            self.quit()

    def create_frames(self):
        """Create main layout frames"""
        # Top frame for periodic table
        self.periodic_frame = tk.Frame(self, bg="#f0f0f0", padx=10, pady=10)
        self.periodic_frame.pack(fill=tk.BOTH, expand=False)

        # Remove the "Periodic Table" label
        # tk.Label(self.periodic_frame, text="Periodic Table", font=self.heading_font,
        #         bg="#f0f0f0").pack(side=tk.TOP, pady=(0, 10))

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

        # Create buttons for each element
        for element, (row, col) in self.element_positions.items():
            # Get color based on element category
            category = element_categories.get(element, 'unknown')
            color = colors.get(category, colors['unknown'])

            # Create button
            btn = tk.Button(pt_grid, text=element, width=3, height=1,
                            bg=color, activebackground=self.darken_color(color),
                            command=lambda e=element: self.select_element(e))

            # Check if element is in our dataset
            if element in self.elements:
                btn.config(relief=tk.RAISED)
            else:
                btn.config(relief=tk.SUNKEN, state=tk.DISABLED)

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
        tk.Label(left_frame, text="Selected Element:", bg="#e0e0e0").grid(row=0, column=0, sticky='w', pady=2)
        self.element_label = tk.Label(left_frame, text="None", width=6, relief=tk.SUNKEN, bg="white")
        self.element_label.grid(row=0, column=1, sticky='w', padx=5, pady=2)

        # Line selection dropdown
        tk.Label(left_frame, text="XPS Line:", bg="#e0e0e0").grid(row=1, column=0, sticky='w', pady=2)
        self.line_var = tk.StringVar()
        self.line_dropdown = ttk.Combobox(left_frame, textvariable=self.line_var, width=15, state="readonly")
        self.line_dropdown['values'] = ['All Lines'] + list(self.lines)
        self.line_dropdown.current(0)
        self.line_dropdown.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        self.line_dropdown.bind('<<ComboboxSelected>>', self.update_results)

        # Create a frame for search boxes
        search_frame = tk.Frame(right_frame, bg="#e0e0e0")
        search_frame.pack(fill=tk.X)

        # Formula search
        tk.Label(search_frame, text="Search Formula:", bg="#e0e0e0").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.on_search_change)
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=20)
        self.search_entry.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        # Name search
        tk.Label(search_frame, text="Search Name:", bg="#e0e0e0").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.name_search_var = tk.StringVar()
        self.name_search_var.trace("w", self.on_search_change)
        self.name_search_entry = tk.Entry(search_frame, textvariable=self.name_search_var, width=20)
        self.name_search_entry.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # # Reset button
        # self.reset_btn = tk.Button(search_frame, text="Reset All", command=self.reset_all)
        # self.reset_btn.grid(row=0, column=3, padx=10, pady=2, sticky='w')

        # Add element properties button
        self.properties_btn = tk.Button(search_frame, text="Properties", command=self.show_element_properties)
        self.properties_btn.grid(row=0, column=3, sticky='w', padx=10, pady=2)

        # Plot button - ADD THIS NEW BUTTON
        self.plot_btn = tk.Button(search_frame, text="Plot Results", command=self.plot_results)
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
        style.configure("Treeview", rowheight=20)
        style.configure("Treeview.Heading", font=('Helvetica', 10))  # Smaller heading font

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

        # Add right-click context menu for the Journal column
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

            # Create a single frame to hold the entire table
            table_frame = tk.Frame(scrollable_frame, bd=1, relief=tk.SOLID)
            table_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

            # Define column widths
            headers = ["Line", "Avg BE (eV)", "Min BE (eV)", "Max BE (eV)", "Count"]
            column_widths = [8, 10, 10, 10, 7]  # Width in characters for each column

            # Add header row
            for col, (header, width) in enumerate(zip(headers, column_widths)):
                header_label = tk.Label(table_frame, text=header,
                                        font=("Helvetica", 10, "bold"), bg="#e8e8e8",
                                        padx=5, pady=2, relief=tk.RIDGE, width=width,
                                        borderwidth=1)
                header_label.grid(row=0, column=col, sticky="nsew")

            # Add data rows
            for i, (_, line_data) in enumerate(lines_summary.iterrows()):
                # Line column
                tk.Label(table_frame, text=line_data['Line'],
                         font=("Helvetica", 10), padx=5, pady=2,
                         relief=tk.RIDGE, width=column_widths[0],
                         borderwidth=1).grid(row=i + 1, column=0, sticky="nsew")

                # Avg BE column
                tk.Label(table_frame, text=f"{line_data['mean']:.2f}",
                         font=("Helvetica", 10), padx=5, pady=2,
                         relief=tk.RIDGE, width=column_widths[1],
                         borderwidth=1).grid(row=i + 1, column=1, sticky="nsew")

                # Min BE column
                tk.Label(table_frame, text=f"{line_data['min']:.2f}",
                         font=("Helvetica", 10), padx=5, pady=2,
                         relief=tk.RIDGE, width=column_widths[2],
                         borderwidth=1).grid(row=i + 1, column=2, sticky="nsew")

                # Max BE column
                tk.Label(table_frame, text=f"{line_data['max']:.2f}",
                         font=("Helvetica", 10), padx=5, pady=2,
                         relief=tk.RIDGE, width=column_widths[3],
                         borderwidth=1).grid(row=i + 1, column=3, sticky="nsew")

                # Count column
                tk.Label(table_frame, text=str(line_data['count']),
                         font=("Helvetica", 10), padx=5, pady=2,
                         relief=tk.RIDGE, width=column_widths[4],
                         borderwidth=1).grid(row=i + 1, column=4, sticky="nsew")

            # Configure all columns to expand properly
            for i in range(len(headers)):
                table_frame.columnconfigure(i, weight=1)

            row += 1  # Move to the next row after the table
        else:
            no_data_label = tk.Label(scrollable_frame, text="No XPS data available for this element",
                                     font=("Helvetica", 10, "italic"), padx=5, pady=5)
            no_data_label.grid(row=row, column=0, columnspan=2, sticky="ew")
            row += 1



    def get_element_properties(self, element_symbol):
        """Get properties for a specific element"""
        # This is a database of element properties
        # In a real application, this would come from a database or API
        # For demonstration, I'll include data for a few common elements
        element_properties = {
            "H": {
                "Name": "Hydrogen",
                "Atomic Number": 1,
                "Atomic Mass": "1.008 u",
                "Density": "0.00008988 g/cm³",
                "Melting Point": "-259.16 °C",
                "Boiling Point": "-252.87 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s¹",
                "Electronegativity": 2.20,
                "Atomic Radius": "25 pm",
                "Ionization Energy": "13.598 eV",
                "Group": 1,
                "Period": 1,
                "Category": "Nonmetal",
                "Discovered By": "Henry Cavendish",
                "Year of Discovery": 1766,
                "Common Core Levels": "1s",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.2 eV",
                "Chemical Shift Range": "0-3 eV"
            },
            "C": {
                "Name": "Carbon",
                "Atomic Number": 6,
                "Atomic Mass": "12.011 u",
                "Density": "2.267 g/cm³",
                "Melting Point": "3550 °C",
                "Boiling Point": "4027 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s² 2p²",
                "Electronegativity": 2.55,
                "Atomic Radius": "70 pm",
                "Ionization Energy": "11.260 eV",
                "Group": 14,
                "Period": 2,
                "Category": "Nonmetal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric",
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.8-1.2 eV",
                "Chemical Shift Range": "0-10 eV"
            },
            "O": {
                "Name": "Oxygen",
                "Atomic Number": 8,
                "Atomic Mass": "15.999 u",
                "Density": "0.001429 g/cm³",
                "Melting Point": "-218.79 °C",
                "Boiling Point": "-182.95 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁴",
                "Electronegativity": 3.44,
                "Atomic Radius": "60 pm",
                "Ionization Energy": "13.618 eV",
                "Group": 16,
                "Period": 2,
                "Category": "Nonmetal",
                "Discovered By": "Joseph Priestley, Carl Wilhelm Scheele",
                "Year of Discovery": 1774,
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.3 eV",
                "Chemical Shift Range": "0-8 eV"
            }
        }

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
            "Electronegativity": "N/A",
            "Atomic Radius": "N/A",
            "Ionization Energy": "N/A",
            "Group": "N/A",
            "Period": "N/A",
            "Category": "N/A",
            "Discovered By": "N/A",
            "Year of Discovery": "N/A",
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