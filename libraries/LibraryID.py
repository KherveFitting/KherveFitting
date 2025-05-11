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


class PeriodicTableXPS(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("My Dream NIST Library")
        self.geometry("740x900")
        # Fix the width but allow height to vary
        self.minsize(740, 500)  # Minimum width set to 740, minimum height can be 0
        self.maxsize(740, 10000)  # Maximum width fixed at 740, height can be very large

        # self.SetMinSize((262, 380))
        # self.SetMaxSize((262, 380))

        # Set up styles and fonts
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(size=9)
        self.heading_font = tkfont.Font(family="Helvetica", size=10, weight="bold")

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
            btn = tk.Button(pt_grid, text=element, width=4, height=2,
                            bg=color, activebackground=self.darken_color(color),
                            command=lambda e=element: self.select_element(e))

            # Check if element is in our dataset
            if element in self.elements:
                btn.config(relief=tk.RAISED)
            else:
                btn.config(relief=tk.SUNKEN, state=tk.DISABLED)

            btn.grid(row=row, column=col, padx=1, pady=1, sticky='nsew')

        # Add labels for lanthanides and actinides
        tk.Label(pt_grid, text="*", font=("Helvetica", 14)).grid(row=6, column=2)
        tk.Label(pt_grid, text="**", font=("Helvetica", 14)).grid(row=7, column=2)

        # # Create a legend frame
        # legend_frame = tk.Frame(self.periodic_frame, bg="#f0f0f0", pady=5)
        # legend_frame.pack(fill=tk.X)
        #
        # # Add color legend
        # legend_categories = [
        #     ('Alkali Metals', colors['alkali_metal']),
        #     ('Alkaline Earth Metals', colors['alkaline_earth']),
        #     ('Transition Metals', colors['transition_metal']),
        #     ('Post-Transition Metals', colors['post_transition']),
        #     ('Metalloids', colors['metalloid']),
        #     ('Nonmetals', colors['nonmetal']),
        #     ('Halogens', colors['halogen']),
        #     ('Noble Gases', colors['noble_gas']),
        #     ('Lanthanides*', colors['lanthanide']),
        #     ('Actinides**', colors['actinide'])
        # ]
        #
        # # Create legend items
        # for i, (name, color) in enumerate(legend_categories):
        #     row = i // 5
        #     col = i % 5
        #     frame = tk.Frame(legend_frame, padx=5)
        #     frame.grid(row=row, column=col, padx=10, pady=2)
        #
        #     color_box = tk.Label(frame, bg=color, width=2, height=1)
        #     color_box.pack(side=tk.LEFT, padx=(0, 5))
        #
        #     tk.Label(frame, text=name, bg="#f0f0f0", font=("Helvetica", 8)).pack(side=tk.LEFT)

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

        # Reset button
        self.reset_btn = tk.Button(search_frame, text="Reset All", command=self.reset_all)
        self.reset_btn.grid(row=0, column=3, padx=10, pady=2, sticky='w')

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
        columns = ("Element", "Line", "BE (eV)", "Formula", "Name", "Journal")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Treeview")

        # Define column widths and headings
        widths = {"Element": 10, "Line": 10, "BE (eV)": 10, "Formula": 10, "Name": 10, "Journal": 200}
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=widths[col])

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

    def get_filtered_data(self):
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




if __name__ == "__main__":
    app = PeriodicTableXPS()
    app.mainloop()