# libraries/grid_operations.py

import wx

def populate_results_grid(window):
    if 'Results' in window.Data and 'Peak' in window.Data['Results']:
        results = window.Data['Results']['Peak']

        # Clear existing data in the grid
        window.results_grid.ClearGrid()

        # Resize the grid if necessary
        num_rows = len(results)
        num_cols = 26  # Based on your Export.py structure
        if window.results_grid.GetNumberRows() < num_rows:
            window.results_grid.AppendRows(num_rows - window.results_grid.GetNumberRows())
        if window.results_grid.GetNumberCols() < num_cols:
            window.results_grid.AppendCols(num_cols - window.results_grid.GetNumberCols())

        # Populate the grid
        for row, (peak_label, peak_data) in enumerate(results.items()):
            window.results_grid.SetCellValue(row, 0, peak_data['Name'])
            window.results_grid.SetCellValue(row, 1, str(peak_data['Position']))
            window.results_grid.SetCellValue(row, 2, str(peak_data['Height']))
            window.results_grid.SetCellValue(row, 3, str(peak_data['FWHM']))
            window.results_grid.SetCellValue(row, 4, str(peak_data['L/G']))
            window.results_grid.SetCellValue(row, 5, f"{peak_data['Area']:.2f}")
            window.results_grid.SetCellValue(row, 6, f"{peak_data['at. %']:.2f}")

            # Set up checkbox and its state
            checkbox_state = peak_data.get('Checkbox', '0')
            window.results_grid.SetCellRenderer(row, 7, wx.grid.GridCellBoolRenderer())
            window.results_grid.SetCellValue(row, 7, checkbox_state)


            window.results_grid.SetCellValue(row, 8, f"{peak_data['RSF']:.2f}")
            window.results_grid.SetCellValue(row, 9, f"{peak_data['TXFN']:.2f}")
            window.results_grid.SetCellValue(row, 10, str(peak_data['ECF']))
            window.results_grid.SetCellValue(row, 11, str(peak_data.get('Instrument', 'Al1486')))
            window.results_grid.SetCellValue(row, 12, peak_data['Fitting Model'])
            window.results_grid.SetCellValue(row, 13, f"{peak_data['Rel. Area']:.2f}")
            window.results_grid.SetCellValue(row, 14, str(peak_data['Sigma']))
            window.results_grid.SetCellValue(row, 15, str(peak_data['Gamma']))
            window.results_grid.SetCellValue(row, 16, peak_data.get('Bkg Type', ''))  # Bkg Type
            window.results_grid.SetCellValue(row, 17, str(peak_data['Bkg Low']))
            window.results_grid.SetCellValue(row, 18, str(peak_data['Bkg High']))
            window.results_grid.SetCellValue(row, 19, str(peak_data.get('Bkg Offset Low', '')))
            window.results_grid.SetCellValue(row, 20, str(peak_data.get('Bkg Offset High', '')))
            window.results_grid.SetCellValue(row, 21, peak_data['Sheetname'])
            window.results_grid.SetCellValue(row, 22, peak_data['Pos. Constraint'])
            window.results_grid.SetCellValue(row, 23, peak_data['Height Constraint'])
            window.results_grid.SetCellValue(row, 24, peak_data['FWHM Constraint'])
            window.results_grid.SetCellValue(row, 25, peak_data['L/G Constraint'])
            window.results_grid.SetCellValue(row, 26, peak_data['Area Constraint'])
            window.results_grid.SetCellValue(row, 27, peak_data['Sigma Constraint'])
            window.results_grid.SetCellValue(row, 28, peak_data['Gamma Constraint'])

        # Bind events
        # window.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, window.on_checkbox_update)
        window.results_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, window.on_cell_changed)

        # Refresh the grid
        window.results_grid.ForceRefresh()
        window.results_grid.Refresh()

    else:
        print("No results data found in window.Data")

    # Calculate atomic percentages for checked elements
    window.update_atomic_percentages()