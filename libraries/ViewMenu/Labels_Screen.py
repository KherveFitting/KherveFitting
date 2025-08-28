import wx
import wx.grid
from matplotlib.patches import Rectangle
from matplotlib.transforms import Affine2D
import numpy as np

class LabelWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Labels Manager", style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetSize((330, 400))

        # Position on right side of main frame
        main_pos = self.parent.GetPosition()
        main_size = self.parent.GetSize()
        window_size = self.GetSize()
        x = main_pos.x + main_size.width - 550
        y = main_pos.y + (main_size.height - window_size.height) // 2
        self.SetPosition((x, y))

        panel = wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(220, 230, 220))
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Create grid for labels
        self.labels_grid = wx.grid.Grid(panel)
        self.labels_grid.CreateGrid(0, 4)
        self.labels_grid.SetColLabelValue(0, "Name")
        self.labels_grid.SetColLabelValue(1, "Pos X")
        self.labels_grid.SetColLabelValue(2, "Pos Y")
        self.labels_grid.SetColLabelValue(3, "Rotation")

        # Set column widths
        self.labels_grid.SetColLabelSize(35)
        self.labels_grid.SetRowLabelSize(25)
        self.labels_grid.SetDefaultColSize(120)
        self.labels_grid.SetColSize(0, 80)
        self.labels_grid.SetColSize(1, 60)
        self.labels_grid.SetColSize(2, 60)
        self.labels_grid.SetColSize(3, 60)


        # Enable editing for all columns
        self.labels_grid.EnableEditing(True)

        # Set selection mode to allow both cell and row selection
        self.labels_grid.SetSelectionMode(wx.grid.Grid.GridSelectRows)

        # Bind events for cell changes and key events
        self.labels_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_cell_changed)
        self.labels_grid.Bind(wx.EVT_KEY_DOWN, self.on_key_down)
        # Make sure the grid has focus for key events
        self.labels_grid.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.on_cell_select)

        sizer.Add(self.labels_grid, 1, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        add_btn = wx.Button(panel, label="Add\nLabel")
        add_btn.SetMinSize((60, 40))
        edit_btn = wx.Button(panel, label="Edit\nLabel")
        edit_btn.SetMinSize((60, 40))
        remove_btn = wx.Button(panel, label="Remove\nSelected")
        remove_btn.SetMinSize((60, 40))
        remove_all_btn = wx.Button(panel, label="Remove\nAll")
        remove_all_btn.SetMinSize((80, 40))

        btn_sizer.Add(add_btn, 0, wx.ALL, 5)
        btn_sizer.Add(edit_btn, 0, wx.ALL, 5)
        btn_sizer.Add(remove_btn, 0, wx.ALL, 5)
        btn_sizer.Add(remove_all_btn, 0, wx.ALL, 5)

        add_btn.Bind(wx.EVT_BUTTON, self.on_add)
        edit_btn.Bind(wx.EVT_BUTTON, self.on_edit)
        remove_btn.Bind(wx.EVT_BUTTON, self.on_remove)
        remove_all_btn.Bind(wx.EVT_BUTTON, self.on_remove_all)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)

        buttons_sizer = wx.BoxSizer(wx.HORIZONTAL)
        movement_sizer = wx.BoxSizer(wx.VERTICAL)
        horizontal_sizer = wx.BoxSizer(wx.HORIZONTAL)
        top_sizer = wx.BoxSizer(wx.HORIZONTAL)
        bottom_sizer = wx.BoxSizer(wx.HORIZONTAL)

        up_btn = wx.Button(panel, label="↑")
        down_btn = wx.Button(panel, label="↓")
        left_btn = wx.Button(panel, label="←")
        right_btn = wx.Button(panel, label="→")

        top_sizer.AddSpacer(40)
        top_sizer.Add(up_btn, 0, wx.ALL, 2)
        movement_sizer.Add(top_sizer, 0, wx.ALL, 2)
        horizontal_sizer.Add(left_btn, 0, wx.RIGHT, 2)
        horizontal_sizer.Add(right_btn, 0, wx.LEFT, 2)
        movement_sizer.Add(horizontal_sizer, 0, wx.ALL, 2)
        bottom_sizer.AddSpacer(40)
        bottom_sizer.Add(down_btn, 0, wx.ALL, 2)
        movement_sizer.Add(bottom_sizer, 0, wx.ALL, 2)

        buttons_sizer.AddSpacer(40)
        buttons_sizer.Add(movement_sizer, 0, wx.LEFT | wx.TOP, 10)

        sizer.Add(buttons_sizer, 0, wx.LEFT | wx.TOP, 10)

        up_btn.Bind(wx.EVT_BUTTON, self.move_up)
        down_btn.Bind(wx.EVT_BUTTON, self.move_down)
        left_btn.Bind(wx.EVT_BUTTON, self.move_left)
        right_btn.Bind(wx.EVT_BUTTON, self.move_right)

        # Add these new attributes
        self.selected_text = None
        self.selection_box = None
        self.drag_offset = None
        self.is_dragging = False

        # Store connection IDs for proper cleanup
        self.canvas_click_id = self.parent.canvas.mpl_connect('button_press_event', self.on_canvas_click)
        self.canvas_release_id = self.parent.canvas.mpl_connect('button_release_event', self.on_canvas_release)
        self.canvas_motion_id = self.parent.canvas.mpl_connect('motion_notify_event', self.on_canvas_motion)

        # Bind close event and key events
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.labels_grid.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

        # Bind to parent's sheet combobox to update when sheet changes
        self.parent.sheet_combobox.Bind(wx.EVT_COMBOBOX, self.on_parent_sheet_change)

        panel.SetSizer(sizer)

        self.update_list()

    def on_parent_sheet_change_OLD(self, event):
        """Update labels list when parent sheet selection changes"""
        self.update_list()
        # event.Skip()  # Let the parent handle the event too

    def on_parent_sheet_change(self, event):
        # Check if grid still exists before trying to update
        if hasattr(self, 'labels_grid') and self.labels_grid:
            try:
                self.update_list()
                # event.Skip()  # Let the parent handle the event too
            except RuntimeError:
                # Grid has been destroyed, ignore the update
                pass

    def move_label(self, direction):
        # Get selection from grid
        selected_rows = self.labels_grid.GetSelectedRows()
        current_row = self.labels_grid.GetGridCursorRow()

        selection = None
        if selected_rows:
            selection = selected_rows[0]
        elif current_row >= 0:
            selection = current_row
        else:
            return

        sheet_name = self.parent.sheet_combobox.GetValue()

        if 'Labels' not in self.parent.Data['Core levels'][sheet_name] or selection >= len(self.parent.Data['Core levels'][sheet_name]['Labels']):
            return

        label = self.parent.Data['Core levels'][sheet_name]['Labels'][selection]
        max_y = max(self.parent.y_values)

        if direction == 'left':
            label['x'] += 2
        elif direction == 'right':
            label['x'] -= 2
        elif direction == 'up':
            label['y'] += max_y * 0.02
        elif direction == 'down':
            label['y'] -= max_y * 0.02

        for txt in self.parent.ax.texts[:]:
            txt.remove()

        for label_data in self.parent.Data['Core levels'][sheet_name]['Labels']:
            self.parent.ax.text(
                label_data['x'],
                label_data['y'],
                label_data['text'],
                rotation=label_data.get('rotation', 90),
                fontsize=label_data.get('fontsize', 10),
                fontfamily=label_data.get('fontfamily', 'Arial'),
                va='bottom',
                ha='center'
            )

        self.parent.canvas.draw_idle()
        self.update_list()

        # Maintain the same row selection
        self.labels_grid.SelectRow(selection)
        self.labels_grid.SetGridCursor(selection, 0)

    def move_up(self, event):
        self.move_label('up')

    def move_down(self, event):
        self.move_label('down')

    def move_left(self, event):
        self.move_label('left')

    def move_right(self, event):
        self.move_label('right')

    def update_list(self):
        # Check if grid still exists and is valid
        if not hasattr(self, 'labels_grid') or not self.labels_grid:
            return

        try:
            # Clear existing rows
            if self.labels_grid.GetNumberRows() > 0:
                self.labels_grid.DeleteRows(0, self.labels_grid.GetNumberRows())
        except RuntimeError:
            # Grid has been destroyed, abort update
            return

        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Labels' in self.parent.Data['Core levels'][sheet_name]:
            labels = self.parent.Data['Core levels'][sheet_name]['Labels']

            # Add rows for each label
            for i, label in enumerate(labels):
                self.labels_grid.AppendRows(1)
                self.labels_grid.SetCellValue(i, 0, label['text'])
                self.labels_grid.SetCellValue(i, 1, f"{label['x']:.2f}")
                self.labels_grid.SetCellValue(i, 2, f"{label['y']:.2f}")
                self.labels_grid.SetCellValue(i, 3, f"{label.get('rotation', 90):.2f}")

                # Center align numeric values
                self.labels_grid.SetCellAlignment(i, 1, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                self.labels_grid.SetCellAlignment(i, 2, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                self.labels_grid.SetCellAlignment(i, 3, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # self.labels_grid.AutoSizeColumns()

    def on_canvas_click(self, event):
        # Check if LabelWindow is still open and active
        if not self.IsShown() or self.IsIconized():
            return

        if not event.inaxes:
            self.clear_selection()
            return

        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Labels' not in self.parent.Data['Core levels'][sheet_name]:
            return

        # Get axis ranges for percentage-based calculations
        xlim = self.parent.ax.get_xlim()
        ylim = self.parent.ax.get_ylim()
        x_range = abs(xlim[1] - xlim[0])
        y_range = abs(ylim[1] - ylim[0])

        # Find clicked text
        clicked_text = None
        clicked_index = None

        for i, label_data in enumerate(self.parent.Data['Core levels'][sheet_name]['Labels']):
            # Check if click is near text position
            text_x = label_data['x']
            text_y = label_data['y']

            # Create larger clickable area based on percentage of plot size
            bbox_width = x_range * 0.03  # 3% of x-axis range
            bbox_height = y_range * 0.03  # 5% of y-axis range

            if (abs(event.xdata - text_x) < bbox_width and
                    abs(event.ydata - text_y) < bbox_height):
                clicked_text = label_data
                clicked_index = i
                break

        if clicked_text:
            self.select_text(clicked_text, clicked_index)
            self.drag_offset = (event.xdata - clicked_text['x'], event.ydata - clicked_text['y'])
            self.is_dragging = True
            # Select the entire row and set cursor
            self.labels_grid.SelectRow(clicked_index)
            self.labels_grid.SetGridCursor(clicked_index, 0)
        else:
            self.clear_selection()

    def clear_selection(self):
        if self.selection_box:
            self.selection_box.remove()
            self.selection_box = None
        self.selected_text = None
        self.labels_grid.ClearSelection()
        self.parent.canvas.draw_idle()


    def select_text_OLD(self, label_data, index):
        # Remove previous selection box
        if self.selection_box:
            self.selection_box.remove()

        # Create triangle selection indicator
        x = label_data['x']
        y = label_data['y']

        # Get axis ranges
        xlim = self.parent.ax.get_xlim()
        ylim = self.parent.ax.get_ylim()
        x_range = abs(xlim[1] - xlim[0])
        y_range = abs(ylim[1] - ylim[0])

        # Triangle dimensions - smaller and more consistent
        triangle_width = x_range * 0.03  # 0.8% of x-axis range
        triangle_height = y_range * 0.03  # 1.5% of y-axis range

        # Triangle positioned slightly lower
        triangle_y = y - y_range * 0.005  # 0.5% below text

        from matplotlib.patches import Polygon
        triangle_points = [
            [x, triangle_y],  # Top point
            [x - triangle_width / 2, triangle_y - triangle_height],  # Bottom left
            [x + triangle_width / 2, triangle_y - triangle_height]  # Bottom right
        ]

        self.selection_box = Polygon(
            triangle_points,
            linewidth=1,
            edgecolor='black',
            facecolor=(200 / 255, 245 / 255, 228 / 255),
            linestyle='-'
        )

        self.parent.ax.add_patch(self.selection_box)
        self.selected_text = label_data
        self.parent.canvas.draw_idle()

    def select_text(self, label_data, index):
        # Remove previous selection box
        if self.selection_box:
            self.selection_box.remove()

        # Create triangle selection indicator
        x = label_data['x']
        y = label_data['y']

        # Get axis ranges
        xlim = self.parent.ax.get_xlim()
        ylim = self.parent.ax.get_ylim()
        x_range = abs(xlim[1] - xlim[0])
        y_range = abs(ylim[1] - ylim[0])

        # Triangle dimensions - smaller and more consistent
        triangle_width = x_range * 0.03  # 0.8% of x-axis range
        triangle_height = y_range * 0.03  # 1.5% of y-axis range

        # Triangle positioned slightly lower
        triangle_y = y - y_range * 0.005  # 0.5% below text

        from matplotlib.patches import Polygon
        triangle_points = [
            [x, triangle_y],  # Top point
            [x - triangle_width / 2, triangle_y - triangle_height],  # Bottom left
            [x + triangle_width / 2, triangle_y - triangle_height]  # Bottom right
        ]

        self.selection_box = Polygon(
            triangle_points,
            linewidth=1,
            edgecolor='black',
            facecolor=(200 / 255, 245 / 255, 228 / 255),
            linestyle='-'
        )

        self.parent.ax.add_patch(self.selection_box)
        self.selected_text = label_data
        self.parent.canvas.draw_idle()

    def on_canvas_motion(self, event):
        if not self.is_dragging or not self.selected_text or not event.inaxes:
            return

        # Update text position
        new_x = event.xdata - self.drag_offset[0]
        new_y = event.ydata - self.drag_offset[1]

        self.selected_text['x'] = new_x
        self.selected_text['y'] = new_y

        # Update triangle position
        if self.selection_box:
            xlim = self.parent.ax.get_xlim()
            ylim = self.parent.ax.get_ylim()
            x_range = abs(xlim[1] - xlim[0])
            y_range = abs(ylim[1] - ylim[0])

            # Triangle dimensions - smaller and more consistent
            triangle_width = x_range * 0.03  # 0.8% of x-axis range
            triangle_height = y_range * 0.03  # 1.5% of y-axis range
            triangle_y = new_y - y_range * 0.005  # 0.5% below text

            triangle_points = [
                [new_x, triangle_y],
                [new_x - triangle_width / 2, triangle_y - triangle_height],
                [new_x + triangle_width / 2, triangle_y - triangle_height]
            ]

            self.selection_box.set_xy(triangle_points)

        # Redraw all text
        self.redraw_labels()



    def on_canvas_release(self, event):
        self.is_dragging = False
        if self.selected_text:
            self.update_list()

    def redraw_labels(self):
        # Clear existing text
        for txt in self.parent.ax.texts[:]:
            txt.remove()

        # Redraw all labels
        sheet_name = self.parent.sheet_combobox.GetValue()
        for label_data in self.parent.Data['Core levels'][sheet_name]['Labels']:
            self.parent.ax.text(
                label_data['x'],
                label_data['y'],
                label_data['text'],
                rotation=label_data.get('rotation', 90),
                fontsize=label_data.get('fontsize', 10),
                fontfamily=label_data.get('fontfamily', 'Arial'),
                va='bottom',
                ha='center'
            )

        self.parent.canvas.draw_idle()


    def on_add(self, event):
        # Check if add window is already open
        if hasattr(self, 'add_window') and self.add_window:
            self.add_window.Raise()
            return

        # Create non-modal frame instead of dialog
        self.add_window = wx.Frame(self, title="Add Label", size=(280, 220), style=wx.DEFAULT_FRAME_STYLE |
                                wx.STAY_ON_TOP)

        # Center on main frame
        main_pos = self.parent.GetPosition()
        main_size = self.parent.GetSize()
        window_size = self.add_window.GetSize()
        x = main_pos.x + main_size.width - 350 -300
        y = main_pos.y + (main_size.height - window_size.height) // 2
        self.add_window.SetPosition((x, y))

        panel = wx.Panel(self.add_window)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Create labels with fixed width for alignment
        label_width = 70

        text_sizer = wx.BoxSizer(wx.HORIZONTAL)
        text_label = wx.StaticText(panel, label="Text:", size=(label_width, -1))
        self.add_text_ctrl = wx.TextCtrl(panel, size=(150, -1))
        text_sizer.Add(text_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        text_sizer.Add(self.add_text_ctrl, 1, wx.ALL, 5)
        sizer.Add(text_sizer)

        # Rotation control
        rotation_sizer = wx.BoxSizer(wx.HORIZONTAL)
        rotation_label = wx.StaticText(panel, label="Rotation:", size=(label_width, -1))
        rotation_choices = ['0', '90', '270']
        self.add_rotation_ctrl = wx.Choice(panel, choices=rotation_choices, size=(150, -1))
        self.add_rotation_ctrl.SetSelection(1)  # Default to 90
        rotation_sizer.Add(rotation_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        rotation_sizer.Add(self.add_rotation_ctrl, 1, wx.ALL, 5)
        sizer.Add(rotation_sizer)

        font_size_sizer = wx.BoxSizer(wx.HORIZONTAL)
        font_size_label = wx.StaticText(panel, label="Font Size:", size=(label_width, -1))
        self.add_font_size_ctrl = wx.SpinCtrl(panel, min=1, max=72, initial=10, size=(150, -1))
        font_size_sizer.Add(font_size_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        font_size_sizer.Add(self.add_font_size_ctrl, 1, wx.ALL, 5)
        sizer.Add(font_size_sizer)

        # Font family control
        font_family_sizer = wx.BoxSizer(wx.HORIZONTAL)
        font_family_label = wx.StaticText(panel, label="Font:", size=(label_width, -1))
        self.add_font_families = ['Arial', 'Times New Roman', 'Courier New', 'Helvetica', 'Verdana',
                                  'Calibri', 'Tahoma', 'Georgia', 'Comic Sans MS', 'Impact',
                                  'Trebuchet MS', 'Century Gothic', 'Palatino']
        self.add_font_family_ctrl = wx.Choice(panel, choices=self.add_font_families, size=(150, -1))
        self.add_font_family_ctrl.SetSelection(0)
        font_family_sizer.Add(font_family_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        font_family_sizer.Add(self.add_font_family_ctrl, 1, wx.ALL, 5)
        sizer.Add(font_family_sizer)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        add_btn = wx.Button(panel, label="Add")
        add_btn.SetMinSize((60, 40))
        close_btn = wx.Button(panel, label="Close")
        close_btn.SetMinSize((60, 40))
        btn_sizer.Add(add_btn, 0, wx.ALL, 5)
        btn_sizer.Add(close_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)

        panel.SetSizer(sizer)

        def on_add_click(event):
            sheet_name = self.parent.sheet_combobox.GetValue()
            maxY = max(self.parent.y_values)
            x = (self.parent.ax.get_xlim()[0] + self.parent.ax.get_xlim()[1]) / 2
            y = maxY * 1.1

            if 'Labels' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Labels'] = []

            # Store properties
            self.parent.Data['Core levels'][sheet_name]['Labels'].append({
                'text': self.add_text_ctrl.GetValue(),
                'x': x,
                'y': y,
                'rotation': int(rotation_choices[self.add_rotation_ctrl.GetSelection()]),
                'fontsize': self.add_font_size_ctrl.GetValue(),
                'fontfamily': self.add_font_families[self.add_font_family_ctrl.GetSelection()]
            })

            # Clear all existing text annotations
            for txt in self.parent.ax.texts[:]:
                txt.remove()

            # Apply properties when creating text
            for label_data in self.parent.Data['Core levels'][sheet_name]['Labels']:
                self.parent.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=label_data.get('rotation', 90),
                    fontsize=label_data.get('fontsize', 10),
                    fontfamily=label_data.get('fontfamily', 'Arial'),
                    va='bottom',
                    ha='center'
                )

            self.parent.canvas.draw_idle()
            self.update_list()

            # Clear the text field for next entry
            self.add_text_ctrl.SetValue("")

        def on_close_window(event):
            self.add_window.Destroy()
            self.add_window = None

        add_btn.Bind(wx.EVT_BUTTON, on_add_click)
        close_btn.Bind(wx.EVT_BUTTON, on_close_window)
        self.add_window.Bind(wx.EVT_CLOSE, on_close_window)

        self.add_window.Show()

    def on_edit(self, event):
        selected_rows = self.labels_grid.GetSelectedRows()
        selected_cells = self.labels_grid.GetSelectedCells()

        selection = None
        if selected_rows:
            selection = selected_rows[0]
        elif selected_cells:
            selection = selected_cells[0][0]
        else:
            return

        sheet_name = self.parent.sheet_combobox.GetValue()

        if 'Labels' not in self.parent.Data['Core levels'][sheet_name] or selection >= len(self.parent.Data['Core levels'][sheet_name]['Labels']):
            return

        label = self.parent.Data['Core levels'][sheet_name]['Labels'][selection]

        dlg = wx.Dialog(self, title="Edit Label", size=(250, 250), style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)

        # Center on main frame
        main_pos = self.parent.GetPosition()
        main_size = self.parent.GetSize()
        dialog_size = dlg.GetSize()
        x = main_pos.x + (main_size.width - dialog_size.width) // 2
        y = main_pos.y + (main_size.height - dialog_size.height) // 2
        dlg.SetPosition((x, y))

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Add text label and control
        text_sizer = wx.BoxSizer(wx.HORIZONTAL)
        text_label = wx.StaticText(dlg, label="Label:")
        text_ctrl = wx.TextCtrl(dlg, value=label['text'], size=(120, -1))
        text_sizer.Add(text_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        text_sizer.Add(text_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(text_sizer, 0, wx.EXPAND)

        # Add x control
        x_sizer = wx.BoxSizer(wx.HORIZONTAL)
        x_label = wx.StaticText(dlg, label="X:")
        x_ctrl = wx.SpinCtrlDouble(dlg, value=str(label['x']), min=-10, max=1e4, size=(120, -1))
        x_sizer.Add(x_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        x_sizer.Add(x_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(x_sizer, 0, wx.EXPAND)

        # Add y control
        y_sizer = wx.BoxSizer(wx.HORIZONTAL)
        y_label = wx.StaticText(dlg, label="Y:")
        y_ctrl = wx.SpinCtrlDouble(dlg, value=str(label['y']), min=-10000, max=1e10, size=(120, -1))
        y_sizer.Add(y_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        y_sizer.Add(y_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(y_sizer, 0, wx.EXPAND)

        # Add rotation control
        rotation_sizer = wx.BoxSizer(wx.HORIZONTAL)
        rotation_label = wx.StaticText(dlg, label="Rotation:")
        rotation_choices = ['0', '90', '270']
        rotation_ctrl = wx.Choice(dlg, choices=rotation_choices, size=(120, -1))

        # Set current rotation value
        current_rotation = str(label.get('rotation', 90))
        if current_rotation in rotation_choices:
            rotation_ctrl.SetSelection(rotation_choices.index(current_rotation))
        else:
            rotation_ctrl.SetSelection(1)  # Default to 90

        rotation_sizer.Add(rotation_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        rotation_sizer.Add(rotation_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(rotation_sizer, 0, wx.EXPAND)

        # Add OK and Cancel buttons
        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(dlg, wx.ID_OK)
        cancel_btn = wx.Button(dlg, wx.ID_CANCEL)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        sizer.Add(btn_sizer, 0, wx.ALL | wx.EXPAND, 5)

        dlg.SetSizer(sizer)

        if dlg.ShowModal() == wx.ID_OK:
            # Update label data
            label['text'] = text_ctrl.GetValue()
            label['x'] = x_ctrl.GetValue()
            label['y'] = y_ctrl.GetValue()
            label['rotation'] = int(rotation_choices[rotation_ctrl.GetSelection()])

            # Clear all existing text annotations
            for txt in self.parent.ax.texts[:]:
                txt.remove()

            # Redraw all labels
            for label_data in self.parent.Data['Core levels'][sheet_name]['Labels']:
                self.parent.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=label_data.get('rotation', 90),
                    va='bottom',
                    ha='center'
                )

            self.parent.canvas.draw_idle()
            self.update_list()

        dlg.Destroy()

    def on_remove(self, event):
        selected_rows = self.labels_grid.GetSelectedRows()
        current_row = self.labels_grid.GetGridCursorRow()

        # Determine which row to delete
        if selected_rows:
            selection = selected_rows[0]
        elif current_row >= 0:
            selection = current_row
        else:
            return

        sheet_name = self.parent.sheet_combobox.GetValue()

        if 'Labels' not in self.parent.Data['Core levels'][sheet_name]:
            return

        labels = self.parent.Data['Core levels'][sheet_name]['Labels']
        if selection < len(labels):
            # Remember the row position for reselection
            row_to_select = selection
            total_rows = len(labels)

            # Remove the label
            labels.pop(selection)

            # Clear and redraw plot
            self.parent.clear_and_replot()

            # Redraw remaining labels
            for label_data in labels:
                self.parent.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=label_data.get('rotation', 90),
                    va='bottom',
                    ha='center'
                )

            # Clear selection when removing
            if self.selection_box:
                self.selection_box.remove()
                self.selection_box = None
            self.selected_text = None

            self.parent.canvas.draw_idle()
            self.update_list()

            # Reselect appropriate row after deletion
            if len(labels) > 0:  # If there are still labels
                # If we deleted the last row, select the new last row
                if row_to_select >= len(labels):
                    row_to_select = len(labels) - 1

                # Select the row
                self.labels_grid.SelectRow(row_to_select)
                self.labels_grid.SetGridCursor(row_to_select, 0)

    def on_remove_all(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()

        if 'Labels' in self.parent.Data['Core levels'][sheet_name]:
            # Confirm removal
            dlg = wx.MessageDialog(self,
                                   "Are you sure you want to remove all labels?",
                                   "Confirm Remove All",
                                   wx.YES_NO | wx.ICON_QUESTION)

            if dlg.ShowModal() == wx.ID_YES:
                # Clear all labels
                self.parent.Data['Core levels'][sheet_name]['Labels'].clear()

                # Clear all text annotations from plot
                for txt in self.parent.ax.texts[:]:
                    txt.remove()

                # Clear selection
                if self.selection_box:
                    self.selection_box.remove()
                    self.selection_box = None
                self.selected_text = None

                self.parent.canvas.draw_idle()
                self.update_list()

            dlg.Destroy()

    def on_key_down(self, event):
        if event.GetKeyCode() == wx.WXK_DELETE:
            # Get the current selection
            current_row = self.labels_grid.GetGridCursorRow()
            if current_row >= 0:
                # Create a mock event for on_remove
                self.on_remove(event)
        else:
            event.Skip()

    def on_cell_changed(self, event):
        row = event.GetRow()
        col = event.GetCol()
        new_value = self.labels_grid.GetCellValue(row, col)

        sheet_name = self.parent.sheet_combobox.GetValue()

        if 'Labels' not in self.parent.Data['Core levels'][sheet_name] or row >= len(self.parent.Data['Core levels'][sheet_name]['Labels']):
            return

        label = self.parent.Data['Core levels'][sheet_name]['Labels'][row]

        try:
            if col == 0:  # Name column
                label['text'] = new_value
            elif col == 1:  # Pos X column
                label['x'] = float(new_value)
            elif col == 2:  # Pos Y column
                label['y'] = float(new_value)
            elif col == 3:  # Rotation column
                # Convert to int for rotation
                label['rotation'] = int(float(new_value))

            # Update the plot immediately
            for txt in self.parent.ax.texts[:]:
                txt.remove()

            for label_data in self.parent.Data['Core levels'][sheet_name]['Labels']:
                self.parent.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=label_data.get('rotation', 90),
                    fontsize=label_data.get('fontsize', 10),
                    fontfamily=label_data.get('fontfamily', 'Arial'),
                    va='bottom',
                    ha='center'
                )

            self.parent.canvas.draw_idle()

            # Update the grid display to show proper formatting
            if col in [1, 2]:  # Pos X, Pos Y columns
                self.labels_grid.SetCellValue(row, col, f"{float(new_value):.2f}")
            elif col == 3:  # Rotation column
                self.labels_grid.SetCellValue(row, col, f"{int(float(new_value)):.2f}")

        except ValueError:
            # If invalid number entered, revert to original value
            if col == 1:
                self.labels_grid.SetCellValue(row, col, f"{label['x']:.2f}")
            elif col == 2:
                self.labels_grid.SetCellValue(row, col, f"{label['y']:.2f}")
            elif col == 3:
                self.labels_grid.SetCellValue(row, col, f"{label.get('rotation', 90):.2f}")

            wx.MessageBox("Please enter a valid number", "Invalid Input", wx.OK | wx.ICON_ERROR)

        event.Skip()

    def on_cell_select(self, event):
        # When a cell is selected, select the entire row
        row = event.GetRow()
        self.labels_grid.SelectRow(row)
        self.labels_grid.SetFocus()
        event.Skip()

    def on_close(self, event):
        # Unbind parent sheet change event first
        if hasattr(self.parent, 'sheet_combobox') and self.parent.sheet_combobox:
            try:
                self.parent.sheet_combobox.Unbind(wx.EVT_COMBOBOX, handler=self.on_parent_sheet_change)
            except:
                pass

        # Clear any selection before closing
        if self.selection_box:
            self.selection_box.remove()
            self.selection_box = None
        self.selected_text = None
        self.parent.canvas.draw_idle()

        # Disconnect canvas events using stored IDs
        if hasattr(self, 'canvas_click_id'):
            self.parent.canvas.mpl_disconnect(self.canvas_click_id)
        if hasattr(self, 'canvas_release_id'):
            self.parent.canvas.mpl_disconnect(self.canvas_release_id)
        if hasattr(self, 'canvas_motion_id'):
            self.parent.canvas.mpl_disconnect(self.canvas_motion_id)

        # Reset parent reference to labels_window
        self.parent.labels_window = None

        self.Destroy()