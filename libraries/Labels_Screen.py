import wx
from matplotlib.patches import Rectangle
from matplotlib.transforms import Affine2D
import numpy as np

class LabelWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Labels Manager", style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetSize((300, 300))

        # Position on right side of main frame
        main_pos = self.parent.GetPosition()
        main_size = self.parent.GetSize()
        window_size = self.GetSize()
        x = main_pos.x + main_size.width - 350
        y = main_pos.y + (main_size.height - window_size.height) // 2
        self.SetPosition((x, y))

        panel = wx.Panel(self)
        # panel.SetBackgroundColour(wx.WHITE)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.list_box = wx.ListBox(panel, style=wx.LB_SINGLE)
        sizer.Add(self.list_box, 1, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        add_btn = wx.Button(panel, label="Add")
        add_btn.SetMinSize((60, 40))
        edit_btn = wx.Button(panel, label="Edit")
        edit_btn.SetMinSize((60, 40))
        remove_btn = wx.Button(panel, label="Remove")
        remove_btn.SetMinSize((60, 40))

        btn_sizer.Add(add_btn, 0, wx.ALL, 5)
        btn_sizer.Add(edit_btn, 0, wx.ALL, 5)
        btn_sizer.Add(remove_btn, 0, wx.ALL, 5)

        add_btn.Bind(wx.EVT_BUTTON, self.on_add)
        edit_btn.Bind(wx.EVT_BUTTON, self.on_edit)
        remove_btn.Bind(wx.EVT_BUTTON, self.on_remove)

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

        sizer.Add(buttons_sizer, 0, wx.LEFT | wx.TOP, 10)  # Changed ALIGN_RIGHT to LEFT

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

        # Bind close event
        self.Bind(wx.EVT_CLOSE, self.on_close)

        panel.SetSizer(sizer)

        self.update_list()

    def move_label(self, direction):
        selection = self.list_box.GetSelection()
        if selection == wx.NOT_FOUND:
            return

        sheet_name = self.parent.sheet_combobox.GetValue()
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
        self.list_box.SetSelection(selection)  # Maintain selection

    def move_up(self, event):
        self.move_label('up')

    def move_down(self, event):
        self.move_label('down')

    def move_left(self, event):
        self.move_label('left')

    def move_right(self, event):
        self.move_label('right')


    def update_list(self):
        self.list_box.Clear()
        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Labels' in self.parent.Data['Core levels'][sheet_name]:
            for label in self.parent.Data['Core levels'][sheet_name]['Labels']:
                self.list_box.Append(f"{label['text']} ({label['x']:.1f}, {label['y']:.1f})")

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

        # Find clicked text
        clicked_text = None
        clicked_index = None

        for i, label_data in enumerate(self.parent.Data['Core levels'][sheet_name]['Labels']):
            # Check if click is near text position
            text_x = label_data['x']
            text_y = label_data['y']

            # Create approximate bounding box for text
            bbox_width = len(label_data['text']) * 0.1  # Approximate width
            bbox_height = 0.05 * max(self.parent.y_values)  # Approximate height

            if (abs(event.xdata - text_x) < bbox_width and
                    abs(event.ydata - text_y) < bbox_height):
                clicked_text = label_data
                clicked_index = i
                break

        if clicked_text:
            self.select_text(clicked_text, clicked_index)
            self.drag_offset = (event.xdata - clicked_text['x'], event.ydata - clicked_text['y'])
            self.is_dragging = True
            self.list_box.SetSelection(clicked_index)

    def select_text_rectangle(self, label_data, index):
        # Remove previous selection box
        if self.selection_box:
            self.selection_box.remove()

        # Create selection box around text
        x = label_data['x']
        y = label_data['y']
        text_width = len(label_data['text']) * 0.4
        text_height = 0.03 * max(self.parent.y_values)

        from matplotlib.patches import Rectangle
        self.selection_box = Rectangle(
            (x - text_width / 2, y),
            text_width, text_height,
            linewidth=1, edgecolor=(200 / 255, 245 / 255, 228 / 255), facecolor='none', linestyle='-')
        self.parent.ax.add_patch(self.selection_box)
        self.selected_text = label_data
        self.parent.canvas.draw_idle()

    def on_canvas_motion_rectangle(self, event):
        if not self.is_dragging or not self.selected_text or not event.inaxes:
            return

        # Update text position
        new_x = event.xdata - self.drag_offset[0]
        new_y = event.ydata - self.drag_offset[1]

        self.selected_text['x'] = new_x
        self.selected_text['y'] = new_y

        # Update selection box position
        if self.selection_box:
            text_width = len(self.selected_text['text']) * 0.4
            text_height = 0.03 * max(self.parent.y_values)
            self.selection_box.set_xy((new_x - text_width / 2, new_y))# - text_height / 2))

        # Redraw all text
        self.redraw_labels()

    def select_text(self, label_data, index):
        # Remove previous selection box
        if self.selection_box:
            self.selection_box.remove()

        # Create triangle selection indicator
        x = label_data['x']
        y = label_data['y']

        # Get axis ranges for proper scaling
        xlim = self.parent.ax.get_xlim()
        ylim = self.parent.ax.get_ylim()
        x_range = abs(xlim[1] - xlim[0])
        y_range = abs(ylim[1] - ylim[0])

        # Triangle dimensions based on axis ranges
        triangle_width = x_range * 0.02  # 2% of x-axis range (in eV)
        triangle_height = y_range * 0.03  # 3% of y-axis range (in CPS)

        # Triangle positioned slightly lower
        triangle_y = y - y_range * 0.01  # 1% below text

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
            facecolor= (200 / 255, 245 / 255, 228 / 255),
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

            triangle_width = x_range * 0.02
            triangle_height = y_range * 0.03
            triangle_y = new_y - y_range * 0.01

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
        selection = self.list_box.GetSelection()
        if selection != wx.NOT_FOUND:
            sheet_name = self.parent.sheet_combobox.GetValue()
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

            # Replace font family control (add after existing controls)
            font_family_sizer = wx.BoxSizer(wx.HORIZONTAL)
            font_family_label = wx.StaticText(dlg, label="Font:")
            font_families = ['Arial', 'Times New Roman', 'Courier New', 'Helvetica', 'Verdana',
                             'Calibri', 'Tahoma', 'Georgia', 'Comic Sans MS', 'Impact',
                             'Trebuchet MS', 'Century Gothic', 'Palatino']
            font_family_ctrl = wx.Choice(dlg, choices=font_families, size=(120, -1))

            # Set current font
            current_font = label.get('fontfamily', 'Arial')
            if current_font in font_families:
                font_family_ctrl.SetSelection(font_families.index(current_font))
            else:
                font_family_ctrl.SetSelection(0)  # Default to Arial

            font_family_sizer.Add(font_family_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
            font_family_sizer.Add(font_family_ctrl, 1, wx.EXPAND | wx.ALL, 5)
            sizer.Add(font_family_sizer, 0, wx.EXPAND)


            # Add buttons
            btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
            ok_btn = wx.Button(dlg, wx.ID_OK, "OK", size=(60, 40))
            cancel_btn = wx.Button(dlg, wx.ID_CANCEL, "Cancel", size=(60, 40))
            btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
            btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
            sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)

            dlg.SetSizer(sizer)
            sizer.Fit(dlg)

            if dlg.ShowModal() == wx.ID_OK:
                label['text'] = text_ctrl.GetValue()
                label['x'] = x_ctrl.GetValue()
                label['y'] = y_ctrl.GetValue()
                label['rotation'] = int(rotation_choices[rotation_ctrl.GetSelection()])  # Convert to int
                label['fontfamily'] = font_families[font_family_ctrl.GetSelection()]  # Add this line

                # Clear all existing text annotations
                for txt in self.parent.ax.texts[:]:
                    txt.remove()

                # self.parent.clear_and_replot()  # Full replot to ensure consistency

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
        selection = self.list_box.GetSelection()
        if selection != wx.NOT_FOUND:
            sheet_name = self.parent.sheet_combobox.GetValue()
            labels = self.parent.Data['Core levels'][sheet_name]['Labels']
            labels.pop(selection)

            # # Clear all existing text annotations
            # for text in self.parent.ax.texts[:]:
            #     text.remove()


            self.parent.clear_and_replot()  # Full replot to ensure consistency

            # Redraw remaining labels
            for label_data in labels:
                self.parent.ax.text(
                    label_data['x'],
                    label_data['y'],
                    label_data['text'],
                    rotation=90,
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

    def on_close(self, event):
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
