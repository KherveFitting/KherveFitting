# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

import wx
import os
import sys


class CustomSplashScreen(wx.Frame):
    def __init__(self):
        super().__init__(None, style=wx.FRAME_NO_TASKBAR | wx.STAY_ON_TOP | wx.BORDER_NONE)

        # Find image - optimized for compiled exe
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            base_path = sys._MEIPASS
        else:
            # Running from source
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # Try different image names and locations (prioritize most likely)
        image_paths = [
            # os.path.join(base_path, "Images", "SplashScreen4.png"),
            # os.path.join(base_path, "libraries", "Images", "SplashScreen4.png"),
            os.path.join(base_path, "Images", "splash_600.png"),
            os.path.join(base_path, "libraries", "Images", "splash_600.png"),
        ]

        self.bitmap = None
        for img_path in image_paths:
            if os.path.exists(img_path):
                try:
                    # Load with explicit format for speed
                    img = wx.Image(img_path, wx.BITMAP_TYPE_PNG)
                    if img.IsOk():
                        # Scale down if too large (faster display)
                        if img.GetWidth() > 600:
                            img.Rescale(600, int(600 * img.GetHeight() / img.GetWidth()), wx.IMAGE_QUALITY_HIGH)
                        self.bitmap = wx.Bitmap(img)
                        break
                except:
                    pass

        # Fast fallback: simple colored bitmap
        if not self.bitmap or not self.bitmap.IsOk():
            self.bitmap = wx.Bitmap(600, 400)
            dc = wx.MemoryDC(self.bitmap)
            dc.SetBackground(wx.Brush(wx.Colour(45, 85, 125)))
            dc.Clear()
            dc.SetFont(wx.Font(28, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            dc.SetTextForeground(wx.WHITE)
            dc.DrawText("KherveFitting", 150, 180)
            dc.SelectObject(wx.NullBitmap)

        # Set window size
        w, h = self.bitmap.GetWidth(), self.bitmap.GetHeight()
        self.SetClientSize((w, h))
        self.CenterOnScreen()

        # Progress message
        self.message = "Starting..."

        # Create a panel that we'll paint on
        self.panel = wx.Panel(self)
        self.panel.SetSize((w, h))
        self.panel.Bind(wx.EVT_PAINT, self.on_paint)

        # Force immediate display
        self.Show()
        self.Raise()
        self.Update()
        self.Refresh()
        wx.SafeYield()

    def on_paint(self, event):
        """Paint the bitmap and text"""
        dc = wx.PaintDC(self.panel)

        # Draw the splash image
        dc.DrawBitmap(self.bitmap, 0, 0, True)

        # Draw text in pure white
        dc.SetFont(wx.Font(11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        dc.SetTextForeground(wx.WHITE)

        # Get text size and center it at bottom
        text_w, text_h = dc.GetTextExtent(self.message)
        w = self.bitmap.GetWidth()
        h = self.bitmap.GetHeight()
        x = (w - text_w) // 2
        y = h - text_h - 15

        # Draw text directly on image (no background)
        dc.DrawText(self.message, x, y)

    def update_message(self, message):
        """Update progress message"""
        self.message = message
        self.panel.Refresh()
        self.Update()
        wx.SafeYield()


def show_splash():
    """Create and show splash screen immediately"""
    splash = CustomSplashScreen()
    return splash

