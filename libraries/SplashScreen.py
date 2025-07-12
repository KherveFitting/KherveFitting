import wx
import wx.adv
import os
import sys

class SplashScreen:
    def __init__(self, duration=3000):  # delay in milliseconds
        self.duration=duration
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        splash_image_path = os.path.join(application_path, "libraries", "Images", "splash_600.png")
        self.splash = None

        if os.path.exists(splash_image_path):
            bitmap = wx.Bitmap(splash_image_path)
            if bitmap.IsOk():
                self.splash = wx.adv.SplashScreen(
                    bitmap,
                    wx.adv.SPLASH_CENTRE_ON_SCREEN,
                    0,  # we will manually close it
                    None,
                    -1
                )
                wx.YieldIfNeeded()
            else:
                print(f"Failed to load bitmap: {splash_image_path}")
        else:
            print(f"Splash image not found: {splash_image_path}")

    def Show(self):
        if self.splash:
            self.splash.Show()
            wx.YieldIfNeeded()
            wx.CallLater(self.duration, self.Destroy)

    def Destroy(self):
        print('destroy')
        if self.splash:
            self.splash.Destroy()
            self.splash = None

def show_splash(duration=3000):
    splash = SplashScreen(duration)
    splash.Show()
    return splash
