import wx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import numpy as np
from scipy.stats import linregress
from matplotlib.ticker import ScalarFormatter


class NoiseAnalysisWindow2(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Noise Analysis", style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetSize(1200, 400)  # Adjusted window size

        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create a figure with three subplots
        self.figure, (self.ax1, self.ax2, self.ax3) = plt.subplots(1, 3, figsize=(10, 1))
        self.figure.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
        self.canvas = FigureCanvas(panel, -1, self.figure)

        main_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        noise_button = wx.Button(panel, label="Calculate Noise")
        noise_button.Bind(wx.EVT_BUTTON, self.on_noise)
        button_sizer.Add(noise_button, 0, wx.ALL, 2)

        clear_noise_button = wx.Button(panel, label="Clear Noise")
        clear_noise_button.Bind(wx.EVT_BUTTON, self.on_clear_noise)
        button_sizer.Add(clear_noise_button, 0, wx.ALL, 2)

        main_sizer.Add(button_sizer, 0, wx.ALIGN_CENTER)

        panel.SetSizer(main_sizer)

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def on_noise(self, event):
        result = PlotManager.plot_noise(self.parent)
        if result:
            x_values, y_values, linear_fit, noise_subtracted, std_value = result
            self.update_plots(x_values, y_values, linear_fit, noise_subtracted, std_value)

    def update_plots(self, x_values, y_values, linear_fit, noise_subtracted, std_value):
        # Clear previous plots
        for ax in [self.ax1, self.ax2, self.ax3]:
            ax.clear()

        # Set global font sizes
        plt.rcParams.update({'font.size': 6})

        # Plot 1: Raw data and linear fit
        self.ax1.scatter(x_values, y_values, color='black', s=10)
        self.ax1.plot(x_values, linear_fit, color='red', linewidth=2)
        self.ax1.set_title('Raw Data and Linear Fit', fontsize=8)
        self.ax1.set_xlabel('Binding Energy (eV)', fontsize=8)
        self.ax1.set_ylabel('Intensity (CTS)', fontsize=8)
        self.ax1.tick_params(axis='both', which='major', labelsize=6)

        # Plot 2: Corrected noise data
        self.ax2.scatter(x_values, noise_subtracted, color='blue', s=10)
        self.ax2.axhline(y=0, color='red', linestyle='--', linewidth=2)
        self.ax2.set_title('Corrected Noise Data', fontsize=8)
        self.ax2.set_xlabel('Binding Energy (eV)', fontsize=8)
        self.ax2.set_ylabel('Residual Intensity', fontsize=8)
        self.ax2.tick_params(axis='both', which='major', labelsize=6)

        # Plot 3: Histogram of results
        self.ax3.hist(noise_subtracted, bins=5, edgecolor='black', linewidth=2)
        self.ax3.set_title(f'Noise Histogram\nSTD: {int(std_value)} cts', fontsize=8)
        self.ax3.set_xlabel('Residual Intensity', fontsize=8)
        self.ax3.set_ylabel('Frequency', fontsize=8)
        self.ax3.tick_params(axis='both', which='major', labelsize=6)

        self.figure.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
        self.canvas.draw()


    def on_clear_noise(self, event):
        PlotManager.remove_noise_inset(self.parent)
        # Clear the plots
        for ax in [self.ax1, self.ax2, self.ax3]:
            ax.clear()
        self.canvas.draw()

    def on_close(self, event):
        self.parent.noise_window_closed()
        self.Destroy()





class NoiseAnalysisWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Noise Analysis", style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetSize(1200, 400)

        self.init_ui()
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        """Initialize the user interface components."""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        self.init_figure(panel, main_sizer)
        self.init_buttons(panel, main_sizer)

        panel.SetSizer(main_sizer)

    def init_figure(self, panel, sizer):
        """Initialize the matplotlib figure with three subplots."""
        self.figure, (self.ax1, self.ax2, self.ax3) = plt.subplots(1, 3, figsize=(10, 1))
        self.figure.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
        self.canvas = FigureCanvas(panel, -1, self.figure)
        sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

    def init_buttons(self, panel, sizer):
        """Initialize buttons and add them to the sizer."""
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        noise_button = wx.Button(panel, label="Calculate Noise")
        noise_button.Bind(wx.EVT_BUTTON, self.on_noise)
        button_sizer.Add(noise_button, 0, wx.ALL, 2)

        clear_noise_button = wx.Button(panel, label="Clear Noise")
        clear_noise_button.Bind(wx.EVT_BUTTON, self.on_clear_noise)
        button_sizer.Add(clear_noise_button, 0, wx.ALL, 2)

        sizer.Add(button_sizer, 0, wx.ALIGN_CENTER)

    def on_noise(self, event):
        """Handle the Calculate Noise button click."""
        result = self.plot_noise(self.parent)
        if result:
            x_values, y_values, linear_fit, noise_subtracted, std_value = result
            self.update_plots(x_values, y_values, linear_fit, noise_subtracted, std_value)

    def update_plots2(self, x_values, y_values, linear_fit, noise_subtracted, std_value):
        """Update the plots with new data."""
        # Clear previous plots
        for ax in [self.ax1, self.ax2, self.ax3]:
            ax.clear()

        # Set global font sizes
        plt.rcParams.update({'font.size': 6})

        # Plot 1: Raw data and linear fit
        self.ax1.scatter(x_values, y_values, color='black', s=10)
        self.ax1.plot(x_values, linear_fit, color='red', linewidth=2)
        self.ax1.set_title('Raw Data and Linear Fit', fontsize=8)
        self.ax1.set_xlabel('Binding Energy (eV)', fontsize=8)
        self.ax1.set_ylabel('Intensity (CTS)', fontsize=8)
        self.ax1.tick_params(axis='both', which='major', labelsize=6)

        # Plot 2: Corrected noise data
        self.ax2.scatter(x_values, noise_subtracted, color='blue', s=10)
        self.ax2.axhline(y=0, color='red', linestyle='--', linewidth=2)
        self.ax2.set_title('Corrected Noise Data', fontsize=8)
        self.ax2.set_xlabel('Binding Energy (eV)', fontsize=8)
        self.ax2.set_ylabel('Residual Intensity', fontsize=8)
        self.ax2.tick_params(axis='both', which='major', labelsize=6)

        # Plot 3: Histogram of results
        self.ax3.hist(noise_subtracted, bins=20, edgecolor='black', linewidth=2)
        self.ax3.set_title(f'Noise Histogram\nSTD: {int(std_value)} cts', fontsize=8)
        self.ax3.set_xlabel('Residual Intensity', fontsize=8)
        self.ax3.set_ylabel('Frequency', fontsize=8)
        self.ax3.tick_params(axis='both', which='major', labelsize=6)

        self.figure.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
        self.canvas.draw()

    def update_plots(self, x_values, y_values, linear_fit, noise_subtracted, std_value):
        """Update the plots with new data."""
        # Clear previous plots
        for ax in [self.ax1, self.ax2, self.ax3]:
            ax.clear()

        # Store original rcParams
        original_fontsize = plt.rcParams['font.size']

        try:
            # Set font sizes locally for each subplot
            for ax in [self.ax1, self.ax2, self.ax3]:
                ax.tick_params(axis='both', which='major', labelsize=8)
                for item in ([ax.title, ax.xaxis.label, ax.yaxis.label] +
                             ax.get_xticklabels() + ax.get_yticklabels()):
                    item.set_fontsize(8)

            # Plot 1: Raw data and linear fit
            self.ax1.scatter(x_values, y_values, color='black', s=10)
            self.ax1.plot(x_values, linear_fit, color='red', linewidth=2)
            self.ax1.set_title('Raw Data and Linear Fit', fontsize=10)
            self.ax1.set_xlabel('Binding Energy (eV)')
            self.ax1.set_ylabel('Intensity (CTS)')

            # Plot 2: Corrected noise data
            self.ax2.scatter(x_values, noise_subtracted, color='blue', s=10)
            self.ax2.axhline(y=0, color='red', linestyle='--', linewidth=2)
            self.ax2.set_title('Corrected Noise Data', fontsize=10)
            self.ax2.set_xlabel('Binding Energy (eV)')
            self.ax2.set_ylabel('Residual Intensity')

            # Plot 3: Histogram of results
            self.ax3.hist(noise_subtracted, bins=20, edgecolor='black', linewidth=2)
            self.ax3.set_title(f'Noise Histogram, STD: {int(std_value)} cts', fontsize=10)
            self.ax3.set_xlabel('Residual Intensity')
            self.ax3.set_ylabel('Frequency')

            self.figure.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
            self.canvas.draw()

        finally:
            # Reset to original rcParams
            plt.rcParams.update({'font.size': original_fontsize})

    def plot_noise(self, window):
        """
        Plot noise analysis in an inset axes.

        Args:
            window: The main application window

        Returns:
            tuple: Filtered x and y values, linear fit, noise subtracted, and standard deviation
        """
        if window.noise_min_energy is not None and window.noise_max_energy is not None:
            mask = (window.x_values >= window.noise_min_energy) & (window.x_values <= window.noise_max_energy)
            x_values_filtered = window.x_values[mask]
            y_values_filtered = window.y_values[mask]

            if len(x_values_filtered) > 0 and len(y_values_filtered) > 0:
                # Fit a linear line to the noise data
                slope, intercept, r_value, p_value, std_err = linregress(x_values_filtered, y_values_filtered)
                linear_fit = slope * x_values_filtered + intercept
                noise_subtracted = y_values_filtered - linear_fit

                # Calculate the standard deviation
                std_value = np.std(noise_subtracted)
                window.noise_std_value = std_value  # Save noise STD value in window instance

                # Create the inset plot in the main window
                if hasattr(window, 'noise_inset_ax') and window.noise_inset_ax:
                    window.noise_inset_ax.clear()
                else:
                    window.noise_inset_ax = window.ax.inset_axes(
                        [0.05, 0.05, 0.2, 0.2])  # Use window.ax instead of self.ax

                window.noise_inset_ax.hist(noise_subtracted, bins=15, histtype='bar', edgecolor='black')
                window.noise_inset_ax.tick_params(axis='both', which='major', labelsize=8)
                window.noise_inset_ax.xaxis.set_major_formatter(ScalarFormatter(useMathText=True))
                window.noise_inset_ax.set_facecolor('none')

                # Hide top, right, and left borders, and the y-axis
                window.noise_inset_ax.spines['top'].set_visible(False)
                window.noise_inset_ax.spines['right'].set_visible(False)
                window.noise_inset_ax.spines['left'].set_visible(False)
                window.noise_inset_ax.yaxis.set_visible(False)

                # Add standard deviation text
                std_value_int = int(std_value)
                window.noise_inset_ax.text(0.5, 1.1, f'STD: {std_value_int} cts',
                                           transform=window.noise_inset_ax.transAxes,
                                           fontsize=8,
                                           verticalalignment='top',
                                           horizontalalignment='center')

                window.canvas.draw()

                # Return the data for the NoiseAnalysisWindow
                return x_values_filtered, y_values_filtered, linear_fit, noise_subtracted, std_value

        return None

    def remove_noise_inset(self, window):
        """
        Remove the noise inset plot.

        Args:
            window: The main application window
        """
        if hasattr(window, 'noise_inset_ax') and window.noise_inset_ax:
            window.noise_inset_ax.clear()
            window.noise_inset_ax.remove()
            window.noise_inset_ax = None
            window.canvas.draw()

    def on_clear_noise(self, event):
        """Handle the Clear Noise button click."""
        self.remove_noise_inset(self.parent)
        # Clear the plots
        for ax in [self.ax1, self.ax2, self.ax3]:
            ax.clear()
        self.canvas.draw()

    def on_close(self, event):
        """Handle the window close event."""
        self.parent.noise_window_closed()
        self.Destroy()
