# -*- coding: utf-8 -*-
"""pyArea Preferences Dialog"""

__title__ = "Preferences"
__author__ = "pyArea"

import os
import sys
import clr

# Add lib to path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, "..", "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')

import System.Windows
from System.Windows import Window, Application, Thickness, WindowStartupLocation, GridLength, GridUnitType, HorizontalAlignment
from System.Windows.Controls import (
    StackPanel, GroupBox, Label, TextBox, CheckBox, Button,
    RadioButton, Orientation, Grid, ColumnDefinition, RowDefinition
)
from System.Windows.Input import FocusNavigationDirection
from System.Windows.Forms import FolderBrowserDialog, DialogResult

from pyrevit import forms
from data_manager import get_preferences, set_preferences
from export_utils import get_default_preferences


class PreferencesWindow(Window):
    """WPF Window for editing export preferences"""
    
    def __init__(self):
        self.Title = "pyArea Preferences"
        self.Width = 500
        self.Height = 520
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        
        # Load current preferences
        self.preferences = get_preferences()
        
        # Create UI
        self._create_ui()
    
    def _create_ui(self):
        """Create the WPF UI"""
        main_panel = StackPanel()
        main_panel.Margin = Thickness(15)
        
        # Export folder section
        export_group = self._create_export_folder_section()
        main_panel.Children.Add(export_group)
        
        # DXF settings section
        dxf_group = self._create_dxf_section()
        main_panel.Children.Add(dxf_group)
        
        # DWFx Postprocessing section (emphasized)
        dwfx_postprocess_group = self._create_dwfx_postprocessing_section()
        main_panel.Children.Add(dwfx_postprocess_group)
        
        # DWFX settings section
        dwfx_group = self._create_dwfx_section()
        main_panel.Children.Add(dwfx_group)
        
        # Buttons
        button_panel = self._create_buttons()
        main_panel.Children.Add(button_panel)
        
        self.Content = main_panel
    
    def _create_export_folder_section(self):
        """Create export folder input section"""
        group = GroupBox()
        group.Header = "Export Folder"
        group.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        # Folder path row
        folder_grid = Grid()
        folder_grid.ColumnDefinitions.Add(ColumnDefinition())
        folder_grid.ColumnDefinitions.Add(ColumnDefinition())
        folder_grid.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Star)
        folder_grid.ColumnDefinitions[1].Width = GridLength.Auto
        
        self.folder_textbox = TextBox()
        self.folder_textbox.Text = self.preferences.get("ExportFolder", "Desktop/Export")
        self.folder_textbox.Margin = Thickness(0, 0, 5, 0)
        Grid.SetColumn(self.folder_textbox, 0)
        folder_grid.Children.Add(self.folder_textbox)
        
        browse_btn = Button()
        browse_btn.Content = "Browse..."
        browse_btn.Width = 80
        browse_btn.Click += self._on_browse_folder
        Grid.SetColumn(browse_btn, 1)
        folder_grid.Children.Add(browse_btn)
        
        panel.Children.Add(folder_grid)
        group.Content = panel
        return group
    
    def _create_dxf_section(self):
        """Create DXF settings section"""
        group = GroupBox()
        group.Header = "DXF Settings"
        group.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        self.dxf_dat_checkbox = CheckBox()
        self.dxf_dat_checkbox.Content = "Create .dat file with DWFx scale"
        self.dxf_dat_checkbox.IsChecked = self.preferences.get("DXF_CreateDatFile", True)
        panel.Children.Add(self.dxf_dat_checkbox)
        
        group.Content = panel
        return group
    
    def _create_dwfx_postprocessing_section(self):
        """Create DWFx Postprocessing section (emphasized)"""
        group = GroupBox()
        group.Header = "DWFx Postprocessing"
        group.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        # Remove opaque white checkbox
        self.dwfx_remove_white_checkbox = CheckBox()
        self.dwfx_remove_white_checkbox.Content = "Remove opaque white background"
        self.dwfx_remove_white_checkbox.IsChecked = self.preferences.get("DWFx_RemoveOpaqueWhite", True)
        self.dwfx_remove_white_checkbox.FontWeight = System.Windows.FontWeights.SemiBold
        panel.Children.Add(self.dwfx_remove_white_checkbox)
        
        # Description label
        desc_label = Label()
        desc_label.Content = "Processes exported DWFx files to make white fills transparent"
        desc_label.Foreground = System.Windows.Media.Brushes.Gray
        desc_label.FontSize = 11
        desc_label.Padding = Thickness(0)
        desc_label.Margin = Thickness(18, 0, 0, 0)
        panel.Children.Add(desc_label)
        
        group.Content = panel
        return group
    
    def _create_dwfx_section(self):
        """Create DWFx settings section"""
        group = GroupBox()
        group.Header = "DWFx Settings"
        group.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        # Export Element Data checkbox
        self.dwfx_element_data_checkbox = CheckBox()
        self.dwfx_element_data_checkbox.Content = "Export Element Data"
        self.dwfx_element_data_checkbox.IsChecked = self.preferences.get("DWFx_ExportElementData", True)
        self.dwfx_element_data_checkbox.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.dwfx_element_data_checkbox)
        
        # --- Graphics Settings subsection ---
        graphics_label = Label()
        graphics_label.Content = "Graphics Settings:"
        graphics_label.FontWeight = System.Windows.FontWeights.Bold
        graphics_label.Margin = Thickness(0, 5, 0, 5)
        graphics_label.Padding = Thickness(0)
        panel.Children.Add(graphics_label)
        
        # Standard format radio
        self.graphics_standard_radio = RadioButton()
        self.graphics_standard_radio.Content = "Use standard format"
        self.graphics_standard_radio.GroupName = "GraphicsFormat"
        self.graphics_standard_radio.IsChecked = not self.preferences.get("DWFx_UseCompressedRaster", False)
        self.graphics_standard_radio.Margin = Thickness(10, 0, 0, 3)
        self.graphics_standard_radio.Checked += self._on_graphics_format_changed
        panel.Children.Add(self.graphics_standard_radio)
        
        # Compressed raster format radio
        self.graphics_compressed_radio = RadioButton()
        self.graphics_compressed_radio.Content = "Use compressed raster format"
        self.graphics_compressed_radio.GroupName = "GraphicsFormat"
        self.graphics_compressed_radio.IsChecked = self.preferences.get("DWFx_UseCompressedRaster", False)
        self.graphics_compressed_radio.Margin = Thickness(10, 0, 0, 3)
        self.graphics_compressed_radio.Checked += self._on_graphics_format_changed
        panel.Children.Add(self.graphics_compressed_radio)
        
        # Image Quality row using Grid for alignment
        image_quality_grid = Grid()
        image_quality_grid.Margin = Thickness(30, 0, 0, 10)
        image_quality_grid.ColumnDefinitions.Add(ColumnDefinition())
        image_quality_grid.ColumnDefinitions.Add(ColumnDefinition())
        image_quality_grid.ColumnDefinitions[0].Width = GridLength(90)
        image_quality_grid.ColumnDefinitions[1].Width = GridLength(1, GridUnitType.Star)
        
        self.image_quality_label = Label()
        self.image_quality_label.Content = "Image Quality:"
        self.image_quality_label.Padding = Thickness(0)
        self.image_quality_label.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(self.image_quality_label, 0)
        image_quality_grid.Children.Add(self.image_quality_label)
        
        image_quality_radios = StackPanel()
        image_quality_radios.Orientation = Orientation.Horizontal
        image_quality_radios.VerticalAlignment = System.Windows.VerticalAlignment.Center
        
        current_image_quality = self.preferences.get("DWFx_ImageQuality", "Low")
        
        self.image_quality_low_radio = RadioButton()
        self.image_quality_low_radio.Content = "Low"
        self.image_quality_low_radio.GroupName = "ImageQuality"
        self.image_quality_low_radio.IsChecked = (current_image_quality == "Low")
        self.image_quality_low_radio.Margin = Thickness(0, 0, 15, 0)
        image_quality_radios.Children.Add(self.image_quality_low_radio)
        
        self.image_quality_medium_radio = RadioButton()
        self.image_quality_medium_radio.Content = "Medium"
        self.image_quality_medium_radio.GroupName = "ImageQuality"
        self.image_quality_medium_radio.IsChecked = (current_image_quality == "Medium")
        self.image_quality_medium_radio.Margin = Thickness(0, 0, 15, 0)
        image_quality_radios.Children.Add(self.image_quality_medium_radio)
        
        self.image_quality_high_radio = RadioButton()
        self.image_quality_high_radio.Content = "High"
        self.image_quality_high_radio.GroupName = "ImageQuality"
        self.image_quality_high_radio.IsChecked = (current_image_quality == "High")
        image_quality_radios.Children.Add(self.image_quality_high_radio)
        
        Grid.SetColumn(image_quality_radios, 1)
        image_quality_grid.Children.Add(image_quality_radios)
        panel.Children.Add(image_quality_grid)
        
        # Store reference to image quality controls for enabling/disabling
        self._image_quality_controls = [self.image_quality_label, self.image_quality_low_radio, 
                                         self.image_quality_medium_radio, self.image_quality_high_radio]
        self._update_image_quality_enabled()
        
        # --- Appearance Settings subsection ---
        appearance_label = Label()
        appearance_label.Content = "Appearance:"
        appearance_label.FontWeight = System.Windows.FontWeights.Bold
        appearance_label.Margin = Thickness(0, 5, 0, 5)
        appearance_label.Padding = Thickness(0)
        panel.Children.Add(appearance_label)
        
        # Raster Quality row using Grid for alignment
        raster_quality_grid = Grid()
        raster_quality_grid.Margin = Thickness(10, 0, 0, 3)
        raster_quality_grid.ColumnDefinitions.Add(ColumnDefinition())
        raster_quality_grid.ColumnDefinitions.Add(ColumnDefinition())
        raster_quality_grid.ColumnDefinitions[0].Width = GridLength(90)
        raster_quality_grid.ColumnDefinitions[1].Width = GridLength(1, GridUnitType.Star)
        
        raster_quality_label = Label()
        raster_quality_label.Content = "Raster quality:"
        raster_quality_label.Padding = Thickness(0)
        raster_quality_label.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(raster_quality_label, 0)
        raster_quality_grid.Children.Add(raster_quality_label)
        
        raster_quality_radios = StackPanel()
        raster_quality_radios.Orientation = Orientation.Horizontal
        raster_quality_radios.VerticalAlignment = System.Windows.VerticalAlignment.Center
        
        current_raster_quality = self.preferences.get("DWFx_RasterQuality", "High")
        
        self.raster_quality_low_radio = RadioButton()
        self.raster_quality_low_radio.Content = "Low"
        self.raster_quality_low_radio.GroupName = "RasterQuality"
        self.raster_quality_low_radio.IsChecked = (current_raster_quality == "Low")
        self.raster_quality_low_radio.Margin = Thickness(0, 0, 15, 0)
        raster_quality_radios.Children.Add(self.raster_quality_low_radio)
        
        self.raster_quality_medium_radio = RadioButton()
        self.raster_quality_medium_radio.Content = "Medium"
        self.raster_quality_medium_radio.GroupName = "RasterQuality"
        self.raster_quality_medium_radio.IsChecked = (current_raster_quality == "Medium")
        self.raster_quality_medium_radio.Margin = Thickness(0, 0, 15, 0)
        raster_quality_radios.Children.Add(self.raster_quality_medium_radio)
        
        self.raster_quality_high_radio = RadioButton()
        self.raster_quality_high_radio.Content = "High"
        self.raster_quality_high_radio.GroupName = "RasterQuality"
        self.raster_quality_high_radio.IsChecked = (current_raster_quality == "High")
        raster_quality_radios.Children.Add(self.raster_quality_high_radio)
        
        Grid.SetColumn(raster_quality_radios, 1)
        raster_quality_grid.Children.Add(raster_quality_radios)
        panel.Children.Add(raster_quality_grid)
        
        # Colors row using Grid for alignment
        colors_grid = Grid()
        colors_grid.Margin = Thickness(10, 0, 0, 3)
        colors_grid.ColumnDefinitions.Add(ColumnDefinition())
        colors_grid.ColumnDefinitions.Add(ColumnDefinition())
        colors_grid.ColumnDefinitions[0].Width = GridLength(90)
        colors_grid.ColumnDefinitions[1].Width = GridLength(1, GridUnitType.Star)
        
        colors_label = Label()
        colors_label.Content = "Colors:"
        colors_label.Padding = Thickness(0)
        colors_label.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(colors_label, 0)
        colors_grid.Children.Add(colors_label)
        
        colors_radios = StackPanel()
        colors_radios.Orientation = Orientation.Horizontal
        colors_radios.VerticalAlignment = System.Windows.VerticalAlignment.Center
        
        current_colors = self.preferences.get("DWFx_Colors", "Color")
        
        self.colors_color_radio = RadioButton()
        self.colors_color_radio.Content = "Color"
        self.colors_color_radio.GroupName = "Colors"
        self.colors_color_radio.IsChecked = (current_colors == "Color")
        self.colors_color_radio.Margin = Thickness(0, 0, 15, 0)
        colors_radios.Children.Add(self.colors_color_radio)
        
        self.colors_grayscale_radio = RadioButton()
        self.colors_grayscale_radio.Content = "Grayscale"
        self.colors_grayscale_radio.GroupName = "Colors"
        self.colors_grayscale_radio.IsChecked = (current_colors == "Grayscale")
        self.colors_grayscale_radio.Margin = Thickness(0, 0, 15, 0)
        colors_radios.Children.Add(self.colors_grayscale_radio)
        
        self.colors_bw_radio = RadioButton()
        self.colors_bw_radio.Content = "Black and White"
        self.colors_bw_radio.GroupName = "Colors"
        self.colors_bw_radio.IsChecked = (current_colors == "BlackAndWhite")
        colors_radios.Children.Add(self.colors_bw_radio)
        
        Grid.SetColumn(colors_radios, 1)
        colors_grid.Children.Add(colors_radios)
        panel.Children.Add(colors_grid)
        
        group.Content = panel
        return group
    
    def _on_graphics_format_changed(self, sender, args):
        """Handle graphics format radio button change"""
        self._update_image_quality_enabled()
    
    def _update_image_quality_enabled(self):
        """Enable/disable image quality controls based on graphics format selection"""
        is_compressed = self.graphics_compressed_radio.IsChecked
        for control in self._image_quality_controls:
            control.IsEnabled = is_compressed
    
    def _create_buttons(self):
        """Create bottom buttons"""
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        panel.HorizontalAlignment = HorizontalAlignment.Right
        panel.Margin = Thickness(0, 15, 0, 0)
        
        reset_btn = Button()
        reset_btn.Content = "Reset to Defaults"
        reset_btn.Width = 120
        reset_btn.Margin = Thickness(0, 0, 10, 0)
        reset_btn.Click += self._on_reset
        panel.Children.Add(reset_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Margin = Thickness(0, 0, 10, 0)
        cancel_btn.Click += self._on_cancel
        panel.Children.Add(cancel_btn)
        
        save_btn = Button()
        save_btn.Content = "Save"
        save_btn.Width = 80
        save_btn.IsDefault = True
        save_btn.Click += self._on_save
        panel.Children.Add(save_btn)
        
        return panel
    
    def _on_browse_folder(self, sender, args):
        """Handle browse folder button click"""
        dialog = FolderBrowserDialog()
        dialog.Description = "Select export folder"
        
        # Set initial directory if possible
        current_path = self.folder_textbox.Text
        if current_path and os.path.isabs(current_path) and os.path.exists(current_path):
            dialog.SelectedPath = current_path
        
        if dialog.ShowDialog() == DialogResult.OK:
            self.folder_textbox.Text = dialog.SelectedPath
    
    def _on_reset(self, sender, args):
        """Reset to default preferences"""
        defaults = get_default_preferences()
        
        self.folder_textbox.Text = defaults["ExportFolder"]
        self.dxf_dat_checkbox.IsChecked = defaults["DXF_CreateDatFile"]
        self.dwfx_element_data_checkbox.IsChecked = defaults["DWFx_ExportElementData"]
        self.dwfx_remove_white_checkbox.IsChecked = defaults["DWFx_RemoveOpaqueWhite"]
        
        # Graphics Settings
        self.graphics_standard_radio.IsChecked = not defaults["DWFx_UseCompressedRaster"]
        self.graphics_compressed_radio.IsChecked = defaults["DWFx_UseCompressedRaster"]
        
        image_quality = defaults["DWFx_ImageQuality"]
        self.image_quality_low_radio.IsChecked = (image_quality == "Low")
        self.image_quality_medium_radio.IsChecked = (image_quality == "Medium")
        self.image_quality_high_radio.IsChecked = (image_quality == "High")
        self._update_image_quality_enabled()
        
        # Appearance Settings
        raster_quality = defaults["DWFx_RasterQuality"]
        self.raster_quality_low_radio.IsChecked = (raster_quality == "Low")
        self.raster_quality_medium_radio.IsChecked = (raster_quality == "Medium")
        self.raster_quality_high_radio.IsChecked = (raster_quality == "High")
        
        colors = defaults["DWFx_Colors"]
        self.colors_color_radio.IsChecked = (colors == "Color")
        self.colors_grayscale_radio.IsChecked = (colors == "Grayscale")
        self.colors_bw_radio.IsChecked = (colors == "BlackAndWhite")
    
    def _on_cancel(self, sender, args):
        """Close dialog without saving"""
        self.DialogResult = False
        self.Close()
    
    def _on_save(self, sender, args):
        """Save preferences and close"""
        # Collect Graphics Settings
        image_quality = "Low"
        if self.image_quality_medium_radio.IsChecked:
            image_quality = "Medium"
        elif self.image_quality_high_radio.IsChecked:
            image_quality = "High"
        
        # Collect Appearance Settings
        raster_quality = "High"
        if self.raster_quality_low_radio.IsChecked:
            raster_quality = "Low"
        elif self.raster_quality_medium_radio.IsChecked:
            raster_quality = "Medium"
        
        colors = "Color"
        if self.colors_grayscale_radio.IsChecked:
            colors = "Grayscale"
        elif self.colors_bw_radio.IsChecked:
            colors = "BlackAndWhite"
        
        new_preferences = {
            "ExportFolder": self.folder_textbox.Text.strip(),
            "DXF_CreateDatFile": bool(self.dxf_dat_checkbox.IsChecked),
            "DWFx_ExportElementData": bool(self.dwfx_element_data_checkbox.IsChecked),
            "DWFx_RemoveOpaqueWhite": bool(self.dwfx_remove_white_checkbox.IsChecked),
            # Graphics Settings
            "DWFx_UseCompressedRaster": bool(self.graphics_compressed_radio.IsChecked),
            "DWFx_ImageQuality": image_quality,
            # Appearance Settings
            "DWFx_RasterQuality": raster_quality,
            "DWFx_Colors": colors
        }
        
        # Save to user's AppData folder (no transaction needed)
        success = set_preferences(new_preferences)
        
        if success:
            self.DialogResult = True
            self.Close()
        else:
            print("ERROR: Failed to save preferences")
            self.DialogResult = False
            self.Close()


def main():
    """Show preferences dialog"""
    try:
        window = PreferencesWindow()
        window.ShowDialog()
    except Exception as e:
        print("ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        forms.alert("Failed to open preferences dialog. See console for details.", title="Error")


if __name__ == "__main__":
    main()
