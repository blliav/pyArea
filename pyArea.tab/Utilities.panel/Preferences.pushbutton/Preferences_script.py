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

from System.Windows import Window, Application, Thickness, WindowStartupLocation, GridLength, GridUnitType, HorizontalAlignment
from System.Windows.Controls import (
    StackPanel, GroupBox, Label, TextBox, CheckBox, Button,
    RadioButton, Orientation, Grid, ColumnDefinition, RowDefinition
)
from System.Windows.Input import FocusNavigationDirection
from System.Windows.Forms import FolderBrowserDialog, DialogResult

from pyrevit import revit, forms
from data_manager import get_preferences, set_preferences
from export_utils import get_default_preferences

doc = revit.doc


class PreferencesWindow(Window):
    """WPF Window for editing export preferences"""
    
    def __init__(self):
        self.Title = "pyArea Preferences"
        self.Width = 500
        self.Height = 450
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
        self.dxf_dat_checkbox.Content = "Create .dat file with DWFX scale"
        self.dxf_dat_checkbox.IsChecked = self.preferences.get("DXF_CreateDatFile", True)
        panel.Children.Add(self.dxf_dat_checkbox)
        
        group.Content = panel
        return group
    
    def _create_dwfx_section(self):
        """Create DWFX settings section"""
        group = GroupBox()
        group.Header = "DWFX Settings"
        group.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        # Export Element Data checkbox
        self.dwfx_element_data_checkbox = CheckBox()
        self.dwfx_element_data_checkbox.Content = "Export Element Data"
        self.dwfx_element_data_checkbox.IsChecked = self.preferences.get("DWFX_ExportElementData", True)
        self.dwfx_element_data_checkbox.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.dwfx_element_data_checkbox)
        
        # Remove opaque white checkbox
        self.dwfx_remove_white_checkbox = CheckBox()
        self.dwfx_remove_white_checkbox.Content = "Remove opaque white background"
        self.dwfx_remove_white_checkbox.IsChecked = self.preferences.get("DWFX_RemoveOpaqueWhite", True)
        self.dwfx_remove_white_checkbox.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.dwfx_remove_white_checkbox)
        
        # Quality radio buttons
        quality_label = Label()
        quality_label.Content = "Export Quality:"
        quality_label.Margin = Thickness(0, 0, 0, 5)
        panel.Children.Add(quality_label)
        
        quality_panel = StackPanel()
        quality_panel.Orientation = Orientation.Horizontal
        quality_panel.Margin = Thickness(10, 0, 0, 0)
        
        current_quality = self.preferences.get("DWFX_Quality", "Medium")
        
        self.quality_low_radio = RadioButton()
        self.quality_low_radio.Content = "Low"
        self.quality_low_radio.GroupName = "Quality"
        self.quality_low_radio.IsChecked = (current_quality == "Low")
        self.quality_low_radio.Margin = Thickness(0, 0, 15, 0)
        quality_panel.Children.Add(self.quality_low_radio)
        
        self.quality_medium_radio = RadioButton()
        self.quality_medium_radio.Content = "Medium"
        self.quality_medium_radio.GroupName = "Quality"
        self.quality_medium_radio.IsChecked = (current_quality == "Medium")
        self.quality_medium_radio.Margin = Thickness(0, 0, 15, 0)
        quality_panel.Children.Add(self.quality_medium_radio)
        
        self.quality_high_radio = RadioButton()
        self.quality_high_radio.Content = "High"
        self.quality_high_radio.GroupName = "Quality"
        self.quality_high_radio.IsChecked = (current_quality == "High")
        quality_panel.Children.Add(self.quality_high_radio)
        
        panel.Children.Add(quality_panel)
        group.Content = panel
        return group
    
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
        self.dwfx_element_data_checkbox.IsChecked = defaults["DWFX_ExportElementData"]
        self.dwfx_remove_white_checkbox.IsChecked = defaults["DWFX_RemoveOpaqueWhite"]
        
        quality = defaults["DWFX_Quality"]
        self.quality_low_radio.IsChecked = (quality == "Low")
        self.quality_medium_radio.IsChecked = (quality == "Medium")
        self.quality_high_radio.IsChecked = (quality == "High")
    
    def _on_cancel(self, sender, args):
        """Close dialog without saving"""
        self.DialogResult = False
        self.Close()
    
    def _on_save(self, sender, args):
        """Save preferences and close"""
        # Collect preferences from UI
        quality = "Medium"
        if self.quality_low_radio.IsChecked:
            quality = "Low"
        elif self.quality_high_radio.IsChecked:
            quality = "High"
        
        new_preferences = {
            "ExportFolder": self.folder_textbox.Text.strip(),
            "DXF_CreateDatFile": bool(self.dxf_dat_checkbox.IsChecked),
            "DWFX_ExportElementData": bool(self.dwfx_element_data_checkbox.IsChecked),
            "DWFX_Quality": quality,
            "DWFX_RemoveOpaqueWhite": bool(self.dwfx_remove_white_checkbox.IsChecked)
        }
        
        # Save to project
        with revit.Transaction("Save Preferences"):
            success = set_preferences(new_preferences)
        
        if success:
            forms.alert("Preferences saved successfully!", title="Success")
            self.DialogResult = True
            self.Close()
        else:
            forms.alert("Failed to save preferences. Check console for errors.", title="Error")


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
