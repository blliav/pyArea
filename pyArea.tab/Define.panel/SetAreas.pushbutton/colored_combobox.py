# -*- coding: utf-8 -*-
"""Reusable Colored ComboBox with HTML Color Tag Support"""

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System.Windows.Controls import ComboBox, TextBlock, StackPanel
from System.Windows.Markup import XamlReader
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Documents import Run
import System


class ColoredComboItem(object):
    """Class to hold combo box item with text and optional color"""
    def __init__(self, text, color_rgb=None):
        self.text = text                   # Text shown and returned
        self.color_rgb = color_rgb         # Optional RGB tuple (R, G, B) for colored square
    
    def __str__(self):
        return self.text


class ColoredComboBox(object):
    """
    A WPF ComboBox wrapper with optional colored squares in dropdown.
    
    Features:
    - Optional colored squares for items (shown only in dropdown)
    - Filterable/searchable dropdown
    - Editable text input
    - Color swatch support
    - Optional RTL (Right-to-Left) support for Hebrew
    
    Usage:
        options = [
            ("Private Office", (74, 144, 226)),  # With color
            "Corridor"                            # Without color
        ]
        combo = ColoredComboBox(parent_combo_control, options, color_swatch_element, rtl=False)
        combo.populate()
        selected = combo.get_text()  # Returns "Private Office"
    """
    
    def __init__(self, combo_control, options, color_swatch=None, rtl=False):
        """
        Initialize the colored combobox.
        
        Args:
            combo_control: WPF ComboBox control from XAML
            options: List of items, where each item can be:
                     - String: text without color
                     - Tuple: (text, (R, G, B)) text with color
            color_swatch: Optional WPF element (Border, Rectangle) to update with selection color
            rtl: If True, set Right-to-Left flow direction for Hebrew/Arabic
                     
        Example:
            options = [
                ("Private Office", (74, 144, 226)),  # With blue color
                "Corridor"                            # Without color
            ]
        """
        self.combo_control = combo_control
        self.color_swatch = color_swatch
        self.rtl = rtl
        
        # Set RTL flow direction if requested
        if rtl:
            from System.Windows import FlowDirection
            self.combo_control.FlowDirection = FlowDirection.RightToLeft
        
        # Create ColoredComboItem objects
        self.all_options = []
        
        for option in options:
            if isinstance(option, tuple):
                # Tuple: (text, color_rgb)
                text, color_rgb = option
                item = ColoredComboItem(text, color_rgb)
            else:
                # Simple string
                item = ColoredComboItem(option, None)
            self.all_options.append(item)
        
        # Store original options
        self.original_options = list(self.all_options)
        
        # Track the selected item for returning the actual value
        self._selected_colored_item = None
        
        # Flag to prevent recursive event handling
        self._updating_selection = False
        
        # Wire up events for filtering
        self.combo_control.AddHandler(
            ComboBox.KeyUpEvent,
            KeyEventHandler(self._on_key_up)
        )
        self.combo_control.DropDownOpened += self._on_dropdown_opened
        self.combo_control.DropDownClosed += self._on_dropdown_closed
        self.combo_control.SelectionChanged += self._on_selection_changed
    
    def populate(self, items=None):
        """
        Populate ComboBox with items.
        
        Args:
            items: Optional list of ColoredComboItem objects. If None, uses all_options.
        """
        if items is None:
            items = self.all_options
        
        self.combo_control.Items.Clear()
        
        for item in items:
            # Create StackPanel to hold colored square (if any) and text
            stack_panel = StackPanel()
            stack_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
            stack_panel.Tag = item  # Store the ColoredComboItem for later retrieval
            
            # Add colored square if color is specified
            if item.color_rgb:
                square = TextBlock()
                square.Text = u"\u25A0"  # Solid square character
                square.FontSize = self.combo_control.FontSize
                square.Margin = System.Windows.Thickness(0, 0, 5, 0)
                # Set color from RGB tuple
                r, g, b = item.color_rgb
                square.Foreground = SolidColorBrush(Color.FromRgb(r, g, b))
                stack_panel.Children.Add(square)
            
            # Add text
            text_block = TextBlock()
            text_block.Text = item.text
            text_block.FontSize = self.combo_control.FontSize
            stack_panel.Children.Add(text_block)
            
            self.combo_control.Items.Add(stack_panel)
    
    
    def _on_key_up(self, sender, args):
        """Filter ComboBox items based on typed text"""
        # Store current text
        current_text = self.combo_control.Text
        search_text = current_text.lower() if current_text else ""
        
        # Clear selected item when user types (so get_text() returns typed text)
        self._selected_colored_item = None
        
        # Don't filter on certain keys (arrow keys, enter, etc.)
        from System.Windows.Input import Key
        if args.Key in [Key.Down, Key.Up, Key.Enter, Key.Escape, Key.Tab]:
            return
        
        if search_text == "":
            # Show all options if text is empty
            filtered = self.original_options
        else:
            # Filter options that contain the search text
            filtered = [opt for opt in self.original_options 
                       if search_text in opt.text.lower()]
        
        # Repopulate with filtered items
        self.populate(filtered)
        
        # Open dropdown if there are filtered results
        if len(filtered) > 0:
            self.combo_control.IsDropDownOpen = True
    
    def _on_dropdown_opened(self, sender, args):
        """Reset filter when dropdown is opened"""
        if self.combo_control.Text == "":
            self.populate(self.original_options)
    
    def _on_dropdown_closed(self, sender, args):
        """Set the text after dropdown closes"""
        if self._selected_colored_item is not None:
            self._updating_selection = True
            try:
                self.combo_control.Text = self._selected_colored_item.text
            finally:
                self._updating_selection = False
    
    def _on_selection_changed(self, sender, args):
        """Handle item selection to display text correctly and update color swatch"""
        # Prevent recursive calls
        if self._updating_selection:
            return
            
        selected_item = self.combo_control.SelectedItem
        
        if selected_item is not None and isinstance(selected_item, StackPanel):
            # Extract the ColoredComboItem from the Tag
            if hasattr(selected_item, 'Tag') and selected_item.Tag is not None:
                colored_item = selected_item.Tag
                
                # Store the selected item for get_text() to retrieve the actual value
                self._selected_colored_item = colored_item
                
                # Update color swatch if provided
                if self.color_swatch and colored_item.color_rgb:
                    r, g, b = colored_item.color_rgb
                    self.color_swatch.Background = SolidColorBrush(Color.FromRgb(r, g, b))
                elif self.color_swatch:
                    # No color - make swatch transparent/empty
                    from System.Windows.Media import Brushes
                    self.color_swatch.Background = Brushes.Transparent
                
                # Text will be set when dropdown closes (_on_dropdown_closed)
    
    def get_text(self):
        """
        Get the text from the combobox.
        If an item was selected from dropdown, returns that item's text.
        If user typed text, returns the typed text.
        """
        if self._selected_colored_item is not None:
            # Return the selected item's text
            return self._selected_colored_item.text
        else:
            # Return what the user typed (or None if empty)
            return self.combo_control.Text if self.combo_control.Text else None
    
    def set_text(self, text):
        """Set the text value in the combobox"""
        self.combo_control.Text = text
    
    def set_initial_value(self, value):
        """
        Set the initial value in the combobox based on a parameter value.
        Looks for matching text in options, and if found, sets the text and color.
        If not found, just sets the text as-is.
        
        Args:
            value: The text value to set
        """
        if not value:
            return
        
        # Look for the value in our options
        for item in self.all_options:
            if item.text == value:
                # Found it - set the text and color swatch
                self._selected_colored_item = item
                self.combo_control.Text = item.text
                
                # Update color swatch if provided
                if self.color_swatch and item.color_rgb:
                    r, g, b = item.color_rgb
                    self.color_swatch.Background = SolidColorBrush(Color.FromRgb(r, g, b))
                elif self.color_swatch:
                    from System.Windows.Media import Brushes
                    self.color_swatch.Background = Brushes.Transparent
                
                return
        
        # Not found in options - just set the text as-is
        self.combo_control.Text = value
        self._selected_colored_item = None
        
        # Clear color swatch
        if self.color_swatch:
            from System.Windows.Media import Brushes
            self.color_swatch.Background = Brushes.Transparent
