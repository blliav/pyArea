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
        
        # Create TextBlock inside color swatch to show number
        self._swatch_text = None
        if self.color_swatch:
            from System.Windows.Controls import TextBlock
            from System.Windows import VerticalAlignment, HorizontalAlignment
            self._swatch_text = TextBlock()
            self._swatch_text.FontSize = 11
            self._swatch_text.FontWeight = System.Windows.FontWeights.Bold
            self._swatch_text.Foreground = System.Windows.Media.Brushes.White
            self._swatch_text.VerticalAlignment = VerticalAlignment.Center
            self._swatch_text.HorizontalAlignment = HorizontalAlignment.Center
            self.color_swatch.Child = self._swatch_text
        
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
        
        # Store initial value to restore if user doesn't select anything valid
        self._initial_value = None
        
        # Wire up events for filtering
        self.combo_control.AddHandler(
            ComboBox.KeyUpEvent,
            KeyEventHandler(self._on_key_up)
        )
        # Use PreviewKeyDown for Enter key to intercept before ComboBox handles it
        self.combo_control.PreviewKeyDown += self._on_preview_key_down
        self.combo_control.DropDownOpened += self._on_dropdown_opened
        self.combo_control.DropDownClosed += self._on_dropdown_closed
        self.combo_control.SelectionChanged += self._on_selection_changed
        # Fix first keystroke: prevent auto-select-all on focus
        self.combo_control.GotKeyboardFocus += self._on_got_keyboard_focus
        # Restore initial value if user doesn't select a valid item
        self.combo_control.LostFocus += self._on_lost_focus
    
    def _extract_number_from_text(self, text):
        """Extract the number from usage type text (format: '1. name' or 'Not defined')"""
        if not text or text == "Not defined":
            return ""
        # Split on '. ' and take the first part
        parts = text.split('. ', 1)
        if len(parts) >= 1:
            return parts[0].strip()
        return ""
    
    def _update_swatch(self, colored_item):
        """Update color swatch background and text"""
        if not self.color_swatch:
            return
        
        if colored_item and colored_item.color_rgb:
            r, g, b = colored_item.color_rgb
            self.color_swatch.Background = SolidColorBrush(Color.FromRgb(r, g, b))
            # Show number in swatch
            if self._swatch_text:
                number = self._extract_number_from_text(colored_item.text)
                self._swatch_text.Text = number
        else:
            # No color - make swatch transparent/empty
            from System.Windows.Media import Brushes
            self.color_swatch.Background = Brushes.Transparent
            if self._swatch_text:
                self._swatch_text.Text = ""
    
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
            
            # Set TextSearch property so ComboBox displays text instead of "System.Windows.Controls.StackPanel"
            from System.Windows.Controls import TextSearch
            TextSearch.SetText(stack_panel, item.text)
            
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
    
    
    def _on_preview_key_down(self, sender, args):
        """Handle Enter key before ComboBox processes it"""
        from System.Windows.Input import Key
        
        # Handle Enter key - select highlighted item and close dropdown
        if args.Key == Key.Enter:
            if self.combo_control.IsDropDownOpen:
                # If an item is highlighted (SelectedIndex >= 0), select it
                if self.combo_control.SelectedIndex >= 0:
                    selected_item = self.combo_control.SelectedItem
                    if selected_item is not None and isinstance(selected_item, StackPanel):
                        if hasattr(selected_item, 'Tag') and selected_item.Tag is not None:
                            colored_item = selected_item.Tag
                            # Store the selected item - this is critical for get_text() to work
                            self._selected_colored_item = colored_item
                            
                            # Reset filter to show all options before closing
                            self._updating_selection = True
                            try:
                                self.populate(self.original_options)
                                
                                # Set the text
                                self.combo_control.Text = colored_item.text
                                
                                # Update swatch with color and number
                                self._update_swatch(colored_item)
                                
                                # Find and select the item in the full list
                                for i, stack_panel in enumerate(self.combo_control.Items):
                                    if isinstance(stack_panel, StackPanel) and hasattr(stack_panel, 'Tag'):
                                        if stack_panel.Tag == colored_item:
                                            self.combo_control.SelectedIndex = i
                                            break
                            finally:
                                self._updating_selection = False
                            
                            # Mark event as handled to prevent ComboBox from processing it
                            args.Handled = True
                # Close dropdown (this will trigger _on_dropdown_closed which will handle final cleanup)
                self.combo_control.IsDropDownOpen = False
    
    def _on_key_up(self, sender, args):
        """Filter ComboBox items based on typed text"""
        from System.Windows.Input import Key
        
        # Don't filter on certain keys (arrow keys, enter, escape, tab)
        # Enter is handled in _on_preview_key_down - must skip here to avoid reopening dropdown
        if args.Key in [Key.Down, Key.Up, Key.Enter, Key.Escape, Key.Tab]:
            return
        
        # Store current text
        current_text = self.combo_control.Text
        search_text = current_text.lower() if current_text else ""
        
        # Clear selected item when user types (so get_text() returns typed text)
        self._selected_colored_item = None
        
        if search_text == "":
            # Show all options if text is empty
            filtered = self.original_options
        else:
            # Filter options that contain the search text
            filtered = [opt for opt in self.original_options 
                       if search_text in opt.text.lower()]
        
        # Repopulate with filtered items
        self.populate(filtered)
        
        # Restore text if it was cleared by populate
        if self.combo_control.Text != current_text:
            self.combo_control.Text = current_text
        
        # Open dropdown if there are filtered results
        if len(filtered) > 0:
            self.combo_control.IsDropDownOpen = True
    
    def _on_got_keyboard_focus(self, sender, args):
        """Clear text and place cursor at beginning when user focuses on ComboBox"""
        try:
            # Store initial value before clearing (for restore on LostFocus)
            # Only store if not already empty (to avoid overwriting on second focus event)
            current_text = self.combo_control.Text
            if current_text:  # Only store non-empty values
                self._initial_value = current_text
                self._initial_selected_item = self._selected_colored_item
            
            # Clear the text and reset selection
            self.combo_control.Text = ""
            self._selected_colored_item = None
            
            # Find the internal editable TextBox
            textbox = self.combo_control.Template.FindName("PART_EditableTextBox", self.combo_control)
            if textbox is not None:
                # Set cursor to beginning
                textbox.SelectionStart = 0
                textbox.SelectionLength = 0
                
                # Also attach to SelectionChanged to prevent select-all after first keystroke
                def on_textbox_selection_changed(tb_sender, tb_args):
                    try:
                        # If all text is selected (common WPF behavior), clear selection
                        if tb_sender.SelectionLength > 0 and tb_sender.SelectionStart == 0:
                            text_len = len(tb_sender.Text) if tb_sender.Text else 0
                            if tb_sender.SelectionLength == text_len and text_len > 0:
                                tb_sender.SelectionStart = text_len
                                tb_sender.SelectionLength = 0
                    except:
                        pass
                
                # Attach the handler (note: this will accumulate handlers, but that's OK for now)
                textbox.SelectionChanged += on_textbox_selection_changed
        except Exception as e:
            pass
    
    def _on_lost_focus(self, sender, args):
        """Restore initial value if user didn't select a valid item"""
        current_text = self.combo_control.Text
        
        # Check if user selected an item from the dropdown
        if self._selected_colored_item is not None:
            # Clear initial value since user made a valid selection
            self._initial_value = None
            self._initial_selected_item = None
            return
        
        # Check if typed text matches any option exactly
        if current_text:
            for option in self.all_options:
                if option.text.lower() == current_text.lower():
                    # Found exact match - keep it
                    self._selected_colored_item = option
                    self.combo_control.Text = option.text
                    # Update swatch with color and number
                    self._update_swatch(option)
                    # Clear initial value since user made a valid selection
                    self._initial_value = None
                    self._initial_selected_item = None
                    return
        
        # No valid selection - restore initial value
        self._updating_selection = True
        try:
            self.combo_control.Text = self._initial_value if self._initial_value else ""
            self._selected_colored_item = self._initial_selected_item if hasattr(self, '_initial_selected_item') else None
            
            # Update color swatch with color and number
            self._update_swatch(self._selected_colored_item)
            
            # Clear initial value after restoring so next focus cycle works correctly
            self._initial_value = None
            self._initial_selected_item = None
        finally:
            self._updating_selection = False
        
        # Always reset filter to show all options when focus is lost
        # Store current text and selection before resetting
        final_text = self.combo_control.Text
        final_selected = self._selected_colored_item
        
        self._updating_selection = True
        try:
            self.populate(self.original_options)
            
            # Restore the text after repopulating
            if final_text:
                self.combo_control.Text = final_text
            
            # Find and select the current item in the repopulated list
            if final_selected is not None:
                for i, stack_panel in enumerate(self.combo_control.Items):
                    if isinstance(stack_panel, StackPanel) and hasattr(stack_panel, 'Tag'):
                        if stack_panel.Tag == final_selected:
                            self.combo_control.SelectedIndex = i
                            break
        finally:
            self._updating_selection = False
    
    def _on_dropdown_opened(self, sender, args):
        """Reset filter when dropdown is opened"""
        if self.combo_control.Text == "":
            self.populate(self.original_options)
    
    def _on_dropdown_closed(self, sender, args):
        """Set the text after dropdown closes and reset filter"""
        # Store current text and selection before resetting filter
        current_text = ""
        selected_item = self._selected_colored_item
        
        if selected_item is not None:
            self._updating_selection = True
            try:
                self.combo_control.Text = selected_item.text
                current_text = selected_item.text
            finally:
                self._updating_selection = False
        
        # Reset filter to show all options after dropdown closes
        self._updating_selection = True
        try:
            self.populate(self.original_options)
            
            # Restore the text after repopulating (populate clears it)
            if current_text:
                self.combo_control.Text = current_text
            
            # Find and select the current item in the repopulated list to maintain position
            if selected_item is not None:
                for i, stack_panel in enumerate(self.combo_control.Items):
                    if isinstance(stack_panel, StackPanel) and hasattr(stack_panel, 'Tag'):
                        if stack_panel.Tag == selected_item:
                            self.combo_control.SelectedIndex = i
                            break
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
                
                # Update color swatch with color and number
                self._update_swatch(colored_item)
                
                # Note: Text will be set when dropdown closes to avoid display issues during navigation
                # Filter will be reset when dropdown closes or focus is lost
    
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
                # Found it - set the text and color swatch with number
                self._selected_colored_item = item
                self.combo_control.Text = item.text
                self._update_swatch(item)
                return
        
        # Not found in options - just set the text as-is
        self.combo_control.Text = value
        self._selected_colored_item = None
        self._update_swatch(None)
