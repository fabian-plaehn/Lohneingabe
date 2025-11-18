import tkinter as tk
from tkinter import ttk


class AutocompleteEntry:
    """Autocomplete functionality for Entry widgets."""

    def __init__(self, entry_widget, suggestions_callback):
        """
        Initialize autocomplete for an entry widget.

        Args:
            entry_widget: The tk.Entry widget to add autocomplete to
            suggestions_callback: Function that returns list of suggestions
        """
        self.entry = entry_widget
        self.suggestions_callback = suggestions_callback
        self.listbox = None
        self.listbox_window = None
        self.listbox_frame = None

        # Bind events - bind navigation keys FIRST so they take priority
        self.entry.bind("<Down>", self.on_down)
        self.entry.bind("<Up>", self.on_up)
        self.entry.bind("<Return>", self.on_return)
        self.entry.bind("<KeyRelease>", self.on_key_release, add="+")
        self.entry.bind("<FocusOut>", self.on_focus_out, add="+")
        self.entry.bind("<Tab>", self.on_tab, add="+")
        self.entry.bind("<Escape>", lambda e: self.hide_listbox(), add="+")

    def on_key_release(self, event):
        """Handle key release event."""
        # Ignore navigation keys
        if event.keysym in ('Down', 'Up', 'Return', 'Escape', 'Tab'):
            return

        full_value = self.entry.get()

        if not full_value:
            self.hide_listbox()
            return

        # Get the current word being typed (after last comma)
        cursor_pos = self.entry.index(tk.INSERT)
        text_before_cursor = full_value[:cursor_pos]

        # Find the last comma before cursor
        last_comma_pos = text_before_cursor.rfind(',')

        if last_comma_pos >= 0:
            # Get text after last comma
            current_word = text_before_cursor[last_comma_pos + 1:].strip()
        else:
            current_word = text_before_cursor.strip()

        if not current_word:
            self.hide_listbox()
            return

        # Get suggestions
        suggestions = self.get_suggestions(current_word)

        if suggestions:
            self.show_listbox(suggestions)
        else:
            self.hide_listbox()

    def get_suggestions(self, value):
        """Get filtered suggestions based on current value."""
        all_suggestions = self.suggestions_callback()
        value_lower = value.lower()

        # Filter suggestions that start with or contain the value
        filtered = [s for s in all_suggestions if s.lower().startswith(value_lower)]

        return filtered[:10]  # Limit to 10 suggestions

    def show_listbox(self, suggestions):
        """Show listbox with suggestions."""
        if not self.listbox_window:
            # Create listbox window
            self.listbox_window = tk.Toplevel(self.entry.winfo_toplevel())
            self.listbox_window.wm_overrideredirect(True)

            # Frame to add border
            self.listbox_frame = tk.Frame(self.listbox_window, relief=tk.SOLID, borderwidth=1)
            self.listbox_frame.pack(fill=tk.BOTH, expand=True)

            self.listbox = tk.Listbox(
                self.listbox_frame,
                height=min(len(suggestions), 10),
                relief=tk.FLAT,
                selectmode=tk.SINGLE,
                exportselection=False
            )
            self.listbox.pack(fill=tk.BOTH, expand=True)

            self.listbox.bind("<ButtonRelease-1>", self.on_click)
            self.listbox.bind("<Return>", self.on_select)

        # Clear and populate
        self.listbox.delete(0, tk.END)
        for suggestion in suggestions:
            self.listbox.insert(tk.END, suggestion)

        # Auto-select first item
        if self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)

        # Update window and get accurate positioning
        self.entry.update_idletasks()
        self.listbox_window.update_idletasks()

        # Position below entry
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = self.entry.winfo_width()
        height = min(len(suggestions) * 20 + 4, 204)  # +4 for border

        self.listbox_window.geometry(f"{width}x{height}+{x}+{y}")
        self.listbox_window.lift()

    def hide_listbox(self, event=None):
        """Hide the listbox."""
        if self.listbox_window:
            self.listbox_window.destroy()
            self.listbox_window = None
            self.listbox = None
            self.listbox_frame = None

    def on_focus_out(self, event):
        """Handle focus out with delay to allow click."""
        # Delay hiding to allow click event to register
        if self.listbox_window:
            self.entry.after(150, self.hide_listbox)

    def on_down(self, event):
        """Handle down arrow key."""
        if self.listbox and self.listbox.winfo_exists():
            current = self.listbox.curselection()
            if current:
                index = current[0]
                if index < self.listbox.size() - 1:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index + 1)
                    self.listbox.activate(index + 1)
                    self.listbox.see(index + 1)
            else:
                self.listbox.selection_set(0)
                self.listbox.activate(0)
            return "break"
        # Allow default behavior when no listbox
        return None

    def on_up(self, event):
        """Handle up arrow key."""
        if self.listbox and self.listbox.winfo_exists():
            current = self.listbox.curselection()
            if current:
                index = current[0]
                if index > 0:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index - 1)
                    self.listbox.activate(index - 1)
                    self.listbox.see(index - 1)
            return "break"
        # Allow default behavior when no listbox
        return None

    def on_return(self, event):
        """Handle return key."""
        if self.listbox and self.listbox.winfo_exists() and self.listbox.curselection():
            self.on_select(event)
            self.hide_listbox()
            return "break"
        # Allow default behavior when no listbox (navigate to next field)
        return None

    def on_tab(self, event):
        """Handle tab key."""
        if self.listbox and self.listbox.size() > 0:
            self.on_select(event)
            self.hide_listbox()
            return "break"

    def on_click(self, event):
        """Handle mouse click."""
        self.on_select(event)
        self.hide_listbox()

    def on_select(self, event=None):
        """Handle selection from listbox."""
        if self.listbox:
            selection = self.listbox.curselection()
            if selection:
                selected_value = self.listbox.get(selection[0])

                # Get current entry content and cursor position
                full_value = self.entry.get()
                cursor_pos = self.entry.index(tk.INSERT)
                text_before_cursor = full_value[:cursor_pos]
                text_after_cursor = full_value[cursor_pos:]

                # Find the last comma before cursor
                last_comma_pos = text_before_cursor.rfind(',')

                if last_comma_pos >= 0:
                    # Replace text after last comma
                    new_text = text_before_cursor[:last_comma_pos + 1] + " " + selected_value + text_after_cursor
                    new_cursor_pos = last_comma_pos + 2 + len(selected_value)
                else:
                    # Replace entire text before cursor
                    new_text = selected_value + text_after_cursor
                    new_cursor_pos = len(selected_value)

                self.entry.delete(0, tk.END)
                self.entry.insert(0, new_text)
                self.entry.icursor(new_cursor_pos)


class BaustelleAutocomplete:
    """Autocomplete for Baustelle field with special formatting."""

    def __init__(self, entry_widget, baustellen_callback):
        """
        Initialize baustelle autocomplete.

        Args:
            entry_widget: The tk.Entry widget to add autocomplete to
            baustellen_callback: Function that returns list of baustelle dicts
        """
        self.entry = entry_widget
        self.baustellen_callback = baustellen_callback
        self.listbox = None
        self.listbox_window = None
        self.listbox_frame = None
        self.baustellen_data = []

        # Bind events - bind navigation keys FIRST so they take priority
        self.entry.bind("<Down>", self.on_down)
        self.entry.bind("<Up>", self.on_up)
        self.entry.bind("<Return>", self.on_return)
        self.entry.bind("<KeyRelease>", self.on_key_release, add="+")
        self.entry.bind("<FocusOut>", self.on_focus_out, add="+")
        self.entry.bind("<Tab>", self.on_tab, add="+")
        self.entry.bind("<Escape>", lambda e: self.hide_listbox(), add="+")

    def on_key_release(self, event):
        """Handle key release event."""
        if event.keysym in ('Down', 'Up', 'Return', 'Escape', 'Tab'):
            return

        value = self.entry.get()

        if not value:
            self.hide_listbox()
            return

        suggestions = self.get_suggestions(value)

        if suggestions:
            self.show_listbox(suggestions)
        else:
            self.hide_listbox()

    def get_suggestions(self, value):
        """Get filtered baustelle suggestions."""
        all_baustellen = self.baustellen_callback()
        value_lower = value.lower()

        # Filter baustellen where nummer or name contains the value
        filtered = []
        for b in all_baustellen:
            display_text = f"{b['nummer']} - {b['name']}"
            if value_lower in b['nummer'].lower() or value_lower in b['name'].lower():
                filtered.append({'display': display_text, 'data': b})

        return filtered[:10]

    def show_listbox(self, suggestions):
        """Show listbox with suggestions."""
        if not self.listbox_window:
            self.listbox_window = tk.Toplevel(self.entry.winfo_toplevel())
            self.listbox_window.wm_overrideredirect(True)

            # Frame to add border
            self.listbox_frame = tk.Frame(self.listbox_window, relief=tk.SOLID, borderwidth=1)
            self.listbox_frame.pack(fill=tk.BOTH, expand=True)

            self.listbox = tk.Listbox(
                self.listbox_frame,
                height=min(len(suggestions), 10),
                relief=tk.FLAT,
                selectmode=tk.SINGLE,
                exportselection=False
            )
            self.listbox.pack(fill=tk.BOTH, expand=True)

            self.listbox.bind("<ButtonRelease-1>", self.on_click)
            self.listbox.bind("<Return>", self.on_select)

        # Clear and populate
        self.listbox.delete(0, tk.END)
        self.baustellen_data = []

        for suggestion in suggestions:
            self.listbox.insert(tk.END, suggestion['display'])
            self.baustellen_data.append(suggestion['data'])

        # Auto-select first item
        if self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)

        # Update window and get accurate positioning
        self.entry.update_idletasks()
        self.listbox_window.update_idletasks()

        # Position below entry
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = max(self.entry.winfo_width(), 300)
        height = min(len(suggestions) * 20 + 4, 204)  # +4 for border

        self.listbox_window.geometry(f"{width}x{height}+{x}+{y}")
        self.listbox_window.lift()

    def hide_listbox(self, event=None):
        """Hide the listbox."""
        if self.listbox_window:
            self.listbox_window.destroy()
            self.listbox_window = None
            self.listbox = None
            self.listbox_frame = None
            self.baustellen_data = []

    def on_focus_out(self, event):
        """Handle focus out with delay to allow click."""
        # Delay hiding to allow click event to register
        if self.listbox_window:
            self.entry.after(150, self.hide_listbox)

    def on_down(self, event):
        """Handle down arrow key."""
        if self.listbox and self.listbox.winfo_exists():
            current = self.listbox.curselection()
            if current:
                index = current[0]
                if index < self.listbox.size() - 1:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index + 1)
                    self.listbox.activate(index + 1)
                    self.listbox.see(index + 1)
            else:
                self.listbox.selection_set(0)
                self.listbox.activate(0)
            return "break"
        # Allow default behavior when no listbox
        return None

    def on_up(self, event):
        """Handle up arrow key."""
        if self.listbox and self.listbox.winfo_exists():
            current = self.listbox.curselection()
            if current:
                index = current[0]
                if index > 0:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index - 1)
                    self.listbox.activate(index - 1)
                    self.listbox.see(index - 1)
            return "break"
        # Allow default behavior when no listbox
        return None

    def on_return(self, event):
        """Handle return key."""
        if self.listbox and self.listbox.winfo_exists() and self.listbox.curselection():
            self.on_select(event)
            self.hide_listbox()
            return "break"
        # Allow default behavior when no listbox (navigate to next field)
        return None

    def on_tab(self, event):
        """Handle tab key."""
        if self.listbox and self.listbox.size() > 0:
            self.on_select(event)
            self.hide_listbox()
            return "break"

    def on_click(self, event):
        """Handle mouse click."""
        self.on_select(event)
        self.hide_listbox()

    def on_select(self, event=None):
        """Handle selection from listbox."""
        if self.listbox:
            selection = self.listbox.curselection()
            if selection:
                value = self.listbox.get(selection[0])
                self.entry.delete(0, tk.END)
                self.entry.insert(0, value)
                self.entry.icursor(tk.END)
