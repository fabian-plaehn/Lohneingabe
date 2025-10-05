import tkinter as tk
from tkinter import messagebox
from data_handler import DataHandler
from utils import get_weekday_abbr

class StundenEingabeGUI:
    def __init__(self, root):
        self.root = root
        self.data_handler = DataHandler()
        self.setup_window()
        self.create_widgets()
        self.setup_bindings()
    
    def setup_window(self):
        """Configure main window."""
        self.root.title("Stunden-Eingabe")
        self.root.geometry("500x400")
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Jahr
        tk.Label(self.root, text="Jahr:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_year = tk.Entry(self.root)
        self.entry_year.grid(row=0, column=1)
        
        # Monat
        tk.Label(self.root, text="Monat:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_month = tk.Entry(self.root)
        self.entry_month.grid(row=1, column=1)
        
        # Tag
        self.label_day = tk.Label(self.root, text="Tag:")
        self.label_day.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_day = tk.Entry(self.root)
        self.entry_day.grid(row=2, column=1)
        
        # Name
        tk.Label(self.root, text="Name:").grid(row=3, column=0, sticky="e", padx=5, pady=2)
        self.entry_name = tk.Entry(self.root)
        self.entry_name.grid(row=3, column=1)
        
        # Stunden
        tk.Label(self.root, text="Stunden:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_hours = tk.Entry(self.root)
        self.entry_hours.grid(row=4, column=1)
        
        # CheckBox1
        self.check_var1 = tk.IntVar()
        check1 = tk.Checkbutton(self.root, text="CheckBox1", variable=self.check_var1)
        check1.grid(row=5, column=0, columnspan=2, pady=2)
        
        # CheckBox2
        self.check_var2 = tk.IntVar()
        check2 = tk.Checkbutton(self.root, text="CheckBox2", variable=self.check_var2, 
                                command=self.toggle_skug)
        check2.grid(row=6, column=0, columnspan=2, pady=2)
        
        # SKUG Feld (nur sichtbar wenn CheckBox2 aktiv)
        self.label_skug = tk.Label(self.root, text="SKUG:")
        self.entry_skug = tk.Entry(self.root)
        
        # Baustelle
        tk.Label(self.root, text="Baustelle:").grid(row=8, column=0, sticky="e", padx=5, pady=2)
        self.entry_bst = tk.Entry(self.root)
        self.entry_bst.grid(row=8, column=1)
        
        # Button
        btn_submit = tk.Button(self.root, text="Speichern", command=self.submit)
        btn_submit.grid(row=9, column=0, columnspan=2, pady=20)
        
        # Field list for navigation
        self.fields = [
            self.entry_year, self.entry_month, self.entry_day, 
            self.entry_name, self.entry_hours, self.entry_skug, self.entry_bst
        ]
    
    def setup_bindings(self):
        """Setup event bindings."""
        # Update weekday on date field changes
        self.entry_year.bind("<KeyRelease>", self.update_weekday)
        self.entry_month.bind("<KeyRelease>", self.update_weekday)
        self.entry_day.bind("<KeyRelease>", self.update_weekday)
        
        # Enter key navigation
        for field in self.fields:
            field.bind("<Return>", self.focus_next)
    
    def update_weekday(self, *args):
        """Update day label with weekday abbreviation."""
        weekday = get_weekday_abbr(
            self.entry_year.get(), 
            self.entry_month.get(), 
            self.entry_day.get()
        )
        
        if weekday:
            self.label_day.config(text=f"Tag ({weekday}):")
        else:
            self.label_day.config(text="Tag:")
    
    def toggle_skug(self):
        """Show/hide SKUG field based on CheckBox2."""
        if self.check_var2.get():
            self.label_skug.grid(row=7, column=0, sticky="e", padx=5, pady=2)
            self.entry_skug.grid(row=7, column=1, padx=5, pady=2)
        else:
            self.label_skug.grid_remove()
            self.entry_skug.grid_remove()
    
    def submit(self):
        """Save entered data and export to Excel."""
        data = {
            "Jahr": self.entry_year.get(),
            "Monat": self.entry_month.get(),
            "Tag": self.entry_day.get(),
            "Wochentag": self.label_day.cget("text").replace("Tag (", "").replace("):", ""),
            "Name": self.entry_name.get(),
            "Stunden": self.entry_hours.get(),
            "CheckBox1": bool(self.check_var1.get()),
            "CheckBox2": bool(self.check_var2.get()),
            "SKUG": self.entry_skug.get() if self.check_var2.get() else "",
            "Baustelle": self.entry_bst.get()
        }
        
        self.data_handler.add_entry(data)
        self.data_handler.save_to_excel()
        
        messagebox.showinfo("Eingabe", f"Daten gespeichert:\n{data}")
        
        # Clear fields for next entry
        self.clear_fields()
        self.entry_day.focus()
    
    def clear_fields(self):
        """Clear input fields after submission."""
        self.entry_hours.delete(0, tk.END)
        self.entry_skug.delete(0, tk.END)
        self.entry_day.delete(0, tk.END)
        self.label_day.config(text="Tag:")
    
    def focus_next(self, event):
        """Navigate to next field on Enter key."""
        widget = event.widget
        try:
            idx = self.fields.index(widget)
            next_idx = idx + 1
            
            # Skip hidden or non-mapped fields
            while next_idx < len(self.fields):
                if self.fields[next_idx].winfo_ismapped():
                    self.fields[next_idx].focus()
                    break
                next_idx += 1
            else:
                # No more fields, submit
                self.submit()
        except ValueError:
            pass