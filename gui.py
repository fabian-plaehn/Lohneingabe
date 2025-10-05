import tkinter as tk
from tkinter import messagebox
from database import Database
from utils import get_weekday_abbr

class StundenEingabeGUI:
    def __init__(self, root):
        self.root = root
        self.db = Database()
        self.setup_window()
        self.create_widgets()
        self.setup_bindings()
    
    def setup_window(self):
        """Configure main window."""
        self.root.title("Stunden-Eingabe")
        self.root.geometry("500x450")
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Jahr
        tk.Label(self.root, text="Jahr:*").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_year = tk.Entry(self.root)
        self.entry_year.grid(row=0, column=1)
        
        # Monat
        tk.Label(self.root, text="Monat:*").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_month = tk.Entry(self.root)
        self.entry_month.grid(row=1, column=1)
        
        # Tag
        self.label_day = tk.Label(self.root, text="Tag:*")
        self.label_day.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_day = tk.Entry(self.root)
        self.entry_day.grid(row=2, column=1)
        
        # Name
        tk.Label(self.root, text="Name:*").grid(row=3, column=0, sticky="e", padx=5, pady=2)
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
        
        # Required fields note
        tk.Label(self.root, text="* Pflichtfelder", font=("Arial", 8), fg="gray").grid(
            row=9, column=0, columnspan=2, sticky="w", padx=5
        )
        
        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.grid(row=10, column=0, columnspan=2, pady=20)
        
        btn_submit = tk.Button(btn_frame, text="Speichern", command=self.submit)
        btn_submit.pack(side=tk.LEFT, padx=5)
        
        btn_export = tk.Button(btn_frame, text="Excel Export", command=self.export_excel)
        btn_export.pack(side=tk.LEFT, padx=5)
        
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
            field.bind("<Down>", self.focus_next)
            field.bind("<Up>", self.focus_previous)
    
    def update_weekday(self, *args):
        """Update day label with weekday abbreviation."""
        weekday = get_weekday_abbr(
            self.entry_year.get(), 
            self.entry_month.get(), 
            self.entry_day.get()
        )
        
        if weekday:
            self.label_day.config(text=f"Tag ({weekday}):*")
        else:
            self.label_day.config(text="Tag:*")
    
    def toggle_skug(self):
        """Show/hide SKUG field based on CheckBox2."""
        if self.check_var2.get():
            self.label_skug.grid(row=7, column=0, sticky="e", padx=5, pady=2)
            self.entry_skug.grid(row=7, column=1, padx=5, pady=2)
        else:
            self.label_skug.grid_remove()
            self.entry_skug.grid_remove()
    
    def validate_required_fields(self) -> tuple[bool, str]:
        """
        Validate that all required fields are filled.
        Returns (is_valid, error_message)
        """
        jahr = self.entry_year.get().strip()
        monat = self.entry_month.get().strip()
        tag = self.entry_day.get().strip()
        name = self.entry_name.get().strip()
        
        if not jahr:
            return (False, "Jahr ist erforderlich!")
        
        if not monat:
            return (False, "Monat ist erforderlich!")
        
        if not tag:
            return (False, "Tag ist erforderlich!")
        
        if not name:
            return (False, "Name ist erforderlich!")
        
        # Validate that they are valid numbers
        try:
            jahr_int = int(jahr)
            monat_int = int(monat)
            tag_int = int(tag)
            
            if not (1900 <= jahr_int <= 2100):
                return (False, "Jahr muss zwischen 1900 und 2100 liegen!")
            
            if not (1 <= monat_int <= 12):
                return (False, "Monat muss zwischen 1 und 12 liegen!")
            
            if not (1 <= tag_int <= 31):
                return (False, "Tag muss zwischen 1 und 31 liegen!")
            
        except ValueError:
            return (False, "Jahr, Monat und Tag müssen Zahlen sein!")
        
        return (True, "")
    
    def submit(self):
        """Save entered data to database."""
        # Validate required fields
        is_valid, error_msg = self.validate_required_fields()
        
        if not is_valid:
            messagebox.showerror("Validierungsfehler", error_msg)
            return
        
        # Extract weekday from label
        weekday_text = self.label_day.cget("text")
        if "(" in weekday_text and ")" in weekday_text:
            wochentag = weekday_text.split("(")[1].split(")")[0]
        else:
            wochentag = ""
        
        data = {
            "Jahr": self.entry_year.get().strip(),
            "Monat": self.entry_month.get().strip(),
            "Tag": self.entry_day.get().strip(),
            "Name": self.entry_name.get().strip(),
            "Wochentag": wochentag,
            "Stunden": self.entry_hours.get().strip() or "0",
            "CheckBox1": bool(self.check_var1.get()),
            "CheckBox2": bool(self.check_var2.get()),
            "SKUG": self.entry_skug.get().strip() if self.check_var2.get() else "",
            "Baustelle": self.entry_bst.get().strip()
        }
        
        try:
            entry_id, was_updated = self.db.add_or_update_entry(data)
            
            if was_updated:
                messagebox.showinfo(
                    "Aktualisiert", 
                    f"Eintrag #{entry_id} wurde aktualisiert!\n"
                    f"({data['Jahr']}-{data['Monat']}-{data['Tag']}, {data['Name']})"
                )
            else:
                messagebox.showinfo(
                    "Gespeichert", 
                    f"Neuer Eintrag #{entry_id} gespeichert!"
                )
            
            # Clear fields for next entry (but keep the day)
            self.clear_fields()
            self.entry_name.focus()  # Focus on name field instead of day
        
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern:\n{str(e)}")
    
    def export_excel(self):
        """Export database to Excel."""
        try:
            if self.db.export_to_excel():
                messagebox.showinfo("Erfolg", "Daten nach Excel exportiert!")
            else:
                messagebox.showwarning("Warnung", "Keine Daten zum Exportieren vorhanden.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen:\n{str(e)}")
    
    def clear_fields(self):
        """Clear input fields after submission (except day fields)."""
        # Clear only non-day fields
        self.entry_hours.delete(0, tk.END)
        self.entry_skug.delete(0, tk.END)
        # Note: We do NOT clear entry_day, entry_month, entry_year anymore
    
    def get_visible_fields(self):
        """Get list of currently visible/mapped fields."""
        return [f for f in self.fields if f.winfo_ismapped()]
    
    def focus_next(self, event):
        """Navigate to next field on Enter/Down key."""
        widget = event.widget
        visible_fields = self.get_visible_fields()
        
        try:
            idx = visible_fields.index(widget)
            next_idx = idx + 1
            
            if next_idx < len(visible_fields):
                visible_fields[next_idx].focus()
                # Prevent default behavior (like moving cursor in entry)
                return "break"
            else:
                # No more fields, submit
                self.submit()
                return "break"
        except ValueError:
            pass
    
    def focus_previous(self, event):
        """Navigate to previous field on Up key."""
        widget = event.widget
        visible_fields = self.get_visible_fields()
        
        try:
            idx = visible_fields.index(widget)
            prev_idx = idx - 1
            
            if prev_idx >= 0:
                visible_fields[prev_idx].focus()
                # Prevent default behavior
                return "break"
        except ValueError:
            pass