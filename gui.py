import tkinter as tk
from tkinter import ttk, messagebox
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
        self.root.geometry("1000x600")
    
    def create_widgets(self):
        """Create all GUI widgets."""
        # Create main container with two columns
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left side - Input form
        input_frame = tk.Frame(main_frame)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Right side - Data displays
        display_frame = tk.Frame(main_frame)
        display_frame.grid(row=0, column=1, sticky="nsew")
        
        # Configure grid weights
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_rowconfigure(0, weight=1)
        
        # --- INPUT FORM ---
        self.create_input_fields(input_frame)
        
        # --- DATA DISPLAYS ---
        self.create_data_displays(display_frame)
    
    def create_input_fields(self, parent):
        """Create input form fields."""
        # Jahr
        tk.Label(parent, text="Jahr:*").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_year = tk.Entry(parent)
        self.entry_year.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        
        # Monat
        tk.Label(parent, text="Monat:*").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_month = tk.Entry(parent)
        self.entry_month.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        
        # Tag
        self.label_day = tk.Label(parent, text="Tag:*")
        self.label_day.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_day = tk.Entry(parent)
        self.entry_day.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        
        # Name
        tk.Label(parent, text="Name:*").grid(row=3, column=0, sticky="e", padx=5, pady=2)
        self.entry_name = tk.Entry(parent)
        self.entry_name.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        
        # Stunden
        tk.Label(parent, text="Stunden:*").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_hours = tk.Entry(parent)
        self.entry_hours.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        
        # CheckBox1
        self.check_unter_8h = tk.IntVar()
        check1 = tk.Checkbutton(parent, text="Unter 8h", variable=self.check_unter_8h)
        check1.grid(row=5, column=0, columnspan=2, pady=2)
        
        # CheckBox2
        self.check_skug = tk.IntVar()
        check2 = tk.Checkbutton(parent, text="SKUG", variable=self.check_skug, 
                                command=self.toggle_skug)
        check2.grid(row=6, column=0, columnspan=2, pady=2)
        
        # SKUG Feld (nur sichtbar wenn CheckBox2 aktiv)
        self.label_skug = tk.Label(parent, text="SKUG:")
        self.entry_skug = tk.Entry(parent)
        
        # Baustelle
        tk.Label(parent, text="Baustelle:").grid(row=8, column=0, sticky="e", padx=5, pady=2)
        self.entry_bst = tk.Entry(parent)
        self.entry_bst.grid(row=8, column=1, padx=5, pady=2, sticky="ew")
        
        # Required fields note
        tk.Label(parent, text="* Pflichtfelder", font=("Arial", 8), fg="gray").grid(
            row=9, column=0, columnspan=2, sticky="w", padx=5
        )
        
        # Buttons
        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=10, column=0, columnspan=2, pady=20)
        
        btn_submit = tk.Button(btn_frame, text="Speichern", command=self.submit)
        btn_submit.pack(side=tk.LEFT, padx=5)
        
        btn_export = tk.Button(btn_frame, text="Excel Export", command=self.export_excel)
        btn_export.pack(side=tk.LEFT, padx=5)
        
        # Configure column weight for resizing
        parent.grid_columnconfigure(1, weight=1)
        
        # Field list for navigation
        self.fields = [
            self.entry_year, self.entry_month, self.entry_day, 
            self.entry_name, self.entry_hours, self.entry_skug, self.entry_bst
        ]
    
    def create_data_displays(self, parent):
        """Create data display panels."""
        # --- MONTH VIEW (for person) ---
        month_frame = tk.LabelFrame(parent, text="Monat Übersicht (Jahr/Monat/Name)", padx=5, pady=5)
        month_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview for month data
        month_columns = ('Tag', 'Wochentag', 'Baustelle', 'Stunden', 'SKUG', 'Unter 8h')
        self.month_tree = ttk.Treeview(month_frame, columns=month_columns, show='headings', height=8)
        
        for col in month_columns:
            self.month_tree.heading(col, text=col)
            self.month_tree.column(col, width=80)
        
        # Scrollbar
        month_scrollbar = ttk.Scrollbar(month_frame, orient=tk.VERTICAL, command=self.month_tree.yview)
        self.month_tree.configure(yscrollcommand=month_scrollbar.set)
        
        self.month_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        month_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # --- DAY VIEW (for construction site) ---
        day_frame = tk.LabelFrame(parent, text="Tages Übersicht (Jahr/Monat/Tag/Baustelle)", padx=5, pady=5)
        day_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview for day data
        day_columns = ('Name', 'Stunden', 'SKUG', 'Unter 8h')
        self.day_tree = ttk.Treeview(day_frame, columns=day_columns, show='headings', height=8)
        
        for col in day_columns:
            self.day_tree.heading(col, text=col)
            self.day_tree.column(col, width=100)
        
        # Scrollbar
        day_scrollbar = ttk.Scrollbar(day_frame, orient=tk.VERTICAL, command=self.day_tree.yview)
        self.day_tree.configure(yscrollcommand=day_scrollbar.set)
        
        self.day_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        day_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_bindings(self):
        """Setup event bindings."""
        # Update weekday on date field changes
        self.entry_year.bind("<KeyRelease>", self.update_weekday)
        self.entry_month.bind("<KeyRelease>", self.update_weekday)
        self.entry_day.bind("<KeyRelease>", self.update_weekday)
        
        # Update month view when year, month, or name changes
        self.entry_year.bind("<KeyRelease>", self.update_month_view, add="+")
        self.entry_month.bind("<KeyRelease>", self.update_month_view, add="+")
        self.entry_name.bind("<KeyRelease>", self.update_month_view, add="+")
        
        # Update day view when year, month, day, or baustelle changes
        self.entry_year.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_month.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_day.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_bst.bind("<KeyRelease>", self.update_day_view, add="+")
        
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
    
    def update_month_view(self, *args):
        """Update the month overview display."""
        # Clear existing items
        for item in self.month_tree.get_children():
            self.month_tree.delete(item)
        
        year = self.entry_year.get().strip()
        month = self.entry_month.get().strip()
        name = self.entry_name.get().strip()
        
        # Only query if all required fields are filled
        if not (year and month and name):
            return
        
        try:
            year_int = int(year)
            month_int = int(month)
            
            # Get data from database
            entries = self.db.get_entries_by_month_and_name(year_int, month_int, name)
            
            # Populate treeview
            for entry in entries:
                self.month_tree.insert('', tk.END, values=(
                    entry['tag'],
                    entry['wochentag'] or '',
                    entry['baustelle'] or '',
                    entry['stunden'] or '',
                    entry['skug'] or 'Nein',
                    "Ja" if entry['unter_8h'] else "Nein"
                ))
        
        except (ValueError, TypeError):
            # Invalid year/month format
            pass
    
    def update_day_view(self, *args):
        """Update the day overview display."""
        # Clear existing items
        for item in self.day_tree.get_children():
            self.day_tree.delete(item)
        
        year = self.entry_year.get().strip()
        month = self.entry_month.get().strip()
        day = self.entry_day.get().strip()
        baustelle = self.entry_bst.get().strip()
        
        # Only query if all required fields are filled
        if not (year and month and day and baustelle):
            return
        
        try:
            year_int = int(year)
            month_int = int(month)
            day_int = int(day)
            
            # Get data from database
            entries = self.db.get_entries_by_date_and_baustelle(year_int, month_int, day_int, baustelle)
            
            # Populate treeview
            for entry in entries:
                self.day_tree.insert('', tk.END, values=(
                    entry['name'],
                    entry['stunden'] or '',
                    entry['skug'] or 'Nein',
                    "Ja" if entry['unter_8h'] else "Nein"
                ))
        
        except (ValueError, TypeError):
            # Invalid date format
            pass
    
    def toggle_skug(self):
        """Show/hide SKUG field based on CheckBox2."""
        if self.check_skug.get():
            self.label_skug.grid(row=7, column=0, sticky="e", padx=5, pady=2)
            self.entry_skug.grid(row=7, column=1, padx=5, pady=2, sticky="ew")
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
        stunden = self.entry_hours.get().strip()
        
        if not jahr:
            return (False, "Jahr ist erforderlich!")
        
        if not monat:
            return (False, "Monat ist erforderlich!")
        
        if not tag:
            return (False, "Tag ist erforderlich!")
        
        if not name:
            return (False, "Name ist erforderlich!")
        
        if not stunden:
            return (False, "Stunden sind erforderlich!")
        
        # Validate that they are valid numbers
        try:
            jahr_int = int(jahr)
            monat_int = int(monat)
            tag_int = int(tag)
            stunden_float = float(stunden)
            
            if not (1900 <= jahr_int <= 2100):
                return (False, "Jahr muss zwischen 1900 und 2100 liegen!")
            
            if not (1 <= monat_int <= 12):
                return (False, "Monat muss zwischen 1 und 12 liegen!")
            
            if not (1 <= tag_int <= 31):
                return (False, "Tag muss zwischen 1 und 31 liegen!")
            
            if not (0 <= stunden_float <= 24):
                return (False, "Stunden müssen zwischen 0 und 24 liegen!")
            
        except ValueError:
            return (False, "Jahr, Monat, Tag und Stunden müssen Zahlen sein!")
        
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
            "Stunden": float(self.entry_hours.get().strip()) or 0.0,
            "unter_8h": bool(self.check_unter_8h.get()),
            "check_skug": bool(self.check_skug.get()),
            "SKUG": self.entry_skug.get().strip() if self.check_skug.get() else "",
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
            
            # Refresh data displays
            self.update_month_view()
            self.update_day_view()
            
            # Clear fields for next entry (but keep the day)
            self.clear_fields()
            self.entry_day.focus()
        
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
        self.entry_hours.delete(0, tk.END)
        self.entry_skug.delete(0, tk.END)
    
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
                return "break"
            else:
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
                return "break"
        except ValueError:
            pass