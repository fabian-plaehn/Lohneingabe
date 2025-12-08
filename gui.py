import tkinter as tk
from tkinter import ttk, messagebox
from database import Database
from excel_export import export_to_excel
from utils import get_weekday_abbr, parse_date_range, parse_multiple_names, validate_days_in_month, calculate_skug, get_effective_fahrzeit
from datetime import datetime, timedelta
from master_data import MasterDataDatabase
from manager_dialogs import NameManagerDialog, BaustelleManagerDialog
from autocomplete import AutocompleteEntry, BaustelleAutocomplete
from settings_dialog import Settings, SettingsDialog
from datatypes import WorkerTypes, TravelStatus

class StundenEingabeGUI:
    def __init__(self, root):
        self.root = root
        self.db = Database()
        self.master_db = MasterDataDatabase()
        self.settings = Settings()
        self.setup_window()
        self.create_widgets()
        self.setup_bindings()
        self.apply_settings()

    def setup_window(self):
        """Configure main window."""
        self.root.title("Stunden-Eingabe")
        self.root.geometry("1000x600")

    def create_widgets(self):
        """Create all GUI widgets."""
        # Create a PanedWindow for resizable sections
        paned_window = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=5,
                                      sashrelief=tk.RAISED, bg='gray')
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left side - Input form
        input_frame = tk.Frame(paned_window)
        paned_window.add(input_frame, minsize=300)

        # Right side - Data displays
        display_frame = tk.Frame(paned_window)
        paned_window.add(display_frame, minsize=400)

        # --- INPUT FORM ---
        self.create_input_fields(input_frame)

        # --- DATA DISPLAYS ---
        self.create_data_displays(display_frame)

    def create_input_fields(self, parent):
        """Create input form fields."""
        # Jahr
        tk.Label(parent, text="Jahr:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_year = tk.Entry(parent)
        self.entry_year.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        # Monat
        tk.Label(parent, text="Monat:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_month = tk.Entry(parent)
        self.entry_month.grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        # Tag (with range support)
        self.label_day = tk.Label(parent, text="Tag(e):*")
        self.label_day.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_day = tk.Entry(parent)
        self.entry_day.grid(row=2, column=1, padx=5, pady=2, sticky="ew")

        # Add hint for Tag format
        tk.Label(parent, text="(z.B. 3-7,9,11-13)", font=("Arial", 7), fg="gray").grid(
            row=2, column=2, sticky="w", padx=2
        )

        # Name with manager button
        tk.Label(parent, text="Name(n):").grid(row=3, column=0, sticky="e", padx=5, pady=2)
        name_frame = tk.Frame(parent)
        name_frame.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.entry_name = tk.Entry(name_frame)
        self.entry_name.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_name_manager = tk.Button(name_frame, text="⚙", width=2, command=self.open_name_manager)
        btn_name_manager.pack(side=tk.LEFT, padx=(2, 0))

        # Add hint for Name format
        tk.Label(parent, text="(z.B. Max, Anna)", font=("Arial", 7), fg="gray").grid(
            row=3, column=2, sticky="w", padx=2
        )

        # Stunden
        tk.Label(parent, text="Stunden:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_hours = tk.Entry(parent)
        self.entry_hours.grid(row=4, column=1, padx=5, pady=2, sticky="ew")

        # Frühstück checkbox (+0.25 hours)
        self.check_fruehstueck = tk.IntVar()
        check_fruehstueck = tk.Checkbutton(parent, text="Frühstück (+0.25h)", variable=self.check_fruehstueck,
                                            command=self.toggle_fruehstueck)
        check_fruehstueck.grid(row=5, column=1, sticky="w", pady=2, padx=5)

        # Mittagspause checkbox (+0.5 hours)
        self.check_mittagspause = tk.IntVar()
        check_mittagspause = tk.Checkbutton(parent, text="Mittagspause (+0.5h)", variable=self.check_mittagspause,
                                            command=self.toggle_mittagspause)
        check_mittagspause.grid(row=6, column=1, sticky="w", pady=2, padx=5)

        # Urlaub checkbox
        self.check_urlaub = tk.IntVar()
        check_urlaub = tk.Checkbutton(parent, text="Urlaub", variable=self.check_urlaub,
                                      command=self.toggle_urlaub)
        check_urlaub.grid(row=7, column=1, sticky="w", pady=2, padx=5)

        # Krank checkbox
        self.check_krank = tk.IntVar()
        check_krank = tk.Checkbutton(parent, text="Krank", variable=self.check_krank,
                                     command=self.toggle_krank)
        check_krank.grid(row=8, column=1, sticky="w", pady=2, padx=5)

        # SKUG checkbox
        self.check_skug = tk.IntVar()
        check_skug = tk.Checkbutton(parent, text="SKUG", variable=self.check_skug,
                                    command=self.toggle_skug)
        check_skug.grid(row=9, column=1, sticky="w", pady=2, padx=5)

        # Travel Status
        travel_frame = tk.Frame(parent)
        travel_frame.grid(row=10, column=1, sticky="w", pady=2, padx=5)

        self.check_reise = tk.IntVar()
        check_reise = tk.Checkbutton(travel_frame, text="Reise", variable=self.check_reise, command=self.toggle_reise)
        check_reise.pack(side=tk.LEFT)

        self.combo_reise_type = ttk.Combobox(travel_frame, width=10, state="disabled")
        self.combo_reise_type['values'] = [ts.value for ts in TravelStatus]
        self.combo_reise_type.current(0)
        self.combo_reise_type.pack(side=tk.LEFT, padx=5)

        # Baustelle with manager button
        tk.Label(parent, text="Baustelle:").grid(row=11, column=0, sticky="e", padx=5, pady=2)
        bst_frame = tk.Frame(parent)
        bst_frame.grid(row=11, column=1, padx=5, pady=2, sticky="ew")
        self.entry_bst = tk.Entry(bst_frame)
        self.entry_bst.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_bst_manager = tk.Button(bst_frame, text="⚙", width=2, command=self.open_baustelle_manager)
        btn_bst_manager.pack(side=tk.LEFT, padx=(2, 0))

        # Delete Mode
        self.check_delete_mode = tk.IntVar()
        check_delete = tk.Checkbutton(parent, text="Löschen / Entfernen (Nur Checkbox)", variable=self.check_delete_mode, fg="red")
        check_delete.grid(row=12, column=1, sticky="w", pady=2, padx=5)

        # Buttons
        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=13, column=0, columnspan=2, pady=20)

        btn_submit = tk.Button(btn_frame, text="Speichern", command=self.submit)
        btn_submit.pack(side=tk.LEFT, padx=5)

        btn_export = tk.Button(btn_frame, text="Excel Export", command=self.export_excel)
        btn_export.pack(side=tk.LEFT, padx=5)

        btn_settings = tk.Button(btn_frame, text="⚙ Einstellungen", command=self.open_settings)
        btn_settings.pack(side=tk.LEFT, padx=5)

        # Configure column weight for resizing
        parent.grid_columnconfigure(1, weight=1)

        # Field list for navigation
        self.fields = [
            self.entry_year, self.entry_month, self.entry_day,
            self.entry_name, self.entry_hours, self.entry_bst
        ]

        # Setup autocomplete
        self.setup_autocomplete()

    def toggle_reise(self):
        """Enable/disable travel type combobox."""
        if self.check_reise.get():
            self.combo_reise_type.config(state="readonly")
            self.clear_krank()
            self.clear_urlaub()
            self.entry_hours.config(state="normal")
        else:
            self.combo_reise_type.config(state="disabled")

    def toggle_fruehstueck(self):
        if self.check_fruehstueck.get():
            self.clear_krank()
            self.clear_urlaub()
            self.entry_hours.config(state="normal")
     

    def toggle_mittagspause(self):
        if self.check_mittagspause.get():
            self.clear_krank()
            self.clear_urlaub()
            self.entry_hours.config(state="normal")

    def toggle_skug(self):
        if self.check_skug.get():
            self.clear_krank()
            self.clear_urlaub()
            self.entry_hours.config(state="normal")


    def create_data_displays(self, parent):
        """Create data display panels."""
        # Create a vertical PanedWindow for the display section
        display_paned = tk.PanedWindow(parent, orient=tk.VERTICAL, sashwidth=5,
                                      sashrelief=tk.RAISED, bg='gray')
        display_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- MONTH VIEW (for person) ---
        month_frame = tk.LabelFrame(display_paned, text="Monat Übersicht (Jahr/Monat/Name(n))", padx=5, pady=5)
        display_paned.add(month_frame, minsize=200, stretch="always") # Add to paned window instead of pack

        # Treeview for month data
        month_columns = ('Tag', 'Wochentag', 'Name', 'Baustelle', 'Stunden', 'F', 'M', 'Urlaub', 'Krank', 'SKUG', 'Reise', '≤ 8h', 'Löschen')
        self.month_tree = ttk.Treeview(month_frame, columns=month_columns, show='headings', height=8)

        # Configure columns with sorting
        self.month_sort_column = 'Tag'
        self.month_sort_reverse = False

        for col in month_columns:
            if col == 'Löschen':
                self.month_tree.heading(col, text=col)
                self.month_tree.column(col, width=60, anchor='center')
            else:
                self.month_tree.heading(col, text=col, command=lambda c=col: self.sort_month_tree(c))
                if col == 'Tag':
                    self.month_tree.column(col, width=50, anchor='center')
                elif col == 'Wochentag':
                    self.month_tree.column(col, width=80, anchor='center')
                elif col == 'Name':
                    self.month_tree.column(col, width=100, anchor='center')
                elif col == 'Reise':
                    self.month_tree.column(col, width=80, anchor='center')
                elif col in ('F', 'M'):
                    self.month_tree.column(col, width=30, anchor='center')
                else:
                    self.month_tree.column(col, width=80, anchor='center')

        # Configure tags for styling
        self.month_tree.tag_configure('row_red', background='#FF9999')
        self.month_tree.tag_configure('row_even', background='#E0E0E0')
        self.month_tree.tag_configure('row_odd', background='#FFFFFF')

        # Scrollbar
        month_scrollbar = ttk.Scrollbar(month_frame, orient=tk.VERTICAL, command=self.month_tree.yview)
        self.month_tree.configure(yscrollcommand=month_scrollbar.set)

        self.month_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        month_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind click event for delete
        self.month_tree.bind('<ButtonRelease-1>', self.on_month_tree_click)

        # --- DAY VIEW (for construction site) ---
        day_frame = tk.LabelFrame(display_paned, text="Tages Übersicht (Jahr/Monat/Tag(e)/Baustelle)", padx=5, pady=5)
        display_paned.add(day_frame, minsize=200, stretch="always") # Add to paned window instead of pack

        # Treeview for day data
        day_columns = ('Tag', 'Wochentag', 'Name', 'Stunden', 'Urlaub', 'Krank', 'SKUG', 'Reise', '≤ 8h')
        self.day_tree = ttk.Treeview(day_frame, columns=day_columns, show='headings', height=8)

        # Configure columns with sorting
        self.day_sort_column = 'Tag'
        self.day_sort_reverse = False

        for col in day_columns:
            self.day_tree.heading(col, text=col, command=lambda c=col: self.sort_day_tree(c))
            if col == 'Tag':
                self.day_tree.column(col, width=50, anchor='center')
            elif col == 'Wochentag':
                self.day_tree.column(col, width=80, anchor='center')
            elif col == 'Name':
                self.day_tree.column(col, width=100, anchor='center')
            elif col == 'Reise':
                self.day_tree.column(col, width=80, anchor='center')
            else:
                self.day_tree.column(col, width=90, anchor='center')
        
        # Configure tags for styling
        self.day_tree.tag_configure('row_even', background='#E0E0E0')
        self.day_tree.tag_configure('row_odd', background='#FFFFFF')

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
        autocomplete_fields = [self.entry_name, self.entry_bst]
        for field in self.fields:
            if field not in autocomplete_fields:
                field.bind("<Return>", self.focus_next)
                field.bind("<Down>", self.focus_next)
                field.bind("<Up>", self.focus_previous)
            else:
                # For autocomplete fields, add navigation with low priority
                # The autocomplete class handlers will run first and return "break" if dropdown is visible
                field.bind("<Return>", self.focus_next, add="+")
                field.bind("<Down>", self.focus_next, add="+")
                field.bind("<Up>", self.focus_previous, add="+")

    def update_weekday(self, *args):
        """Update day label with weekday abbreviation or range info."""
        tag_input = self.entry_day.get().strip()
        jahr = self.entry_year.get()
        monat = self.entry_month.get()

        # Get settings for filtering
        skip_weekends = self.settings.get("skip_weekends", True)
        skip_holidays = self.settings.get("skip_holidays", True)

        # Try to parse as date range
        try:
            jahr_int = int(jahr) if jahr else None
            monat_int = int(monat) if monat else None

            if jahr_int and monat_int:
                days = parse_date_range(tag_input, jahr_int, monat_int, skip_weekends, skip_holidays)
            else:
                days = parse_date_range(tag_input)
        except (ValueError, TypeError):
            days = parse_date_range(tag_input)

        if days:
            if len(days) == 1:
                # Single day - show weekday
                weekday = get_weekday_abbr(jahr, monat, str(days[0]))
                if weekday:
                    self.label_day.config(text=f"Tag(e) ({weekday}):*")
                else:
                    self.label_day.config(text="Tag(e):*")
            else:
                # Multiple days - show count
                self.label_day.config(text=f"Tag(e) ({len(days)} Tage):*")
        else:
            # Try single day
            try:
                single_day = int(tag_input)
                weekday = get_weekday_abbr(jahr, monat, str(single_day))
                if weekday:
                    self.label_day.config(text=f"Tag(e) ({weekday}):*")
                else:
                    self.label_day.config(text="Tag(e):*")
            except (ValueError, TypeError):
                self.label_day.config(text="Tag(e):*")

    def update_month_view(self, *args):
        """Update the month overview display."""
        # Clear existing items
        for item in self.month_tree.get_children():
            self.month_tree.delete(item)

        year = self.entry_year.get().strip()
        month = self.entry_month.get().strip()
        names_input = self.entry_name.get().strip()

        # Only query if year and month are filled
        if not (year and month):
            return

        try:
            year_int = int(year)
            month_int = int(month)

            # Parse multiple names
            names = parse_multiple_names(names_input)

            if not names:
                return

            # Get data from database for all names
            all_entries = []
            for name in names:
                entries = self.db.get_entries_by_month_and_name(year_int, month_int, name)
                all_entries.extend(entries)

            # Sort by day (default)
            all_entries.sort(key=lambda x: x['tag'])

            # Populate treeview
            for i, entry in enumerate(all_entries):
                # Determine tag based on kg_8h and row index (zebra striping)
                tags = []
                if entry['kg_8h']:
                    tags.append('row_red')
                else:
                    tags.append('row_even' if i % 2 == 0 else 'row_odd')
                
                # Add entry id tag
                tags.append(f"entry_{entry['id']}")

                # Store entry id as a tag for later retrieval
                self.month_tree.insert('', tk.END, values=(
                    entry['tag'],
                    entry['wochentag'] or '',
                    entry['name'],
                    entry['baustelle'] or '',
                    entry['stunden'] or '',
                    "X" if entry.get('fruehstueck') else "",
                    "X" if entry.get('mittag') else "",
                    entry.get('urlaub') or '',
                    entry.get('krank') or '',
                    entry['skug'] or '',
                    entry.get('travel_status') or '',
                    "Ja" if entry['kg_8h'] else ("" if entry['kg_8h'] is None else "Nein"),
                    '🗑'  # Delete icon
                ), tags=tuple(tags))

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
        day_input = self.entry_day.get().strip()
        baustelle = self.entry_bst.get().strip()

        # Only query if year, month, day and baustelle are filled
        if not (year and month and day_input and baustelle):
            return

        try:
            year_int = int(year)
            month_int = int(month)

            # Get settings for filtering
            skip_weekends = self.settings.get("skip_weekends", True)
            skip_holidays = self.settings.get("skip_holidays", True)

            # Parse date range with filtering
            days = parse_date_range(day_input, year_int, month_int, skip_weekends, skip_holidays)

            # If no range, treat as single day
            if days is None:
                try:
                    single_day = int(day_input)
                    if 1 <= single_day <= 31:
                        days = [single_day]
                    else:
                        return
                except ValueError:
                    return

            # Get data from database for all days
            all_entries = []
            for day in days:
                entries = self.db.get_entries_by_date_and_baustelle(year_int, month_int, day, baustelle)
                all_entries.extend(entries)

            # Sort by day (default)
            all_entries.sort(key=lambda x: x['tag'])

            # Populate treeview
            for i, entry in enumerate(all_entries):
                wochentag = get_weekday_abbr(str(year_int), str(month_int), str(entry['tag'])) or ''
                
                # Zebra striping
                row_tag = 'row_even' if i % 2 == 0 else 'row_odd'

                self.day_tree.insert('', tk.END, values=(
                    entry['tag'],
                    wochentag,
                    entry['name'],
                    entry['stunden'] or '',
                    entry.get('urlaub') or '',
                    entry.get('krank') or '',
                    entry['skug'] or '',
                    entry.get('travel_status') or '',
                    "Ja" if entry['kg_8h'] else ("" if entry['kg_8h'] is None else "Nein"),
                ), tags=(row_tag,))

        except (ValueError, TypeError):
            # Invalid date format
            pass

    def on_month_tree_click(self, event):
        """Handle click on month tree view to detect delete button clicks."""
        # Identify the region clicked
        region = self.month_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        # Get the column clicked
        column = self.month_tree.identify_column(event.x)

        # Column #13 is the delete column (0-indexed internally but #-indexed in identify)
        if column == '#13':  # Löschen column
            # Get the item clicked
            item = self.month_tree.identify_row(event.y)
            if item:
                # Get the entry ID from tags
                tags = self.month_tree.item(item, 'tags')
                if tags:
                    entry_id_str = tags[1]  # Format: "entry_123"
                    entry_id = int(entry_id_str.split('_')[1])

                    # Get entry details for confirmation
                    values = self.month_tree.item(item, 'values')
                    name = values[2]
                    tag = values[0]

                    # Confirm deletion
                    if messagebox.askyesno("Eintrag löschen",
                                          f"Möchten Sie den Eintrag für {name} am Tag {tag} wirklich löschen?"):
                        # Delete from database
                        if self.db.delete_entry(entry_id):
                            # Remove from treeview
                            self.month_tree.delete(item)
                            # Also refresh day view in case it's affected
                            self.update_day_view()
                        else:
                            messagebox.showerror("Fehler", "Eintrag konnte nicht gelöscht werden.")

    def toggle_krank(self):
        """Handle mutual exclusion of Krank and set hours to 0."""
        self.entry_hours.delete(0, tk.END)
        self.entry_hours.insert(0, "0")
        if self.check_krank.get():
            # If Urlaub is checked, uncheck Krank
            self.check_urlaub.set(0)
            self.check_fruehstueck.set(0)
            self.check_mittagspause.set(0)
            self.check_skug.set(0)
            self.check_reise.set(0)
            self.combo_reise_type.set(TravelStatus.Nicht)
            self.entry_hours.config(state="disabled")
            self.entry_bst.delete(0, tk.END)
            self.entry_bst.insert(0, "Krank")
            self.entry_bst.config(state="disabled")
        else:
            self.clear_krank()

    def clear_krank(self):
        self.check_krank.set(0)
        self.entry_hours.config(state="normal")
        self.entry_hours.delete(0, tk.END)
        self.entry_bst.config(state="normal")
        self.entry_bst.delete(0, tk.END)

    def toggle_urlaub(self):
        self.entry_hours.delete(0, tk.END)
        self.entry_hours.insert(0, "0")
        if self.check_urlaub.get():
            # If Urlaub is checked, uncheck Krank
            self.check_krank.set(0)
            self.check_fruehstueck.set(0)
            self.check_mittagspause.set(0)
            self.check_skug.set(0)
            self.check_reise.set(0)
            self.combo_reise_type.set(TravelStatus.Nicht)
            self.entry_hours.config(state="disabled")
            self.entry_bst.delete(0, tk.END)
            self.entry_bst.insert(0, "940")
            self.entry_bst.config(state="disabled")
        else:
            self.clear_urlaub()

    def clear_urlaub(self):
        self.check_urlaub.set(0)
        self.entry_hours.config(state="normal")
        self.entry_hours.delete(0, tk.END)
        self.entry_bst.config(state="normal")
        self.entry_bst.delete(0, tk.END)

    def validate_required_fields(self) -> tuple[bool, str]:
        """
        Validate that all required fields are filled.
        Returns (is_valid, error_message)
        """
        jahr = self.entry_year.get().strip()
        monat = self.entry_month.get().strip()
        tag = self.entry_day.get().strip()
        name = self.entry_name.get().strip()

        # edit hours in case of , as decimal separator
        stunden = self.entry_hours.get().strip()
        stunden = stunden.replace(',', '.')
        self.entry_hours.delete(0, tk.END)
        self.entry_hours.insert(0, stunden)
        baustelle = self.entry_bst.get().strip()

        worker_type = self.master_db.get_worker_type_by_name(name)

        if not jahr:
            return (False, "Jahr ist erforderlich!")

        if not monat:
            return (False, "Monat ist erforderlich!")

        if not name:
            return (False, "Name ist erforderlich!")

        # Hours are now optional if just updating Travel Status
        # if not stunden:
        #     return (False, "Stunden sind erforderlich!")

        #if not baustelle and worker_type == WorkerTypes.Gewerblich:
        #    return (False, "Baustelle ist erforderlich!")

        # Validate that they are valid numbers
        try:
            jahr_int = int(jahr)
            monat_int = int(monat)

            if stunden:
                stunden_float = float(stunden)
                if not (0 <= stunden_float <= 24):
                    return (False, "Stunden müssen zwischen 0 und 24 liegen!")

            if not (1900 <= jahr_int <= 2100):
                return (False, "Jahr muss zwischen 1900 und 2100 liegen!")

            if not (1 <= monat_int <= 12):
                return (False, "Monat muss zwischen 1 und 12 liegen!")

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

        # Parse multiple names
        names_input = self.entry_name.get().strip()
        names = parse_multiple_names(names_input)

        if not names:
            messagebox.showerror("Fehler", "Mindestens ein Name muss angegeben werden!")
            return

        # Get year and month for date parsing
        jahr_input = self.entry_year.get().strip()
        monat_input = self.entry_month.get().strip()
        
        # New Feature: Only clear Baustelle if Krank or 940 was applied
        should_clear_baustelle = False

        try:
            jahr_int = int(jahr_input)
            monat_int = int(monat_input)
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Jahr oder Monat!")
            return

        # Get settings for filtering
        skip_weekends = self.settings.get("skip_weekends", True)
        skip_holidays = self.settings.get("skip_holidays", True)

        # Parse date range with filtering
        tag_input = self.entry_day.get().strip()
        days = parse_date_range(tag_input, jahr_int, monat_int, skip_weekends, skip_holidays)

        # If no range, treat as single day
        if days is None:
            try:
                single_day = int(tag_input)
                if 1 <= single_day <= 31:
                    days = [single_day]
                else:
                    messagebox.showerror("Fehler", "Tag muss zwischen 1 und 31 liegen!")
                    return
            except ValueError:
                messagebox.showerror("Fehler", "Ungültiges Tag-Format! Verwenden Sie z.B. '3-7,9,11-13' oder '5'")
                return

        # Check if days list is empty after filtering
        if not days:
            messagebox.showwarning("Warnung",
                "Alle eingegebenen Tage wurden gefiltert (Wochenenden/Feiertage).\n"
                "Bitte passen Sie die Einstellungen an oder wählen Sie andere Tage.")
            return

        # Validate that all days are valid for the given month/year
        is_valid, invalid_days = validate_days_in_month(jahr_int, monat_int, days)
        if not is_valid:
            invalid_days_str = ', '.join(map(str, invalid_days))
            messagebox.showerror("Fehler",
                f"Die folgenden Tage existieren nicht im Monat {monat_input}/{jahr_input}:\n{invalid_days_str}")
            return

        # Get common data
        jahr = self.entry_year.get().strip()
        monat = self.entry_month.get().strip()
        baustelle_input = self.entry_bst.get().strip()
        
        # Determine Input Hours
        stunden_input = self.entry_hours.get().strip()
        new_stunden = float(stunden_input) if stunden_input else None

        # Checkbox Inputs
        input_fruehstueck = bool(self.check_fruehstueck.get())
        input_mittag = bool(self.check_mittagspause.get())
        input_urlaub = bool(self.check_urlaub.get())
        input_krank = bool(self.check_krank.get())
        input_skug = bool(self.check_skug.get())
        input_reise = bool(self.check_reise.get())
        
        if input_urlaub or input_krank:
            should_clear_baustelle = True
        
        # Delete Mode
        delete_mode = bool(self.check_delete_mode.get())
        
        # Check if we should skip (No hours and no checkboxes)
        # Note: Delete mode implies we want to do something, so don't skip
        if new_stunden is None and not any([input_fruehstueck, input_mittag, input_urlaub, input_krank, input_skug, input_reise, delete_mode]):
            messagebox.showinfo("Info", "Keine Stunden und keine Optionen gewählt - nichts zu tun.")
            return

        # Get SKUG settings for calculation
        skug_settings = self.master_db.get_skug_settings()

        # Prepare to save multiple entries
        total_entries = 0
        updated_entries = 0
        errors = []

        # Travel Status Logic (Global for the batch if hours/reise enabled)
        travel_type_input = self.combo_reise_type.get()
        if travel_type_input == TravelStatus.Nicht:
            travel_type_input = None
        
        # Sort days for Smart Range logic (Travel needs order)
        sorted_days = sorted(days)

        try:
            # Loop through all combinations of names and days
            for name in names:
                for i, day in enumerate(sorted_days):
                    # Get weekday for this specific day
                    wochentag = get_weekday_abbr(jahr, monat, str(day)) or ""

                    # --- RESOLVE EXISTING DATA ---
                    existing_entry = self.db.get_entry(jahr_int, monat_int, day, name) or {}
                    
                    # Resolve Hours
                    current_stunden = new_stunden
                    if current_stunden is None:
                        # Use existing hours if available
                        current_stunden = existing_entry.get('stunden')

                    # Resolve Baustelle
                    current_baustelle = baustelle_input
                    if not current_baustelle:
                        current_baustelle = existing_entry.get('baustelle', '')
                    
                    # Validation: Missing Baustelle with Hours
                    # If we have hours (and not clearing them via Urlaub/Krank), we need a Baustelle
                    # unless it's a specific worker type? User request: "Applied Stunden in a range without bst -> didnt tell me that it did not work"
                    # We check this later after resolving urlaub/krank, because urlaub/krank force hours to 0.

                    # --- RESOLVE CHECKBOXES (Additive / Subtractive) ---
                    
                    def resolve_flag(key, input_val):
                        existing_val = existing_entry.get(key)
                        existing_bool = bool(existing_val)
                        if delete_mode:
                            if input_val: return False
                            return existing_bool
                        else:
                            if input_val: return True
                            return existing_bool

                    final_fruehstueck = resolve_flag('fruehstueck', input_fruehstueck)
                    final_mittag = resolve_flag('mittag', input_mittag)
                    
                    # SKUG Flag Logic
                    existing_skug_bool = bool(existing_entry.get('skug'))
                    wants_skug = False
                    if delete_mode:
                        if input_skug: wants_skug = False
                        else: wants_skug = existing_skug_bool
                    else:
                        if input_skug: wants_skug = True
                        else: wants_skug = existing_skug_bool

                    # Urlaub/Krank Logic
                    final_urlaub_flag = False
                    final_krank_flag = False
                    
                    if delete_mode:
                        if input_urlaub: final_urlaub_flag = False
                        else: final_urlaub_flag = bool(existing_entry.get('urlaub'))
                        
                        if input_krank: final_krank_flag = False
                        else: final_krank_flag = bool(existing_entry.get('krank'))
                    else:
                        if input_urlaub: final_urlaub_flag = True
                        else: final_urlaub_flag = bool(existing_entry.get('urlaub'))
                        
                        if input_krank: final_krank_flag = True
                        else: final_krank_flag = bool(existing_entry.get('krank'))

                    # CONFLICT RESOLUTION: Hours vs Urlaub/Krank
                    # Bugfix Case 1: "Applied Stunden in a range with Bst -> Urlaub still stayed in Urlaub field"
                    # If New Hours are explicitly provided (new_stunden is not None), AND they are > 0, 
                    # we must clear Urlaub and Krank.
                    if new_stunden is not None and new_stunden > 0:
                        final_urlaub_flag = False
                        final_krank_flag = False

                    # Mutual exclusion for Urlaub/Krank (New input overrides old)
                    if input_urlaub and not delete_mode:
                        final_krank_flag = False
                        final_fruehstueck = False
                        final_mittag = False
                        wants_skug = False
                        final_travel_status = None
                        current_stunden = 0.0 # Force 0
                        current_kg_8h = None # Reset 8h check
                        
                    if input_krank and not delete_mode:
                        final_urlaub_flag = False
                        final_fruehstueck = False
                        final_mittag = False
                        wants_skug = False
                        final_travel_status = None
                        current_stunden = 0.0 # Force 0
                        current_kg_8h = None
                    
                    # Fix: If we set Urlaub/Krank, we should update the baustelle text?
                    # Actually, usually user logic sets Baustelle input to "940" or "Krank".
                    # Whatever is in current_baustelle is what we use.
                    
                    # Bugfix Case 2: "Removed Krank in a range -> Krank stayed in Baustelle"
                    # If we are removing Krank (delete_mode + input_krank), check if baustelle text is "Krank".
                    if delete_mode and input_krank:
                         if current_baustelle == "Krank":
                             current_baustelle = ""

                    # Bugfix Case 1 (Continued): "Urlaub stayed in Urlaub field"
                    # Handled by `if new_stunden > 0: clear flags` check above.
                    
                    # Travel Status
                    final_travel_status = existing_entry.get('travel_status')
                    if input_reise:
                        if delete_mode:
                             final_travel_status = None
                        else:
                            # Calculate smart travel status
                            if travel_type_input == TravelStatus.Auto:
                                if len(sorted_days) == 1:
                                    final_travel_status = "Anreise"
                                else:
                                    if i == 0:
                                        final_travel_status = "Anreise"
                                    elif i == len(sorted_days) - 1:
                                        final_travel_status = "Abreise"
                                    else:
                                        final_travel_status = "24h_away"
                            else:
                                final_travel_status = travel_type_input
                    
                    # --- DEPENDENCY CHECK ---
                    # Breakfast/Lunch require hours.
                    if (final_fruehstueck or final_mittag) and (current_stunden is None or current_stunden == 0):
                        final_fruehstueck = False
                        final_mittag = False
                    
                    # --- VALIDATION (Missing Baustelle) ---
                    # Bugfix Case 3: "Applied Stunden without bst -> didnt tell me that it did not work"
                    if current_stunden is not None and current_stunden > 0 and not current_baustelle:
                        errors.append(f"{name}, Tag {day}: Baustelle fehlt für {current_stunden}h")
                        continue # Skip saving this entry

                    # --- CALCULATIONS ---
                    
                    # Recalculate SKUG
                    final_skug_val = ""
                    if wants_skug and current_stunden is not None:
                        skug_value = calculate_skug(int(jahr), int(monat), day, current_stunden, skug_settings)
                        final_skug_val = str(skug_value) if skug_value != 0.0 else ""
                    
                    # Recalculate Urlaub (Value)
                    final_urlaub_val = ""
                    if final_urlaub_flag:
                        urlaub_value = calculate_skug(int(jahr), int(monat), day, 0, skug_settings)
                        final_urlaub_val = str(urlaub_value) if urlaub_value != 0.0 else ""
                        current_stunden = 0.0
                        # Also ensure Baustelle is set to 940 if empty? User usually inputs it via toggle.
                        # If user just checks box but forgets input field, we might want to auto-set it?
                        # Existing code toggles input field.
                    
                    # Recalculate Krank (Value)
                    final_krank_val = ""
                    if final_krank_flag:
                        krank_value = calculate_skug(int(jahr), int(monat), day, 0, skug_settings)
                        final_krank_val = str(krank_value) if krank_value != 0.0 else ""
                        current_stunden = 0.0

                    # Recalculate kg_8h
                    current_kg_8h = None
                    if current_stunden is not None and not final_urlaub_flag and not final_krank_flag:
                        verpflegungs_stunden = float(current_stunden)
                        if final_fruehstueck: verpflegungs_stunden += 0.25
                        if final_mittag: verpflegungs_stunden += 0.5
                        
                        # Add travel time if applicable
                        if current_baustelle:
                            bst_nummer = current_baustelle.split('-')[0].strip() if '-' in current_baustelle else current_baustelle
                            bst_data = self.master_db.get_baustelle_by_nummer(bst_nummer)
                            if bst_data:
                                 worker_id = self.master_db.get_worker_id_by_name(name)
                                 fahrzeit = get_effective_fahrzeit(self.master_db, worker_id, bst_data['id'], bst_data.get('fahrzeit', 0.0))
                                 verpflegungs_stunden += float(fahrzeit)
                        
                        if verpflegungs_stunden <= 8.0:
                            current_kg_8h = True
                        else:
                            current_kg_8h = False 
                    
                    # --- CLEANUP (Empty Entries) ---
                    # Bugfix Case 2 hint: "maybe remove empty entries"
                    is_empty = (
                        (current_stunden is None or current_stunden == 0) and
                        not final_urlaub_flag and
                        not final_krank_flag and
                        not wants_skug and # Should check final_skug_val really, but wants_skug is intent
                        not final_travel_status and
                        not final_fruehstueck and 
                        not final_mittag
                    )
                    
                    if is_empty:
                        # If existing entry exists, delete it
                        if existing_entry and existing_entry.get('id'):
                            self.db.delete_entry(existing_entry['id'])
                            updated_entries += 1 # Count as update?
                        continue # Don't save

                    # --- DATA PREPARATION ---
                    data = {
                        "Jahr": jahr,
                        "Monat": monat,
                        "Tag": str(day),
                        "Name": name,
                        "Wochentag": wochentag,
                        "Urlaub": final_urlaub_val,
                        "Krank": final_krank_val,
                        "kg_8h": current_kg_8h,
                        "SKUG": final_skug_val,
                        "Baustelle": current_baustelle,
                        "fruehstueck": final_fruehstueck,
                        "mittag": final_mittag,
                        "travel_status": final_travel_status
                    }
                    
                    if current_stunden is not None:
                         data["Stunden"] = current_stunden

                    # --- SAVE ---
                    entry_id, was_updated = self.db.add_or_update_entry(data)
                    total_entries += 1
                    if was_updated:
                        updated_entries += 1

            # Show summary message including errors
            if errors:
                error_msg = f"{total_entries} Einträge verarbeitet.\n\nFehler:\n" + "\n".join(errors[:10])
                if len(errors) > 10:
                    error_msg += f"\n... und {len(errors) - 10} weitere Fehler"
                messagebox.showwarning("Hinweis", error_msg)

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern:\n{str(e)}")

        # Refresh data displays
        self.update_month_view()
        self.update_day_view()

        # Auto-increment day if enabled
        if self.settings.get("auto_increment_day", False) and not delete_mode:
            # Get the last day from the range
            last_day = max(days)

            # Check if we should skip weekends
            if self.settings.get("skip_weekends", True):
                next_year, next_month, next_day = self.get_next_day_skip_weekend(jahr, monat, last_day)
            else:
                next_year, next_month, next_day = self.get_next_day(jahr, monat, last_day)

            # Update month/year if they changed
            if next_year != int(jahr):
                self.entry_year.delete(0, tk.END)
                self.entry_year.insert(0, str(next_year))
            if next_month != int(monat):
                self.entry_month.delete(0, tk.END)
                self.entry_month.insert(0, str(next_month))

            # Update the day field
            self.entry_day.delete(0, tk.END)
            self.entry_day.insert(0, str(next_day))

            # Update the weekday label
            self.update_weekday()

        # Clear fields for next entry
        self.clear_fields(clear_baustelle=should_clear_baustelle)

        # Jump to configured field
        cursor_target = self.settings.get("cursor_jump_target", "Tag")
        if cursor_target == "Tag":
            self.entry_day.focus()
        elif cursor_target == "Name":
            self.entry_name.focus()
        elif cursor_target == "Stunden":
            self.entry_hours.focus()
        elif cursor_target == "Baustelle":
            self.entry_bst.focus()
        else:
            self.entry_day.focus()  # Default fallback

    def export_excel(self):
        """Export database to Excel for the currently selected year and month."""
        try:
            # Get year and month from entry fields
            jahr_str = self.entry_year.get().strip()
            monat_str = self.entry_month.get().strip()

            if not jahr_str or not monat_str:
                messagebox.showwarning("Warnung", "Bitte Jahr und Monat eingeben.")
                return

            try:
                jahr = int(jahr_str)
                monat = int(monat_str)

                if monat < 1 or monat > 12:
                    messagebox.showwarning("Warnung", "Monat muss zwischen 1 und 12 liegen.")
                    return

            except ValueError:
                messagebox.showwarning("Warnung", "Jahr und Monat müssen gültige Zahlen sein.")
                return

            # Use the new month-specific export
            if export_to_excel(jahr, monat, self.db, self.master_db):
                messagebox.showinfo("Erfolg", f"Daten für {monat:02d}/{jahr} nach Excel exportiert!")
            else:
                messagebox.showwarning("Warnung", "Keine Daten zum Exportieren vorhanden.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen:\n{str(e)}")

    def get_next_day(self, year, month, day):
        """
        Get the next day without skipping weekends.
        Returns (year, month, day) tuple.
        """
        try:
            current_date = datetime(int(year), int(month), int(day))
            next_date = current_date + timedelta(days=1)
            return (next_date.year, next_date.month, next_date.day)
        except (ValueError, TypeError):
            # If invalid date, just increment day by 1
            return (year, month, day + 1)

    def get_next_day_skip_weekend(self, year, month, day):
        """
        Get the next day, skipping weekends.
        If day is Friday, return Monday.
        Returns (year, month, day) tuple.
        """
        try:
            current_date = datetime(int(year), int(month), int(day))
            next_date = current_date + timedelta(days=1)

            # Check if next day is Saturday (5) or Sunday (6)
            while next_date.weekday() >= 5:  # 5 = Saturday, 6 = Sunday
                next_date += timedelta(days=1)

            return (next_date.year, next_date.month, next_date.day)
        except (ValueError, TypeError):
            # If invalid date, just increment day by 1
            return (year, month, day + 1)

    def clear_fields(self, clear_baustelle=True):
        """Clear input fields after submission (except day fields)."""
        self.entry_hours.delete(0, tk.END)
        self.check_urlaub.set(False)
        self.check_krank.set(False)
        self.check_reise.set(False)
        self.check_fruehstueck.set(False)
        self.check_mittagspause.set(False)
        self.entry_hours.config(state="normal")
        self.entry_hours.delete(0, tk.END)
        self.entry_bst.config(state="normal")
        
        if clear_baustelle:
            self.entry_bst.delete(0, tk.END)
        
        self.check_delete_mode.set(False)



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

    def setup_autocomplete(self):
        """Setup autocomplete for Name and Baustelle fields."""
        # Name autocomplete
        self.name_autocomplete = AutocompleteEntry(
            self.entry_name,
            self.get_name_suggestions
        )

        # Baustelle autocomplete
        self.baustelle_autocomplete = BaustelleAutocomplete(
            self.entry_bst,
            self.get_baustelle_suggestions
        )

    def get_name_suggestions(self):
        """Get list of name suggestions from database."""
        names_data = self.master_db.get_all_names()
        return [n['name'] for n in names_data]

    def get_baustelle_suggestions(self):
        """Get list of baustelle suggestions from database."""
        return self.master_db.get_all_baustellen()

    def sort_month_tree(self, col):
        """Sort month treeview by column."""
        # Toggle sort direction if clicking same column
        if col == self.month_sort_column:
            self.month_sort_reverse = not self.month_sort_reverse
        else:
            self.month_sort_column = col
            self.month_sort_reverse = False

        # Get all items
        items = [(self.month_tree.set(item, col), item) for item in self.month_tree.get_children('')]

        # Sort items
        if col in ('Tag', 'Stunden'):
            # Numeric sort
            try:
                items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=self.month_sort_reverse)
            except ValueError:
                items.sort(reverse=self.month_sort_reverse)
        else:
            # String sort
            items.sort(reverse=self.month_sort_reverse)

        # Rearrange items
        for index, (val, item) in enumerate(items):
            self.month_tree.move(item, '', index)

        # Update heading to show sort indicator
        for column in self.month_tree['columns']:
            heading_text = column
            if column == col:
                heading_text = f"{column} {'▼' if self.month_sort_reverse else '▲'}"
            self.month_tree.heading(column, text=heading_text)

    def sort_day_tree(self, col):
        """Sort day treeview by column."""
        # Toggle sort direction if clicking same column
        if col == self.day_sort_column:
            self.day_sort_reverse = not self.day_sort_reverse
        else:
            self.day_sort_column = col
            self.day_sort_reverse = False

        # Get all items
        items = [(self.day_tree.set(item, col), item) for item in self.day_tree.get_children('')]

        # Sort items
        if col in ('Tag', 'Stunden'):
            # Numeric sort
            try:
                items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=self.day_sort_reverse)
            except ValueError:
                items.sort(reverse=self.day_sort_reverse)
        else:
            # String sort
            items.sort(reverse=self.day_sort_reverse)

        # Rearrange items
        for index, (val, item) in enumerate(items):
            self.day_tree.move(item, '', index)

        # Update heading to show sort indicator
        for column in self.day_tree['columns']:
            heading_text = column
            if column == col:
                heading_text = f"{column} {'▼' if self.day_sort_reverse else '▲'}"
            self.day_tree.heading(column, text=heading_text)

    def open_name_manager(self):
        """Open the name manager dialog."""
        NameManagerDialog(self.root)
        # Refresh autocomplete will happen automatically on next keystroke

    def open_baustelle_manager(self):
        """Open the baustelle manager dialog."""
        BaustelleManagerDialog(self.root)
        # Refresh autocomplete will happen automatically on next keystroke

    def open_settings(self):
        """Open the settings dialog."""
        dialog = SettingsDialog(self.root, self.settings, self.master_db)
        # Wait for dialog to close, then reload settings
        self.root.wait_window(dialog.dialog)
        # Reload settings from file (they may have been changed)
        self.settings.current_settings = self.settings.load()
        self.apply_settings()

    def apply_settings(self):
        """Apply loaded settings to the GUI."""
        # Settings are now applied dynamically from self.settings.get() calls
        # This method is here for future use if needed
        pass