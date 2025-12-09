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
from datatypes import TravelStatus

class StundenEingabeGUI:
    def __init__(self, root):
        self.root = root
        self.db = Database()
        self.master_db = MasterDataDatabase()
        self.settings = Settings()
        self.setup_window()
        self.create_widgets()
        self.setup_bindings()

    def setup_window(self):
        self.root.title("Stunden-Eingabe")
        self.root.geometry("1000x600")

    def create_widgets(self):
        paned_window = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=5,
                                      sashrelief=tk.RAISED, bg='gray')
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        input_frame = tk.Frame(paned_window)
        paned_window.add(input_frame, minsize=300)
        display_frame = tk.Frame(paned_window)
        paned_window.add(display_frame, minsize=400)
        self.create_input_fields(input_frame)
        self.create_data_displays(display_frame)

    def create_input_fields(self, parent):
        tk.Label(parent, text="Jahr:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_year = tk.Entry(parent)
        self.entry_year.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        tk.Label(parent, text="Monat:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_month = tk.Entry(parent)
        self.entry_month.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.label_day = tk.Label(parent, text="Tag(e):*")
        self.label_day.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_day = tk.Entry(parent)
        self.entry_day.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        tk.Label(parent, text="(z.B. 3-7,9,11-13)", font=("Arial", 7), fg="gray").grid(
            row=2, column=2, sticky="w", padx=2
        )
        tk.Label(parent, text="Name(n):").grid(row=3, column=0, sticky="e", padx=5, pady=2)
        name_frame = tk.Frame(parent)
        name_frame.grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        self.entry_name = tk.Entry(name_frame)
        self.entry_name.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_name_manager = tk.Button(name_frame, text="⚙", width=2, command=self.open_name_manager)
        btn_name_manager.pack(side=tk.LEFT, padx=(2, 0))
        tk.Label(parent, text="(z.B. Max, Anna)", font=("Arial", 7), fg="gray").grid(
            row=3, column=2, sticky="w", padx=2
        )
        tk.Label(parent, text="Stunden:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_hours = tk.Entry(parent)
        self.entry_hours.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        self.check_fruehstueck = tk.IntVar()
        check_fruehstueck = tk.Checkbutton(parent, text="Frühstück (+0.25h)", variable=self.check_fruehstueck,
                                            command=self.toggle_fruehstueck)
        check_fruehstueck.grid(row=5, column=1, sticky="w", pady=2, padx=5)
        self.check_mittagspause = tk.IntVar()
        check_mittagspause = tk.Checkbutton(parent, text="Mittagspause (+0.5h)", variable=self.check_mittagspause,
                                            command=self.toggle_mittagspause)
        check_mittagspause.grid(row=6, column=1, sticky="w", pady=2, padx=5)
        self.check_urlaub = tk.IntVar()
        check_urlaub = tk.Checkbutton(parent, text="Urlaub", variable=self.check_urlaub,
                                      command=self.toggle_urlaub)
        check_urlaub.grid(row=7, column=1, sticky="w", pady=2, padx=5)
        self.check_krank = tk.IntVar()
        check_krank = tk.Checkbutton(parent, text="Krank", variable=self.check_krank,
                                     command=self.toggle_krank)
        check_krank.grid(row=8, column=1, sticky="w", pady=2, padx=5)
        
        self.check_skug = tk.IntVar()
        check_skug = tk.Checkbutton(parent, text="SKUG", variable=self.check_skug,
                                    command=self.toggle_skug)
        check_skug.grid(row=9, column=1, sticky="w", pady=2, padx=5)
        
        travel_frame = tk.Frame(parent)
        travel_frame.grid(row=10, column=1, sticky="w", pady=2, padx=5)

        self.check_reise = tk.IntVar()
        check_reise = tk.Checkbutton(travel_frame, text="Reise", variable=self.check_reise, command=self.toggle_reise)
        check_reise.pack(side=tk.LEFT)

        self.combo_reise_type = ttk.Combobox(travel_frame, width=10, state="disabled")
        self.combo_reise_type['values'] = [ts.value for ts in TravelStatus]
        self.combo_reise_type.current(0)
        self.combo_reise_type.pack(side=tk.LEFT, padx=5)

        tk.Label(parent, text="Baustelle:").grid(row=11, column=0, sticky="e", padx=5, pady=2)
        bst_frame = tk.Frame(parent)
        bst_frame.grid(row=11, column=1, padx=5, pady=2, sticky="ew")
        self.entry_bst = tk.Entry(bst_frame)
        self.entry_bst.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_bst_manager = tk.Button(bst_frame, text="⚙", width=2, command=self.open_baustelle_manager)
        btn_bst_manager.pack(side=tk.LEFT, padx=(2, 0))

        self.check_delete_mode = tk.IntVar()
        check_delete = tk.Checkbutton(parent, text="Entfernt Checkbox", variable=self.check_delete_mode, fg="red")
        check_delete.grid(row=12, column=1, sticky="w", pady=2, padx=5)

        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=13, column=0, columnspan=2, pady=20)

        btn_submit = tk.Button(btn_frame, text="Speichern", command=self.submit)
        btn_submit.pack(side=tk.LEFT, padx=5)

        btn_export = tk.Button(btn_frame, text="Excel Export", command=self.export_excel)
        btn_export.pack(side=tk.LEFT, padx=5)

        btn_settings = tk.Button(btn_frame, text="⚙ Einstellungen", command=self.open_settings)
        btn_settings.pack(side=tk.LEFT, padx=5)

        parent.grid_columnconfigure(1, weight=1)

        self.fields = [
            self.entry_year, self.entry_month, self.entry_day,
            self.entry_name, self.entry_hours, self.entry_bst
        ]

        self.setup_autocomplete()

    def toggle_reise(self):
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
        display_paned = tk.PanedWindow(parent, orient=tk.VERTICAL, sashwidth=5,
                                      sashrelief=tk.RAISED, bg='gray')
        display_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        month_frame = tk.LabelFrame(display_paned, text="Monat Übersicht (Jahr/Monat/Name(n))", padx=5, pady=5)
        display_paned.add(month_frame, minsize=200, stretch="always")

        month_columns = ('Tag', 'Wochentag', 'Name', 'Baustelle', 'Stunden', 'F', 'M', 'Urlaub', 'Krank', 'SKUG', 'Reise', '≤ 8h', 'Löschen')
        self.month_tree = ttk.Treeview(month_frame, columns=month_columns, show='headings', height=8)

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

        self.month_tree.tag_configure('row_red', background='#FF9999')
        self.month_tree.tag_configure('row_even', background='#E0E0E0')
        self.month_tree.tag_configure('row_odd', background='#FFFFFF')
        month_scrollbar = ttk.Scrollbar(month_frame, orient=tk.VERTICAL, command=self.month_tree.yview)
        self.month_tree.configure(yscrollcommand=month_scrollbar.set)

        self.month_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        month_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.month_tree.bind('<ButtonRelease-1>', self.on_month_tree_click)

        day_frame = tk.LabelFrame(display_paned, text="Tages Übersicht (Jahr/Monat/Tag(e)/Baustelle)", padx=5, pady=5)
        display_paned.add(day_frame, minsize=200, stretch="always")

        day_columns = ('Tag', 'Wochentag', 'Name', 'Stunden', 'Urlaub', 'Krank', 'SKUG', 'Reise', '≤ 8h')
        self.day_tree = ttk.Treeview(day_frame, columns=day_columns, show='headings', height=8)

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
        
        self.day_tree.tag_configure('row_even', background='#E0E0E0')
        self.day_tree.tag_configure('row_odd', background='#FFFFFF')

        day_scrollbar = ttk.Scrollbar(day_frame, orient=tk.VERTICAL, command=self.day_tree.yview)
        self.day_tree.configure(yscrollcommand=day_scrollbar.set)

        self.day_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        day_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_bindings(self):
        self.entry_year.bind("<KeyRelease>", self.update_weekday)
        self.entry_month.bind("<KeyRelease>", self.update_weekday)
        self.entry_day.bind("<KeyRelease>", self.update_weekday)

        self.entry_year.bind("<KeyRelease>", self.update_month_view, add="+")
        self.entry_month.bind("<KeyRelease>", self.update_month_view, add="+")
        self.entry_name.bind("<KeyRelease>", self.update_month_view, add="+")

        self.entry_year.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_month.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_day.bind("<KeyRelease>", self.update_day_view, add="+")
        self.entry_bst.bind("<KeyRelease>", self.update_day_view, add="+")

        autocomplete_fields = [self.entry_name, self.entry_bst]
        for field in self.fields:
            if field not in autocomplete_fields:
                field.bind("<Return>", self.focus_next)
                field.bind("<Down>", self.focus_next)
                field.bind("<Up>", self.focus_previous)
            else:
                field.bind("<Return>", self.focus_next, add="+")
                field.bind("<Down>", self.focus_next, add="+")
                field.bind("<Up>", self.focus_previous, add="+")

    def update_weekday(self, *args):
        tag_input = self.entry_day.get().strip()
        jahr = self.entry_year.get()
        monat = self.entry_month.get()

        skip_weekends = self.settings.get("skip_weekends", True)
        skip_holidays = self.settings.get("skip_holidays", True)

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
                weekday = get_weekday_abbr(jahr, monat, str(days[0]))
                if weekday:
                    self.label_day.config(text=f"Tag(e) ({weekday}):*")
                else:
                    self.label_day.config(text="Tag(e):*")
            else:
                self.label_day.config(text=f"Tag(e) ({len(days)} Tage):*")
        else:
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
        for item in self.month_tree.get_children():
            self.month_tree.delete(item)

        year = self.entry_year.get().strip()
        month = self.entry_month.get().strip()
        names_input = self.entry_name.get().strip()

        if not (year and month):
            return

        try:
            year_int = int(year)
            month_int = int(month)

            names = parse_multiple_names(names_input)

            if not names:
                return

            all_entries = []
            for name in names:
                entries = self.db.get_entries_by_month_and_name(year_int, month_int, name)
                all_entries.extend(entries)

            all_entries.sort(key=lambda x: x['tag'])

            for i, entry in enumerate(all_entries):
                tags = []
                if entry['kg_8h']:
                    tags.append('row_red')
                else:
                    tags.append('row_even' if i % 2 == 0 else 'row_odd')

                tags.append(f"entry_{entry['id']}")

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
                    '🗑'
                ), tags=tuple(tags))

        except (ValueError, TypeError):
            pass
        

    def update_day_view(self, *args):
        for item in self.day_tree.get_children():
            self.day_tree.delete(item)

        year = self.entry_year.get().strip()
        month = self.entry_month.get().strip()
        day_input = self.entry_day.get().strip()
        baustelle = self.entry_bst.get().strip()

        if not (year and month and day_input and baustelle):
            return

        try:
            year_int = int(year)
            month_int = int(month)

            skip_weekends = self.settings.get("skip_weekends", True)
            skip_holidays = self.settings.get("skip_holidays", True)

            days = parse_date_range(day_input, year_int, month_int, skip_weekends, skip_holidays)

            if days is None:
                try:
                    single_day = int(day_input)
                    if 1 <= single_day <= 31:
                        days = [single_day]
                    else:
                        return
                except ValueError:
                    return

            all_entries = []
            for day in days:
                entries = self.db.get_entries_by_date_and_baustelle(year_int, month_int, day, baustelle)
                all_entries.extend(entries)

            all_entries.sort(key=lambda x: x['tag'])

            for i, entry in enumerate(all_entries):
                wochentag = get_weekday_abbr(str(year_int), str(month_int), str(entry['tag'])) or ''

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
            pass

    def on_month_tree_click(self, event):
        region = self.month_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.month_tree.identify_column(event.x)

        if column == '#13':
            item = self.month_tree.identify_row(event.y)
            if item:
                tags = self.month_tree.item(item, 'tags')
                if tags:
                    entry_id_str = tags[1]  # Format: "entry_123"
                    entry_id = int(entry_id_str.split('_')[1])

                    values = self.month_tree.item(item, 'values')
                    name = values[2]
                    tag = values[0]

                    if messagebox.askyesno("Eintrag löschen",
                                          f"Möchten Sie den Eintrag für {name} am Tag {tag} wirklich löschen?"):
                        if self.db.delete_entry(entry_id):
                            self.month_tree.delete(item)
                            self.update_day_view()
                        else:
                            messagebox.showerror("Fehler", "Eintrag konnte nicht gelöscht werden.")

    def toggle_krank(self):
        self.entry_hours.delete(0, tk.END)
        self.entry_hours.insert(0, "0")
        if self.check_krank.get():
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
        jahr = self.entry_year.get().strip()
        monat = self.entry_month.get().strip()
        tag = self.entry_day.get().strip()
        name = self.entry_name.get().strip()

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

        jahr_input = self.entry_year.get().strip()
        monat_input = self.entry_month.get().strip()
        
        try:
            jahr_int = int(jahr_input)
            monat_int = int(monat_input)
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Jahr oder Monat!")
            return

        skip_weekends = self.settings.get("skip_weekends", True)
        skip_holidays = self.settings.get("skip_holidays", True)

        tag_input = self.entry_day.get().strip()
        days = parse_date_range(tag_input, jahr_int, monat_int, skip_weekends, skip_holidays)

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

        if not days:
            messagebox.showwarning("Warnung",
                "Alle eingegebenen Tage wurden gefiltert (Wochenenden/Feiertage).\n"
                "Bitte passen Sie die Einstellungen an oder wählen Sie andere Tage.")
            return

        is_valid, invalid_days = validate_days_in_month(jahr_int, monat_int, days)
        if not is_valid:
            invalid_days_str = ', '.join(map(str, invalid_days))
            messagebox.showerror("Fehler",
                f"Die folgenden Tage existieren nicht im Monat {monat_input}/{jahr_input}:\n{invalid_days_str}")
            return

        jahr = self.entry_year.get().strip()
        monat = self.entry_month.get().strip()
        baustelle_input = self.entry_bst.get().strip()

        stunden_input = self.entry_hours.get().strip()
        new_stunden = float(stunden_input) if stunden_input else None

        input_fruehstueck = bool(self.check_fruehstueck.get())
        input_mittag = bool(self.check_mittagspause.get())
        input_urlaub = bool(self.check_urlaub.get())
        input_krank = bool(self.check_krank.get())
        input_skug = bool(self.check_skug.get())
        input_reise = bool(self.check_reise.get())

        delete_mode = bool(self.check_delete_mode.get())

        if new_stunden is None and not any([input_fruehstueck, input_mittag, input_urlaub, input_krank, input_skug, input_reise, delete_mode]):
            messagebox.showinfo("Info", "Keine Stunden und keine Optionen gewählt - nichts zu tun.")
            return

        skug_settings = self.master_db.get_skug_settings()

        total_entries = 0
        updated_entries = 0
        errors = []

        travel_type_input = self.combo_reise_type.get()
        if travel_type_input == TravelStatus.Nicht:
            travel_type_input = None
 
        sorted_days = sorted(days)

        try:
            for name in names:
                for i, day in enumerate(sorted_days):
                    wochentag = get_weekday_abbr(jahr, monat, str(day)) or ""

                    existing_entries = self.db.get_entries_for_day(jahr_int, monat_int, day, name)
                    if input_krank or input_urlaub:
                        for entry in existing_entries:
                            self.db.delete_entry(entry['id'])
                        
                        final_urlaub_val = ""
                        final_krank_val = ""
                        bst_val = ""

                        if input_krank:
                             krank_value = calculate_skug(int(jahr), int(monat), day, 0, skug_settings)
                             final_krank_val = str(krank_value) if krank_value != 0.0 else ""
                             bst_val = "Krank"
                        elif input_urlaub:
                             urlaub_value = calculate_skug(int(jahr), int(monat), day, 0, skug_settings)
                             final_urlaub_val = str(urlaub_value) if urlaub_value != 0.0 else ""
                             bst_val = "940" # Standard for Urlaub

                        data = {
                            "Jahr": jahr, "Monat": monat, "Tag": str(day), "Name": name, "Wochentag": wochentag,
                            "Stunden": 0.0, "Urlaub": final_urlaub_val, "Krank": final_krank_val,
                            "kg_8h": None, "SKUG": "", "Baustelle": bst_val,
                            "fruehstueck": False, "mittag": False, "travel_status": None
                        }
                        self.db.insert_entry(data)
                        total_entries += 1
                        continue

                    target_entry_id = None
                    entry_data = {}
                    
                    if baustelle_input:
                        match = next((e for e in existing_entries if e['baustelle'] == baustelle_input), None)
                        if match:
                            target_entry_id = match['id']
                            entry_data = dict(match)
                            updated_entries += 1
                        else:
                            for e in existing_entries:
                                if e.get('krank') or e.get('urlaub'):
                                    self.db.delete_entry(e['id'])
                            target_entry_id = None
                            entry_data = {} 
                    else:
                        if len(existing_entries) == 0:
                            errors.append(f"{name}, Tag {day}: Keine Baustelle angegeben und kein Eintrag vorhanden.")
                            continue
                        elif len(existing_entries) == 1:
                            target_entry_id = existing_entries[0]['id']
                            entry_data = dict(existing_entries[0])
                            updated_entries += 1
                        else:
                            errors.append(f"{name}, Tag {day}: Keine Baustelle angegeben, aber mehrere Einträge vorhanden. Bitte Baustelle spezifizieren.")
                            continue
                    
                    if new_stunden is not None:
                         entry_data['Stunden'] = new_stunden
                    
                    if baustelle_input:
                        entry_data['Baustelle'] = baustelle_input
                    
                    if not target_entry_id:
                        entry_data.update({
                            "Jahr": jahr, "Monat": monat, "Tag": str(day), "Name": name, "Wochentag": wochentag,
                            "Stunden": new_stunden if new_stunden is not None else 0.0,
                            "Baustelle": baustelle_input
                        })

                    if input_fruehstueck: entry_data['fruehstueck'] = True
                    if input_mittag: entry_data['mittag'] = True
                    if input_reise: entry_data['travel_status'] = travel_type_input
                    if input_skug: entry_data['SKUG'] = ""

                    if input_fruehstueck: entry_data['fruehstueck'] = True
                    elif target_entry_id:
                         pass

                    if input_fruehstueck: entry_data['fruehstueck'] = True
                    if input_mittag: entry_data['mittag'] = True
                    if input_skug: 
                        entry_data['SKUG'] = str(calculate_skug(int(jahr), int(monat), day, entry_data.get('Stunden', 0), skug_settings))
                    
                    if input_reise:
                        final_travel_status = None
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
                        entry_data['travel_status'] = final_travel_status

                    if target_entry_id:
                        self.db.update_entry(target_entry_id, entry_data)
                    else:
                        target_entry_id = self.db.insert_entry(entry_data)
                    
                    total_entries += 1

                    day_entries = self.db.get_entries_for_day(jahr_int, monat_int, day, name)
                    
                    total_hours = 0.0
                    for e in day_entries:
                        h = float(e.get('stunden') or 0.0)
                        if e.get('fruehstueck'): h += 0.25
                        if e.get('mittag'): h += 0.5
                   
                        bst_name = e.get('baustelle')
                        if bst_name:
                             bst_nummer = bst_name.split('-')[0].strip() if '-' in bst_name else bst_name
                             bst_data = self.master_db.get_baustelle_by_nummer(bst_nummer)
                             if bst_data:
                                  worker_id = self.master_db.get_worker_id_by_name(name)
                                  fahrzeit = get_effective_fahrzeit(self.master_db, worker_id, bst_data['id'], bst_data.get('fahrzeit', 0.0))
                                  h += float(fahrzeit)
                        
                        total_hours += h
                    
                    is_unter_8h = (total_hours <= 8.0)
                    
                    for e in day_entries:
                        if not e.get('urlaub') and not e.get('krank'):
                            self.db.update_entry(e['id'], {'kg_8h': is_unter_8h})

            if errors:
                error_msg = f"{total_entries} Einträge verarbeitet.\n\nFehler:\n" + "\n".join(errors[:10])
                if len(errors) > 10:
                    error_msg += f"\n... und {len(errors) - 10} weitere Fehler"
                messagebox.showwarning("Hinweis", error_msg)

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern:\n{str(e)}")
            print(e) # For debugging

        self.update_month_view()
        self.update_day_view()

        if self.settings.get("auto_increment_day", False) and not delete_mode:
            last_day = max(days)

            if self.settings.get("skip_weekends", True):
                next_year, next_month, next_day = self.get_next_day_skip_weekend(jahr, monat, last_day)
            else:
                next_year, next_month, next_day = self.get_next_day(jahr, monat, last_day)

            if next_year != int(jahr):
                self.entry_year.delete(0, tk.END)
                self.entry_year.insert(0, str(next_year))
            if next_month != int(monat):
                self.entry_month.delete(0, tk.END)
                self.entry_month.insert(0, str(next_month))

            self.entry_day.delete(0, tk.END)
            self.entry_day.insert(0, str(next_day))

            self.update_weekday()

        should_clear_baustelle = True 
        if input_krank or input_urlaub:
             should_clear_baustelle = True
        
        self.clear_fields(clear_baustelle=should_clear_baustelle)

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
            self.entry_day.focus()

    def export_excel(self):
        try:
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

            if export_to_excel(jahr, monat, self.db, self.master_db):
                messagebox.showinfo("Erfolg", f"Daten für {monat:02d}/{jahr} nach Excel exportiert!")
            else:
                messagebox.showwarning("Warnung", "Keine Daten zum Exportieren vorhanden.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen:\n{str(e)}")

    def get_next_day(self, year, month, day):
        try:
            current_date = datetime(int(year), int(month), int(day))
            next_date = current_date + timedelta(days=1)
            return (next_date.year, next_date.month, next_date.day)
        except (ValueError, TypeError):
            # If invalid date, just increment day by 1
            return (year, month, day + 1)

    def get_next_day_skip_weekend(self, year, month, day):
        try:
            current_date = datetime(int(year), int(month), int(day))
            next_date = current_date + timedelta(days=1)

            while next_date.weekday() >= 5:
                next_date += timedelta(days=1)

            return (next_date.year, next_date.month, next_date.day)
        except (ValueError, TypeError):
            return (year, month, day + 1)

    def clear_fields(self, clear_baustelle=True):
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
        return [f for f in self.fields if f.winfo_ismapped()]

    def focus_next(self, event):
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
        self.name_autocomplete = AutocompleteEntry(
            self.entry_name,
            self.get_name_suggestions
        )
        self.baustelle_autocomplete = BaustelleAutocomplete(
            self.entry_bst,
            self.get_baustelle_suggestions
        )

    def get_name_suggestions(self):
        names_data = self.master_db.get_all_names()
        return [n['name'] for n in names_data]

    def get_baustelle_suggestions(self):
        return self.master_db.get_all_baustellen()

    def sort_month_tree(self, col):
        if col == self.month_sort_column:
            self.month_sort_reverse = not self.month_sort_reverse
        else:
            self.month_sort_column = col
            self.month_sort_reverse = False

        items = [(self.month_tree.set(item, col), item) for item in self.month_tree.get_children('')]

        if col in ('Tag', 'Stunden'):
            try:
                items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=self.month_sort_reverse)
            except ValueError:
                items.sort(reverse=self.month_sort_reverse)
        else:
            items.sort(reverse=self.month_sort_reverse)

        for index, (val, item) in enumerate(items):
            self.month_tree.move(item, '', index)

        for column in self.month_tree['columns']:
            heading_text = column
            if column == col:
                heading_text = f"{column} {'▼' if self.month_sort_reverse else '▲'}"
            self.month_tree.heading(column, text=heading_text)

    def sort_day_tree(self, col):
        if col == self.day_sort_column:
            self.day_sort_reverse = not self.day_sort_reverse
        else:
            self.day_sort_column = col
            self.day_sort_reverse = False
        items = [(self.day_tree.set(item, col), item) for item in self.day_tree.get_children('')]

        if col in ('Tag', 'Stunden'):
            try:
                items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=self.day_sort_reverse)
            except ValueError:
                items.sort(reverse=self.day_sort_reverse)
        else:
            items.sort(reverse=self.day_sort_reverse)

        for index, (val, item) in enumerate(items):
            self.day_tree.move(item, '', index)

        for column in self.day_tree['columns']:
            heading_text = column
            if column == col:
                heading_text = f"{column} {'▼' if self.day_sort_reverse else '▲'}"
            self.day_tree.heading(column, text=heading_text)

    def open_name_manager(self):
        NameManagerDialog(self.root)

    def open_baustelle_manager(self):
        BaustelleManagerDialog(self.root)

    def open_settings(self):
        dialog = SettingsDialog(self.root, self.settings, self.master_db)
        self.root.wait_window(dialog.dialog)
        self.settings.current_settings = self.settings.load()