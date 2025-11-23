import tkinter as tk
from tkinter import ttk, messagebox
from master_data import MasterDataDatabase
from datatypes import WorkerTypes


class NameManagerDialog:
    """Dialog for managing names."""

    def __init__(self, parent):
        self.parent = parent
        self.db = MasterDataDatabase()
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Namen verwalten")
        self.dialog.geometry("400x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Position near parent window
        self.position_near_parent()

        self.create_widgets()
        self.refresh_list()

    def position_near_parent(self):
        """Position dialog near the parent window."""
        self.dialog.update_idletasks()
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()

        # Position to the right of parent, or centered if not enough space
        x = parent_x + parent_width + 10
        y = parent_y + 50

        self.dialog.geometry(f"+{x}+{y}")

    def create_widgets(self):
        """Create dialog widgets."""
        # Main frame
        main_frame = tk.Frame(self.dialog, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input frame
        input_frame = tk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(input_frame, text="Name:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_name = tk.Entry(input_frame, width=30)
        self.entry_name.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(input_frame, text="Typ:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.combo_worker_type = ttk.Combobox(input_frame, width=27, state="readonly")
        self.combo_worker_type['values'] = [wt.value for wt in WorkerTypes]
        self.combo_worker_type.current(0)  # Default to first type (Fest)
        self.combo_worker_type.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # New attributes
        self.var_kein_verpflegung = tk.BooleanVar()
        self.check_kein_verpflegung = tk.Checkbutton(input_frame, text="Kein Verpflegungsgeld", variable=self.var_kein_verpflegung)
        self.check_kein_verpflegung.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        self.var_keine_feiertag = tk.BooleanVar()
        self.check_keine_feiertag = tk.Checkbutton(input_frame, text="Keine Feiertagsstunden", variable=self.var_keine_feiertag)
        self.check_keine_feiertag.grid(row=3, column=1, sticky="w", padx=5, pady=2)

        tk.Label(input_frame, text="Wochenstunden:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_weekly_hours = tk.Entry(input_frame, width=10)
        self.entry_weekly_hours.insert(0, "0.0")
        self.entry_weekly_hours.grid(row=4, column=1, sticky="w", padx=5, pady=2)

        # Bind combobox change to toggle weekly hours
        self.combo_worker_type.bind("<<ComboboxSelected>>", self.toggle_weekly_hours)
        self.toggle_weekly_hours() # Initial state

        self.btn_add = tk.Button(input_frame, text="Hinzufügen", command=self.add_name)
        self.btn_add.grid(row=5, column=0, columnspan=2, pady=10)

        input_frame.grid_columnconfigure(1, weight=1)

        # List frame with scrollbar
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.btn_edit = tk.Button(button_frame, text="Bearbeiten", command=self.edit_name)
        self.btn_edit.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_delete = tk.Button(button_frame, text="Löschen", command=self.delete_name)
        self.btn_delete.pack(side=tk.LEFT)

        tk.Button(button_frame, text="Schließen", command=self.dialog.destroy).pack(side=tk.RIGHT)

        # Bindings
        self.entry_name.bind("<Return>", lambda e: self.add_name())
        self.listbox.bind("<Double-Button-1>", lambda e: self.edit_name())

    def toggle_weekly_hours(self, event=None):
        """Enable/disable weekly hours entry based on worker type."""
        worker_type = self.combo_worker_type.get()
        if worker_type == 'Fest': # Assuming 'Fest' is the value for permanent workers
            self.entry_weekly_hours.config(state='normal')
        else:
            self.entry_weekly_hours.delete(0, tk.END)
            self.entry_weekly_hours.insert(0, "0.0")
            self.entry_weekly_hours.config(state='disabled')

    def refresh_list(self):
        """Refresh the names list."""
        self.listbox.delete(0, tk.END)
        self.names_data = self.db.get_all_names()

        for name_entry in self.names_data:
            worker_type = name_entry.get('worker_type', 'Fest')
            display_text = f"{name_entry['name']} ({worker_type})"
            self.listbox.insert(tk.END, display_text)

    def add_name(self):
        """Add a new name."""
        name = self.entry_name.get().strip()
        worker_type = self.combo_worker_type.get()
        kein_verpflegung = self.var_kein_verpflegung.get()
        keine_feiertag = self.var_keine_feiertag.get()
        
        try:
            weekly_hours = float(self.entry_weekly_hours.get().strip())
        except ValueError:
            messagebox.showerror("Fehler", "Wochenstunden muss eine Zahl sein.")
            return

        if not name:
            messagebox.showwarning("Warnung", "Bitte geben Sie einen Namen ein.")
            return

        result = self.db.add_name(name, worker_type, kein_verpflegung, keine_feiertag, weekly_hours)

        if result:
            messagebox.showinfo("Erfolg", f"Name '{name}' ({worker_type}) wurde hinzugefügt.")
            self.entry_name.delete(0, tk.END)
            self.combo_worker_type.current(0)  # Reset to default
            self.var_kein_verpflegung.set(False)
            self.var_keine_feiertag.set(False)
            self.entry_weekly_hours.delete(0, tk.END)
            self.entry_weekly_hours.insert(0, "0.0")
            self.toggle_weekly_hours()
            self.refresh_list()
        else:
            messagebox.showerror("Fehler", f"Name '{name}' existiert bereits.")

    def edit_name(self):
        """Edit selected name."""
        selection = self.listbox.curselection()

        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie einen Namen aus.")
            return

        index = selection[0]
        name_data = self.names_data[index]
        old_name = name_data['name']
        old_worker_type = name_data.get('worker_type', 'Fest')
        old_kein_verpflegung = bool(name_data.get('kein_verpflegungsgeld', 0))
        old_keine_feiertag = bool(name_data.get('keine_feiertagssstunden', 0))
        old_weekly_hours = name_data.get('weekly_hours', 0.0)

        # Create edit dialog
        edit_dialog = tk.Toplevel(self.dialog)
        edit_dialog.title("Name bearbeiten")
        edit_dialog.geometry("400x300")
        edit_dialog.transient(self.dialog)
        edit_dialog.grab_set()

        # Position near parent dialog
        edit_dialog.update_idletasks()
        dialog_x = self.dialog.winfo_x()
        dialog_y = self.dialog.winfo_y()
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        x = dialog_x + (dialog_width - 400) // 2  # Center horizontally
        y = dialog_y + dialog_height // 4 
        edit_dialog.geometry(f"+{x}+{y}")

        tk.Label(edit_dialog, text="Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        entry_edit = tk.Entry(edit_dialog, width=30)
        entry_edit.insert(0, old_name)
        entry_edit.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        entry_edit.select_range(0, tk.END)
        entry_edit.focus()

        tk.Label(edit_dialog, text="Typ:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        combo_edit_type = ttk.Combobox(edit_dialog, width=27, state="readonly")
        combo_edit_type['values'] = [wt.value for wt in WorkerTypes]
        # Set current value
        try:
            combo_edit_type.current([wt.value for wt in WorkerTypes].index(old_worker_type))
        except ValueError:
            combo_edit_type.current(0)  # Default to first if not found
        combo_edit_type.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        var_edit_kein_verpflegung = tk.BooleanVar(value=old_kein_verpflegung)
        check_edit_kein_verpflegung = tk.Checkbutton(edit_dialog, text="Kein Verpflegungsgeld", variable=var_edit_kein_verpflegung)
        check_edit_kein_verpflegung.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        var_edit_keine_feiertag = tk.BooleanVar(value=old_keine_feiertag)
        check_edit_keine_feiertag = tk.Checkbutton(edit_dialog, text="Keine Feiertagsstunden", variable=var_edit_keine_feiertag)
        check_edit_keine_feiertag.grid(row=3, column=1, sticky="w", padx=5, pady=2)

        tk.Label(edit_dialog, text="Wochenstunden:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        entry_edit_weekly_hours = tk.Entry(edit_dialog, width=10)
        entry_edit_weekly_hours.insert(0, str(old_weekly_hours))
        entry_edit_weekly_hours.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        edit_dialog.grid_columnconfigure(1, weight=1)
        
        def toggle_edit_weekly_hours(event=None):
            if combo_edit_type.get() == 'Fest':
                entry_edit_weekly_hours.config(state='normal')
            else:
                entry_edit_weekly_hours.config(state='disabled')
        
        combo_edit_type.bind("<<ComboboxSelected>>", toggle_edit_weekly_hours)
        toggle_edit_weekly_hours()

        def save_edit():
            new_name = entry_edit.get().strip()
            new_worker_type = combo_edit_type.get()
            new_kein_verpflegung = var_edit_kein_verpflegung.get()
            new_keine_feiertag = var_edit_keine_feiertag.get()
            
            try:
                new_weekly_hours = float(entry_edit_weekly_hours.get().strip())
            except ValueError:
                messagebox.showerror("Fehler", "Wochenstunden muss eine Zahl sein.")
                return

            if not new_name:
                messagebox.showwarning("Warnung", "Name darf nicht leer sein.")
                return

            if self.db.update_name(name_data['id'], new_name, new_worker_type, 
                                   new_kein_verpflegung, new_keine_feiertag, new_weekly_hours):
                messagebox.showinfo("Erfolg", f"Name wurde zu '{new_name}' ({new_worker_type}) geändert.")
                edit_dialog.destroy()
                self.refresh_list()
            else:
                messagebox.showerror("Fehler", "Name konnte nicht aktualisiert werden (möglicherweise existiert er bereits).")

        btn_frame = tk.Frame(edit_dialog)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)

        tk.Button(btn_frame, text="Speichern", command=save_edit).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Abbrechen", command=edit_dialog.destroy).pack(side=tk.LEFT, padx=5)

        entry_edit.bind("<Return>", lambda e: save_edit())

    def delete_name(self):
        """Delete selected name."""
        selection = self.listbox.curselection()

        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie einen Namen aus.")
            return

        index = selection[0]
        name_data = self.names_data[index]

        if messagebox.askyesno("Bestätigen", f"Möchten Sie '{name_data['name']}' wirklich löschen?"):
            if self.db.delete_name(name_data['id']):
                messagebox.showinfo("Erfolg", "Name wurde gelöscht.")
                self.refresh_list()
            else:
                messagebox.showerror("Fehler", "Name konnte nicht gelöscht werden.")


class BaustelleManagerDialog:
    """Dialog for managing baustellen."""

    def __init__(self, parent):
        self.parent = parent
        self.db = MasterDataDatabase()
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Baustellen verwalten")
        self.dialog.geometry("600x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Position near parent window
        self.position_near_parent()

        self.create_widgets()
        self.refresh_list()

    def position_near_parent(self):
        """Position dialog near the parent window."""
        self.dialog.update_idletasks()
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()

        # Position to the right of parent, or centered if not enough space
        x = parent_x + parent_width + 10
        y = parent_y + 50

        self.dialog.geometry(f"+{x}+{y}")

    def create_widgets(self):
        """Create dialog widgets."""
        # Main frame
        main_frame = tk.Frame(self.dialog, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input frame
        input_frame = tk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(input_frame, text="Nummer:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.entry_nummer = tk.Entry(input_frame, width=15)
        self.entry_nummer.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        tk.Label(input_frame, text="Name:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.entry_name = tk.Entry(input_frame, width=30)
        self.entry_name.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(input_frame, text="Verpflegungsgeld:").grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.entry_verpflegung = tk.Entry(input_frame, width=15)
        self.entry_verpflegung.insert(0, "0.0")
        self.entry_verpflegung.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        tk.Label(input_frame, text="Fahrzeit (h):").grid(row=3, column=0, sticky="e", padx=5, pady=2)
        self.entry_fahrzeit = tk.Entry(input_frame, width=15)
        self.entry_fahrzeit.insert(0, "0.0")
        self.entry_fahrzeit.grid(row=3, column=1, sticky="w", padx=5, pady=2)

        tk.Label(input_frame, text="Distanz (km):").grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.entry_distance = tk.Entry(input_frame, width=15)
        self.entry_distance.insert(0, "0.0")
        self.entry_distance.grid(row=4, column=1, sticky="w", padx=5, pady=2)

        self.btn_add = tk.Button(input_frame, text="Hinzufügen", command=self.add_baustelle)
        self.btn_add.grid(row=5, column=0, columnspan=2, pady=10)

        input_frame.grid_columnconfigure(1, weight=1)

        # Treeview frame
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview with scrollbar
        columns = ('Nummer', 'Name', 'Verpflegungsgeld', 'Fahrzeit', 'Distanz')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        self.tree.heading('Nummer', text='Nummer')
        self.tree.heading('Name', text='Name')
        self.tree.heading('Verpflegungsgeld', text='Verpflegungsgeld (€)')
        self.tree.heading('Fahrzeit', text='Fahrzeit (h)')
        self.tree.heading('Distanz', text='Distanz (km)')

        self.tree.column('Nummer', width=80)
        self.tree.column('Name', width=200)
        self.tree.column('Verpflegungsgeld', width=120)
        self.tree.column('Fahrzeit', width=80)
        self.tree.column('Distanz', width=80)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.btn_edit = tk.Button(button_frame, text="Bearbeiten", command=self.edit_baustelle)
        self.btn_edit.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_delete = tk.Button(button_frame, text="Löschen", command=self.delete_baustelle)
        self.btn_delete.pack(side=tk.LEFT)

        tk.Button(button_frame, text="Schließen", command=self.dialog.destroy).pack(side=tk.RIGHT)

        # Bindings
        self.entry_verpflegung.bind("<Return>", lambda e: self.add_baustelle())
        self.tree.bind("<Double-Button-1>", lambda e: self.edit_baustelle())

    def refresh_list(self):
        """Refresh the baustellen list."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.baustellen_data = self.db.get_all_baustellen()

        for baustelle in self.baustellen_data:
            self.tree.insert('', tk.END, values=(
                baustelle['nummer'],
                baustelle['name'],
                f"{baustelle['verpflegungsgeld']:.2f}",
                f"{baustelle.get('fahrzeit', 0.0):.2f}",
                f"{baustelle.get('distance_km', 0.0):.2f}"
            ), tags=(baustelle['id'],))

    def add_baustelle(self):
        """Add a new baustelle."""
        nummer = self.entry_nummer.get().strip()
        name = self.entry_name.get().strip()
        verpflegungsgeld_str = self.entry_verpflegung.get().strip()
        fahrzeit_str = self.entry_fahrzeit.get().strip()
        distance_str = self.entry_distance.get().strip()

        if not nummer or not name:
            messagebox.showwarning("Warnung", "Bitte füllen Sie Nummer und Name aus.")
            return

        try:
            verpflegungsgeld = float(verpflegungsgeld_str)
            fahrzeit = float(fahrzeit_str)
            distance_km = float(distance_str)
        except ValueError:
            messagebox.showerror("Fehler", "Verpflegungsgeld, Fahrzeit und Distanz müssen Zahlen sein.")
            return

        result = self.db.add_baustelle(nummer, name, verpflegungsgeld, fahrzeit, distance_km)

        if result:
            messagebox.showinfo("Erfolg", f"Baustelle '{nummer} - {name}' wurde hinzugefügt.")
            self.entry_nummer.delete(0, tk.END)
            self.entry_name.delete(0, tk.END)
            self.entry_verpflegung.delete(0, tk.END)
            self.entry_verpflegung.insert(0, "0.0")
            self.entry_fahrzeit.delete(0, tk.END)
            self.entry_fahrzeit.insert(0, "0.0")
            self.entry_distance.delete(0, tk.END)
            self.entry_distance.insert(0, "0.0")
            self.refresh_list()
        else:
            messagebox.showerror("Fehler", "Baustelle existiert bereits.")

    def edit_baustelle(self):
        """Edit selected baustelle."""
        selection = self.tree.selection()

        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie eine Baustelle aus.")
            return

        item = selection[0]
        baustelle_id = int(self.tree.item(item)['tags'][0])

        # Find the baustelle data
        baustelle_data = next((b for b in self.baustellen_data if b['id'] == baustelle_id), None)

        if not baustelle_data:
            return

        # Create edit dialog
        edit_dialog = tk.Toplevel(self.dialog)
        edit_dialog.title("Baustelle bearbeiten")
        edit_dialog.geometry("350x250")
        edit_dialog.transient(self.dialog)
        edit_dialog.grab_set()

        # Position near parent dialog
        edit_dialog.update_idletasks()
        dialog_x = self.dialog.winfo_x()
        dialog_y = self.dialog.winfo_y()
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        x = dialog_x + (dialog_width - 350) // 2  # Center horizontally
        y = dialog_y + dialog_height // 3  # One third down
        edit_dialog.geometry(f"+{x}+{y}")

        tk.Label(edit_dialog, text="Nummer:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        entry_nummer = tk.Entry(edit_dialog, width=20)
        entry_nummer.insert(0, baustelle_data['nummer'])
        entry_nummer.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(edit_dialog, text="Name:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        entry_name = tk.Entry(edit_dialog, width=30)
        entry_name.insert(0, baustelle_data['name'])
        entry_name.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(edit_dialog, text="Verpflegungsgeld:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        entry_verpflegung = tk.Entry(edit_dialog, width=20)
        entry_verpflegung.insert(0, str(baustelle_data['verpflegungsgeld']))
        entry_verpflegung.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(edit_dialog, text="Fahrzeit (h):").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        entry_fahrzeit = tk.Entry(edit_dialog, width=20)
        entry_fahrzeit.insert(0, str(baustelle_data.get('fahrzeit', 0.0)))
        entry_fahrzeit.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(edit_dialog, text="Distanz (km):").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        entry_distance = tk.Entry(edit_dialog, width=20)
        entry_distance.insert(0, str(baustelle_data.get('distance_km', 0.0)))
        entry_distance.grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        edit_dialog.grid_columnconfigure(1, weight=1)

        def save_edit():
            nummer = entry_nummer.get().strip()
            name = entry_name.get().strip()
            verpflegungsgeld_str = entry_verpflegung.get().strip()
            fahrzeit_str = entry_fahrzeit.get().strip()
            distance_str = entry_distance.get().strip()

            if not nummer or not name:
                messagebox.showwarning("Warnung", "Nummer und Name dürfen nicht leer sein.")
                return

            try:
                verpflegungsgeld = float(verpflegungsgeld_str)
                fahrzeit = float(fahrzeit_str)
                distance_km = float(distance_str)
            except ValueError:
                messagebox.showerror("Fehler", "Verpflegungsgeld, Fahrzeit und Distanz müssen Zahlen sein.")
                return

            if self.db.update_baustelle(baustelle_id, nummer, name, verpflegungsgeld, fahrzeit, distance_km):
                messagebox.showinfo("Erfolg", "Baustelle wurde aktualisiert.")
                edit_dialog.destroy()
                self.refresh_list()
            else:
                messagebox.showerror("Fehler", "Baustelle konnte nicht aktualisiert werden.")

        btn_frame = tk.Frame(edit_dialog)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)

        tk.Button(btn_frame, text="Speichern", command=save_edit).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Abbrechen", command=edit_dialog.destroy).pack(side=tk.LEFT, padx=5)

        entry_verpflegung.bind("<Return>", lambda e: save_edit())

    def delete_baustelle(self):
        """Delete selected baustelle."""
        selection = self.tree.selection()

        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie eine Baustelle aus.")
            return

        item = selection[0]
        baustelle_id = int(self.tree.item(item)['tags'][0])

        # Find the baustelle data
        baustelle_data = next((b for b in self.baustellen_data if b['id'] == baustelle_id), None)

        if not baustelle_data:
            return

        if messagebox.askyesno("Bestätigen",
                               f"Möchten Sie Baustelle '{baustelle_data['nummer']} - {baustelle_data['name']}' wirklich löschen?"):
            if self.db.delete_baustelle(baustelle_id):
                messagebox.showinfo("Erfolg", "Baustelle wurde gelöscht.")
                self.refresh_list()
            else:
                messagebox.showerror("Fehler", "Baustelle konnte nicht gelöscht werden.")
