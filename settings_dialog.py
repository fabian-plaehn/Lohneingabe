import tkinter as tk
from tkinter import ttk, messagebox
import json
import os


class Settings:
    """Handles loading and saving settings to JSON file."""

    def __init__(self, settings_file="settings.json"):
        self.settings_file = settings_file
        self.default_settings = {
            "auto_increment_day": False,
            "skip_weekends": True,
            "cursor_jump_target": "Tag"
        }
        self.current_settings = self.load()

    def load(self):
        """Load settings from JSON file or return defaults."""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    settings = self.default_settings.copy()
                    settings.update(loaded)
                    return settings
            except (json.JSONDecodeError, IOError) as e:
                print(f"Error loading settings: {e}")
                return self.default_settings.copy()
        else:
            return self.default_settings.copy()

    def save(self, settings_dict):
        """Save settings to JSON file."""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_dict, f, indent=4, ensure_ascii=False)
            self.current_settings = settings_dict
            return True
        except IOError as e:
            print(f"Error saving settings: {e}")
            return False

    def get(self, key, default=None):
        """Get a setting value."""
        return self.current_settings.get(key, default)


class SettingsDialog:
    """Settings dialog window."""

    def __init__(self, parent, settings_manager, master_db):
        self.parent = parent
        self.settings_manager = settings_manager
        self.master_db = master_db
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Einstellungen")
        self.dialog.geometry("600x550")
        self.dialog.resizable(False, False)

        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Position near parent
        self.position_near_parent()

        # Variables for settings
        self.auto_increment_var = tk.BooleanVar(
            value=self.settings_manager.get("auto_increment_day", False)
        )
        self.skip_weekends_var = tk.BooleanVar(
            value=self.settings_manager.get("skip_weekends", True)
        )
        self.cursor_target_var = tk.StringVar(
            value=self.settings_manager.get("cursor_jump_target", "Tag")
        )

        # Add traces to auto-save when settings change
        self.auto_increment_var.trace_add("write", self.auto_save_settings)
        self.skip_weekends_var.trace_add("write", self.auto_save_settings)
        self.cursor_target_var.trace_add("write", self.auto_save_settings)

        # SKUG settings variables
        self.skug_vars = {}
        self.load_skug_settings()

        self.create_widgets()

    def position_near_parent(self):
        """Position dialog near the parent window."""
        self.dialog.update_idletasks()
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()

        # Center over parent
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        self.dialog.geometry(f"+{x}+{y}")

    def load_skug_settings(self):
        """Load SKUG settings from database."""
        settings = self.master_db.get_skug_settings()

        # Create StringVars for each setting
        days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']
        seasons = ['winter', 'summer']

        for season in seasons:
            for day in days:
                key = f"{season}_{day}"
                self.skug_vars[key] = tk.StringVar(value=str(settings.get(key, 8.0)))

    def create_widgets(self):
        """Create the settings interface."""
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Auto-increment settings section
        increment_frame = tk.LabelFrame(main_frame, text="Tag-Erhöhung", padx=10, pady=10)
        increment_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Checkbutton(
            increment_frame,
            text="Tag automatisch erhöhen nach Eingabe",
            variable=self.auto_increment_var
        ).pack(anchor="w")

        tk.Checkbutton(
            increment_frame,
            text="Wochenenden überspringen",
            variable=self.skip_weekends_var
        ).pack(anchor="w", pady=(5, 0))

        # Cursor jump settings section
        cursor_frame = tk.LabelFrame(main_frame, text="Cursor-Sprung", padx=10, pady=10)
        cursor_frame.pack(fill=tk.X, pady=(0, 10))

        cursor_label = tk.Label(cursor_frame, text="Nach Eingabe springen zu:")
        cursor_label.pack(anchor="w")

        cursor_options = ["Tag", "Name", "Stunden", "Baustelle"]
        cursor_dropdown = ttk.Combobox(
            cursor_frame,
            textvariable=self.cursor_target_var,
            values=cursor_options,
            width=15,
            state="readonly"
        )
        cursor_dropdown.pack(anchor="w", pady=(5, 0))

        # SKUG settings section
        skug_frame = tk.LabelFrame(main_frame, text="SKUG Einstellungen", padx=10, pady=10)
        skug_frame.pack(fill=tk.X, pady=(0, 10))

        # Info label
        info_label = tk.Label(
            skug_frame,
            text="Winter: Dezember - März | Sommer: April - November",
            font=("Arial", 8),
            fg="gray"
        )
        info_label.grid(row=0, column=0, columnspan=4, pady=(0, 5))

        # Headers
        tk.Label(skug_frame, text="", width=10).grid(row=1, column=0)
        tk.Label(skug_frame, text="Winter", font=("Arial", 9, "bold")).grid(row=1, column=1)
        tk.Label(skug_frame, text="Sommer", font=("Arial", 9, "bold")).grid(row=1, column=2)

        # Days
        day_labels = {
            'monday': 'Montag',
            'tuesday': 'Dienstag',
            'wednesday': 'Mittwoch',
            'thursday': 'Donnerstag',
            'friday': 'Freitag'
        }

        row = 2
        for day_key, day_label in day_labels.items():
            tk.Label(skug_frame, text=day_label + ":", anchor="w").grid(row=row, column=0, sticky="w", pady=2)

            # Winter entry
            winter_entry = tk.Entry(skug_frame, textvariable=self.skug_vars[f'winter_{day_key}'], width=8)
            winter_entry.grid(row=row, column=1, padx=5, pady=2)

            # Summer entry
            summer_entry = tk.Entry(skug_frame, textvariable=self.skug_vars[f'summer_{day_key}'], width=8)
            summer_entry.grid(row=row, column=2, padx=5, pady=2)

            row += 1

        # Save SKUG button
        btn_save_skug = tk.Button(
            skug_frame,
            text="SKUG Einstellungen speichern",
            command=self.save_skug_settings,
            width=25
        )
        btn_save_skug.grid(row=row, column=0, columnspan=3, pady=(10, 0))

        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        btn_reset = tk.Button(
            button_frame,
            text="Zurücksetzen",
            command=self.reset_to_defaults,
            width=15
        )
        btn_reset.pack(side=tk.LEFT)

        btn_close = tk.Button(
            button_frame,
            text="Schließen",
            command=self.dialog.destroy,
            width=15
        )
        btn_close.pack(side=tk.RIGHT)

    def auto_save_settings(self, *args):
        """Automatically save settings when any setting changes."""
        settings_dict = {
            "auto_increment_day": self.auto_increment_var.get(),
            "skip_weekends": self.skip_weekends_var.get(),
            "cursor_jump_target": self.cursor_target_var.get()
        }
        self.settings_manager.save(settings_dict)

    def save_skug_settings(self):
        """Save SKUG settings to database."""
        try:
            settings = {}
            for key, var in self.skug_vars.items():
                value = float(var.get())
                if value < 0 or value > 24:
                    messagebox.showerror(
                        "Fehler",
                        f"Ungültiger Wert für {key}: {value}\nWert muss zwischen 0 und 24 liegen.",
                        parent=self.dialog
                    )
                    return
                settings[key] = value

            if self.master_db.update_skug_settings(settings):
                messagebox.showinfo(
                    "Erfolg",
                    "SKUG Einstellungen wurden gespeichert.",
                    parent=self.dialog
                )
            else:
                messagebox.showerror(
                    "Fehler",
                    "SKUG Einstellungen konnten nicht gespeichert werden.",
                    parent=self.dialog
                )
        except ValueError:
            messagebox.showerror(
                "Fehler",
                "Bitte geben Sie gültige Zahlenwerte ein.",
                parent=self.dialog
            )

    def reset_to_defaults(self):
        """Reset all settings to default values."""
        if messagebox.askyesno("Zurücksetzen",
                              "Alle Einstellungen auf Standardwerte zurücksetzen?",
                              parent=self.dialog):
            # Temporarily remove traces to avoid multiple saves
            self.auto_increment_var.trace_remove("write", self.auto_increment_var.trace_info()[0][1])
            self.skip_weekends_var.trace_remove("write", self.skip_weekends_var.trace_info()[0][1])
            self.cursor_target_var.trace_remove("write", self.cursor_target_var.trace_info()[0][1])

            # Set default values
            self.auto_increment_var.set(False)
            self.skip_weekends_var.set(True)
            self.cursor_target_var.set("Tag")

            # Re-add traces
            self.auto_increment_var.trace_add("write", self.auto_save_settings)
            self.skip_weekends_var.trace_add("write", self.auto_save_settings)
            self.cursor_target_var.trace_add("write", self.auto_save_settings)

            # Save the defaults
            self.auto_save_settings()
