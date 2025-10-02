import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import pandas as pd
import locale
# --- Variablen ---
entries = []

locale.setlocale(locale.LC_TIME, 'de_DE')  # Deutsch
# --- Funktionen ---
def update_weekday(*args):
    """Update Wochentag im Label neben Tag-Feld."""
    try:
        d = datetime(int(entry_year.get()), int(entry_month.get()), int(entry_day.get()))
        label_weekday.config(text=d.strftime("%a"))  # z.B. Mo, Tue
    except:
        label_weekday.config(text="-")

def toggle_skug():
    """Zeigt/versteckt SKUG Feld abhängig von Checkbox2."""
    if check_var2.get():
        label_skug.grid(row=7, column=0, sticky="e", padx=5, pady=2)
        entry_skug.grid(row=7, column=1, padx=5, pady=2)
    else:
        label_skug.grid_remove()
        entry_skug.grid_remove()

def submit():
    """Speichert die eingegebenen Daten und exportiert nach Excel."""
    data = {
        "Jahr": entry_year.get(),
        "Monat": entry_month.get(),
        "Tag": entry_day.get(),
        "Wochentag": label_weekday.cget("text"),
        "Name": entry_name.get(),
        "Stunden": entry_hours.get(),
        "CheckBox1": bool(check_var1.get()),
        "CheckBox2": bool(check_var2.get()),
        "SKUG": entry_skug.get() if check_var2.get() else "",
        "Baustelle": entry_bst.get()
    }
    entries.append(data)

    # Excel export
    df = pd.DataFrame(entries)
    df.to_excel("stundenliste.xlsx", index=False)

    messagebox.showinfo("Eingabe", f"Daten gespeichert:\n{data}")

    # Felder für nächste Eingabe vorbereiten
    entry_hours.delete(0, tk.END)
    entry_skug.delete(0, tk.END)
    entry_day.delete(0, tk.END)
    label_weekday.config(text="-")
    entry_day.focus()

# --- GUI Setup ---
root = tk.Tk()
root.title("Stunden-Eingabe")
root.geometry("500x400")

# Jahr
tk.Label(root, text="Jahr:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
entry_year = tk.Entry(root)
entry_year.grid(row=0, column=1)

# Monat
tk.Label(root, text="Monat:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
entry_month = tk.Entry(root)
entry_month.grid(row=1, column=1)

# Tag
tk.Label(root, text="Tag:").grid(row=2, column=0, sticky="e", padx=5, pady=2)
entry_day = tk.Entry(root)
entry_day.grid(row=2, column=1)
label_weekday = tk.Label(root, text="-")
label_weekday.grid(row=2, column=2)
entry_year.bind("<KeyRelease>", update_weekday)
entry_month.bind("<KeyRelease>", update_weekday)
entry_day.bind("<KeyRelease>", update_weekday)

# Name
tk.Label(root, text="Name:").grid(row=3, column=0, sticky="e", padx=5, pady=2)
entry_name = tk.Entry(root)
entry_name.grid(row=3, column=1)

# Stunden
tk.Label(root, text="Stunden:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
entry_hours = tk.Entry(root)
entry_hours.grid(row=4, column=1)

# CheckBox1
check_var1 = tk.IntVar()
check1 = tk.Checkbutton(root, text="CheckBox1", variable=check_var1)
check1.grid(row=5, column=0, columnspan=2, pady=2)

# CheckBox2
check_var2 = tk.IntVar()
check2 = tk.Checkbutton(root, text="CheckBox2", variable=check_var2, command=toggle_skug)
check2.grid(row=6, column=0, columnspan=2, pady=2)

# SKUG Feld (nur sichtbar wenn CheckBox2 aktiv)
label_skug = tk.Label(root, text="SKUG:")
entry_skug = tk.Entry(root)

# Baustelle
tk.Label(root, text="Baustelle:").grid(row=8, column=0, sticky="e", padx=5, pady=2)
entry_bst = tk.Entry(root)
entry_bst.grid(row=8, column=1)

# Button
btn_submit = tk.Button(root, text="Speichern", command=submit)
btn_submit.grid(row=9, column=0, columnspan=2, pady=20)

# --- Enter-Navigation ---
fields = [entry_year, entry_month, entry_day, entry_name, entry_hours, entry_skug, entry_bst]

def focus_next(event):
    widget = event.widget
    try:
        idx = fields.index(widget)
        # nächstes Feld überspringen, wenn Checkbox oder nicht sichtbar
        next_idx = idx + 1
        while next_idx < len(fields) and (not fields[next_idx].winfo_ismapped() or isinstance(fields[next_idx], tk.Checkbutton)):
            next_idx += 1
        if next_idx < len(fields):
            fields[next_idx].focus()
        else:
            submit()
    except ValueError:
        pass

for f in fields:
    f.bind("<Return>", focus_next)

root.mainloop()
