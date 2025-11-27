from openpyxl import Workbook
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import calendar
from database import Database
from master_data import MasterDataDatabase
from utils import get_days_of_krank, get_days_of_urlaub, get_fahrstunden_for_name, get_normal_hours_per_month, get_verpflegungsgeld_for_name, is_holiday, is_weekend, calculate_skug
from datatypes import WorkerTypes

def AddBorders(border_one:Border, border_two:Border) -> Border:
    sides = ['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal']
    border_kwargs = {}
    for side in sides:
        try: side_one = getattr(border_one, side)
        except AttributeError: side_one = None
        try: side_two = getattr(border_two, side)
        except AttributeError: side_two = None
        if side_one is None and side_two is None: continue
        
        if side_one is not None:
            border_kwargs[side] = side_one if side_one.style is not None else side_two
        else:
            border_kwargs[side] = side_two
    return Border(**border_kwargs)

def addLattice(min_row, max_row, min_col, max_col, ws: Workbook):
    print("add lattice")
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = AddBorders(cell.border, Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')))


def set_create_border(min_row, max_row, min_col, max_col, side_style, ws: Workbook):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if cell.row == min_row:
                cell.border = AddBorders(cell.border, Border(top=side_style))
            if cell.row == max_row:
                cell.border = AddBorders(cell.border, Border(bottom=side_style))
            if cell.column == min_col:
                cell.border = AddBorders(cell.border, Border(left=side_style))
            if cell.column == max_col:
                cell.border = AddBorders(cell.border, Border(right=side_style))
       
SKUG_COLOR = "32a852"
UNTER_8H_COLOR = "42c8f5"
AN_AB_COLOR = "a83232"
FREE_DAY_COLOR = "FFA500"  # Orange color for weekends and holidays
                
def export_to_excel(year:int, month:int, db:Database, master_db: MasterDataDatabase, filename: str = None):
    """
    Export data for a specific month to Excel with custom formatting.

    Layout:
    - Starts at cell A3
    - Datum column (2x2 merged) with dates listed below
    - 9 names with Std./Bst. columns each
    - Repeat Datum column after every 9 names
    - Summary rows after last day
    """
    if filename is None:
        filename = f"stundenliste_{year}_{month:02d}.xlsx"

    # Get all entries for the month
    entries = db.get_entries_by_date(year, month)

    if not entries:
        print(f"No data to export for {year}-{month:02d}")
        return False

    # Get number of days in month
    num_days = calendar.monthrange(year, month)[1]

    # Get unique names from entries, sorted
    unique_names = master_db.get_all_names_list() # sorted(set(entry['name'] for entry in entries))
    
    # Create a lookup for person data
    all_persons = master_db.get_all_names()
    person_lookup = {p['name']: p for p in all_persons}
    
    if not unique_names:
        print("No names found in entries")
        return False

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}-{month:02d}"

    # Define thick border style
    thick_border = Border(
        left=Side(style='thick'),
        right=Side(style='thick'),
        top=Side(style='thick'),
        bottom=Side(style='thick')
    )
    
    # Track current column position
    current_col = 1  # Start at column A
    names_per_section = 9

    # Calculate how many sections we need
    num_sections = (len(unique_names) + names_per_section - 1) // names_per_section

    for section_idx in range(num_sections):
        # Get names for this section
        start_idx = section_idx * names_per_section
        end_idx = min(start_idx + names_per_section, len(unique_names))
        section_names = unique_names[start_idx:end_idx]

        # Write "Datum" header (2x2 merged cell starting at row 3)
        datum_col = current_col
        ws.merge_cells(start_row=3, start_column=datum_col, end_row=4, end_column=datum_col+1)
        datum_cell = ws.cell(row=3, column=datum_col)
        datum_cell.value = "Datum"
        datum_cell.alignment = Alignment(horizontal='center', vertical='center')
        datum_cell.font = Font(bold=True)
        
        info_cell = ws.cell(row=1, column=datum_col)
        info_cell.value = f"Stundenliste - {calendar.month_name[month]} {year}"

        # Apply thick border to Datum header
        for row in range(3, 5):
            for col in range(datum_col, datum_col + 2):
                ws.cell(row=row, column=col).border = thick_border

        # Write dates (starting at row 5)
        for day in range(1, num_days + 1):
            row = 5 + day - 1
            ws.merge_cells(start_row=row, start_column=datum_col, end_row=row, end_column=datum_col+1)
            date_cell = ws.cell(row=row, column=datum_col)
            date_cell.value = f"{day}."
            date_cell.alignment = Alignment(horizontal='center', vertical='center')

            # Color date cell if it's a weekend or holiday
            if is_weekend(year, month, day) or is_holiday(year, month, day):
                for col in range(datum_col, datum_col + 2):
                    ws.cell(row=row, column=col).fill = openpyxl.styles.PatternFill(
                        start_color=FREE_DAY_COLOR,
                        end_color=FREE_DAY_COLOR,
                        fill_type="solid"
                    )

        # Apply thick border around all dates
        set_create_border(
            min_row=5,
            max_row=5 + num_days - 1,
            min_col=datum_col,
            max_col=datum_col + 1,
            side_style=Side(border_style='thick'),
            ws=ws
        )

        # Move to name columns
        current_col += 2

        # Write each name column
        for name in section_names:
            name_col = current_col

            # Write name (1x2 merged cell in row 3)
            ws.merge_cells(start_row=3, start_column=name_col, end_row=3, end_column=name_col+1)
            name_cell = ws.cell(row=3, column=name_col)
            name_cell.value = name
            name_cell.alignment = Alignment(horizontal='center', vertical='center')
            name_cell.font = Font(bold=True)
            person_data = person_lookup.get(name, {})
            worker_type = person_data.get('worker_type', 'Fest')
            kein_verpflegung = bool(person_data.get('kein_verpflegungsgeld', 0))
            keine_feiertag = bool(person_data.get('keine_feiertagssstunden', 0))
            if not kein_verpflegung:
                # color the name blue
                name_cell.fill = openpyxl.styles.PatternFill(start_color=UNTER_8H_COLOR, end_color=UNTER_8H_COLOR, fill_type='solid')


            # Apply thick border to name header
            set_create_border(
                min_row=3,
                max_row=4,
                min_col=name_col,
                max_col=name_col + 1,
                side_style=Side(style='thick'),
                ws=ws
            )
            
            # Write "Std." and "Bst." in row 4
            std_cell = ws.cell(row=4, column=name_col)
            std_cell.value = "Std."
            std_cell.alignment = Alignment(horizontal='center', vertical='center')
            #std_cell.border = thick_border

            bst_cell = ws.cell(row=4, column=name_col + 1)
            bst_cell.value = "Bst."
            bst_cell.alignment = Alignment(horizontal='center', vertical='center')
            #bst_cell.border = thick_border

            # Apply thick border around Std./Bst. Data
            set_create_border(
                min_row=5,
                max_row=5+num_days-1,
                min_col=name_col,
                max_col=name_col + 1,
                side_style=Side(style='thick'),
                ws=ws
            )
            
            person_data = person_lookup.get(name, {})
            worker_type = person_data.get('worker_type', 'Fest')
            name_entries = [e for e in entries if e['name'] == name]
            gesamtstunden = sum(e['stunden'] for e in name_entries)
     
            # Fill in data for each day
            for day in range(1, num_days + 1):
                row = 5 + day - 1

                # Check if this day is a weekend or holiday
                is_free_day = is_weekend(year, month, day) or is_holiday(year, month, day)
                is_bank_holiday = is_holiday(year, month, day) and not is_weekend(year, month, day)

                std_cell_data = ws.cell(row=row, column=name_col)
                bst_cell_data = ws.cell(row=row, column=name_col + 1)

                entry = next((e for e in entries if e['name'] == name and e['tag'] == day), None)
                if is_free_day:
                    # Color orange for weekends and holidays
                    std_cell_data.fill = openpyxl.styles.PatternFill(start_color=FREE_DAY_COLOR, end_color=FREE_DAY_COLOR, fill_type="solid")
                    bst_cell_data.fill = openpyxl.styles.PatternFill(start_color=FREE_DAY_COLOR, end_color=FREE_DAY_COLOR, fill_type="solid")
                elif entry and entry.get('travel_status'):
                    std_cell_data.fill = openpyxl.styles.PatternFill(start_color=AN_AB_COLOR, end_color=AN_AB_COLOR, fill_type="solid")
                    bst_cell_data.fill = openpyxl.styles.PatternFill(start_color=AN_AB_COLOR, end_color=AN_AB_COLOR, fill_type="solid")
                elif entry and entry.get('skug'):
                    std_cell_data.fill = openpyxl.styles.PatternFill(start_color=SKUG_COLOR, end_color=SKUG_COLOR, fill_type="solid")
                    bst_cell_data.fill = openpyxl.styles.PatternFill(start_color=SKUG_COLOR, end_color=SKUG_COLOR, fill_type="solid")
                elif entry and entry.get('unter_8h'):
                    std_cell_data.fill = openpyxl.styles.PatternFill(start_color=UNTER_8H_COLOR, end_color=UNTER_8H_COLOR, fill_type="solid")
                    bst_cell_data.fill = openpyxl.styles.PatternFill(start_color=UNTER_8H_COLOR, end_color=UNTER_8H_COLOR, fill_type="solid")
                
                std_value = ""
                bst_value = ""

                # If it's a bank holiday (not weekend), automatically fill F and 940
                if is_bank_holiday and not entry and worker_type == WorkerTypes.Gewerblich:
                    if keine_feiertag:
                        std_value = ""
                        bst_value = ""
                    else:
                        std_value = "F"
                        bst_value = "940"
                elif entry:
                    std_value = entry.get('stunden', '')
                    baustelle = entry.get('baustelle', '')
                    # Extract number from baustelle (format: "Nummer - Name")
                    if baustelle and ' - ' in baustelle:
                        bst_value = baustelle.split(' - ')[0]
                    elif baustelle:
                        bst_value = baustelle
                        bst_cell_data.number_format = "0"

                    # If stunden = 0, check if its urlaub or krank
                    if std_value == 0:
                        if worker_type==WorkerTypes.Fest and gesamtstunden >0:
                            weekly_hours = person_data.get('weekly_hours', 0.0)
                            if entry.get('urlaub'):
                                std_value = weekly_hours/5.0
                                bst_value = "Urlaub"
                            elif entry.get('krank'):
                                std_value = weekly_hours/5.0
                                bst_value = "Krank"
                        else:
                            if entry.get('urlaub'):
                                std_value = "Urlaub"
                                bst_value = ""
                            elif entry.get('krank'):
                                std_value = "Krank"
                                bst_value = ""
                elif is_bank_holiday and worker_type==WorkerTypes.Fest and gesamtstunden>0:
                    weekly_hours = person_data.get('weekly_hours', 0.0)
                    std_value = weekly_hours/5.0
                    bst_value = "F"
                elif is_bank_holiday:
                    std_value = "F"
                    bst_value = ""

                # Write values
                std_cell_data.value = std_value
                std_cell_data.number_format = "0.00"
                std_cell_data.alignment = Alignment(horizontal='center', vertical='center')

                bst_cell_data.value = bst_value
                bst_cell_data.alignment = Alignment(horizontal='center', vertical='center')

            # Move to next name
            current_col += 2

        # Add summary rows under this section
        summary_start_row = 5 + num_days
        summary_labels = [
            "Gesamtstunden",
            "Feiertag",
            "Urlaubsstunden",
            "Krankstunden",
            "SKUG",
            "Summe",
            "Mehr-/Minderstd",
            "Fahrstunden",
            "V.-Zuschuss [€]"
        ]
        
        # Apply thick border to summary labels
        set_create_border(
            min_row=summary_start_row,
            max_row=summary_start_row+len(summary_labels)-1,
            min_col=datum_col,
            max_col=datum_col,
            side_style=Side(style='thick'),
            ws=ws
        )

        # Write summary labels under Datum column
        for idx, label in enumerate(summary_labels):
            row = summary_start_row + idx
            label_cell = ws.cell(row=row, column=datum_col)
            label_cell.value = label
            label_cell.font = Font(bold=True)
            label_cell.alignment = Alignment(horizontal='left', vertical='center')

            # Merge across both Datum columns
            ws.merge_cells(start_row=row, start_column=datum_col, end_row=row, end_column=datum_col+1)

        # Add legend for colors
        row = summary_start_row + len(summary_labels)

        cell = ws.cell(row=row, column=datum_col)
        cell.value = "Wochenende/Feiertag"
        cell.font = Font(italic=True, color=FREE_DAY_COLOR)
        cell.alignment = Alignment(horizontal='left', vertical='center')

        cell = ws.cell(row=row+1, column=datum_col)
        cell.value = "diesen Tag mit SKUG auffüllen"
        cell.font = Font(italic=True, color=SKUG_COLOR)
        cell.alignment = Alignment(horizontal='left', vertical='center')

        cell = ws.cell(row=row+1, column=datum_col+5)
        cell.value = "weniger oder gleich als 8 Stunden von zu Hause abwesend"
        cell.font = Font(italic=True, color=UNTER_8H_COLOR)
        cell.alignment = Alignment(horizontal='left', vertical='center')

        cell = ws.cell(row=row+1, column=datum_col+11)
        cell.value = "An+Ab/>24"
        cell.font = Font(italic=True, color=AN_AB_COLOR)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        
        

        # Calculate and write summary values for each name
        for name_idx, name in enumerate(section_names):
            name_col = datum_col + 2 + (name_idx * 2)

            # Apply thick border to summary numbers
            set_create_border(
                min_row=summary_start_row,
                max_row=summary_start_row+len(summary_labels)-1,
                min_col=name_col,
                max_col=name_col+1,
                side_style=Side(style='thick'),
                ws=ws
            )
            
            # Get all entries for this name
            name_entries = [e for e in entries if e['name'] == name]
            
            person_data = person_lookup.get(name, {})
            worker_type = person_data.get('worker_type', 'Fest')
            kein_verpflegung = bool(person_data.get('kein_verpflegungsgeld', 0))
            keine_feiertag = bool(person_data.get('keine_feiertagssstunden', 0))
            weekly_hours = person_data.get('weekly_hours', 0.0)

            # Get SKUG settings for calculating Feiertag hours
            skug_settings = master_db.get_skug_settings()

            # Calculate totals
            gesamtstunden = sum(float(e.get('stunden', 0)) for e in name_entries)
            urlaubsstunden = sum(float(e.get('urlaub', 0) or 0) for e in name_entries)
            
            krankstunden = sum(float(e.get('krank', 0) or 0) for e in name_entries)

            # Only calculate SKUG total in winter months (December-March)
            is_winter = month in [12, 1, 2, 3]
            if is_winter:
                skug_total = sum(float(e.get('skug', 0) or 0) for e in name_entries)
            else:
                skug_total = 0

            # Calculate Feiertag hours
            feiertag = 0
            if not keine_feiertag:
                for day in range(1, num_days + 1):
                    # Check if it's a bank holiday (holiday but not weekend)
                    if is_holiday(year, month, day) and not is_weekend(year, month, day):
                        # Logic for Fest with weekly hours
                        if worker_type == WorkerTypes.Fest and gesamtstunden > 0:
                            gesamtstunden += weekly_hours / 5.0
                            feiertag += 1
                        elif worker_type == WorkerTypes.Fest:
                            feiertag += 1
                        else:
                            # Standard logic
                            # Check if there's no entry for this person on this day
                            has_entry = any(e['name'] == name and e['tag'] == day for e in entries)
                            if not has_entry:
                                # Add the target hours for this day based on SKUG settings
                                feiertag_hours = calculate_skug(year, month, day, 0, skug_settings)
                                feiertag += abs(feiertag_hours)

            # Summe and Mehr/Minder Calculation
            summe = 0
            mehr_minder = 0
            
            if worker_type == WorkerTypes.Fest and gesamtstunden > 0:
                # Special logic for Fest with weekly hours
                # Summe is fixed based on weekly hours
                summe = (weekly_hours * 52) / 12
                urlaubsstunden = get_days_of_urlaub(name, month, year, db)*weekly_hours/5.0
                krankstunden = get_days_of_krank(name, month, year, db)*weekly_hours/5.0
                gesamtstunden += (urlaubsstunden + krankstunden)
                # Mehr/Minder is difference
                mehr_minder = gesamtstunden - summe
            elif worker_type == WorkerTypes.Fest and gesamtstunden == 0:
                summe = ""
                mehr_minder = ""
            else:
                # Standard logic
                summe = gesamtstunden + feiertag + urlaubsstunden + krankstunden + skug_total
                mehr_minder = gesamtstunden - get_normal_hours_per_month(year, month, master_db)

            # Verpflegungsgeld
            if kein_verpflegung:
                v_zuschuss = 0.0
            else:
                v_zuschuss = get_verpflegungsgeld_for_name(name, month, year, master_db, db)
                
            fahrstunden = get_fahrstunden_for_name(name, month, year, master_db, db)


            summary_values = [
                gesamtstunden,
                feiertag,
                urlaubsstunden,
                krankstunden,
                skug_total,
                summe,
                mehr_minder,
                fahrstunden,
                v_zuschuss
            ]
            
            if worker_type == WorkerTypes.Gewerblich:
                create_zeitarbeiter_summary(ws, person_lookup, name, summary_values, summary_start_row, name_col)
            elif worker_type == WorkerTypes.Fest:
                create_fest_summary(ws, name, month, year, summary_values, summary_start_row, name_col, worker_type, master_db, db, weekly_hours)
                
            

        # Move to next section (add spacing)
        current_col += 2

    # Adjust column widths
    for col in range(1, current_col):
        ws.column_dimensions[get_column_letter(col)].width = 12

    # Save workbook
    max_row = 4+num_days + len(summary_labels)
    max_col = len(unique_names)*2 + num_sections*2
    addLattice(3, max_row, 1, max_col, ws)
    try:
        wb.save(filename)
        return True
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        return False
    
def create_zeitarbeiter_summary(ws: Workbook, person_lookup, name, summary_values: list, summary_start_row: int = None, name_col: int = None):
    for idx, value in enumerate(summary_values):
        row = summary_start_row + idx
        value_cell = ws.cell(row=row, column=name_col)
        value_cell.value = value if value != 0 else ""
        value_cell.alignment = Alignment(horizontal='center', vertical='center')
        value_cell.number_format = "0.00"

        if idx == 6:
            person_data = person_lookup.get(name, {})
            keine_feiertag = bool(person_data.get('kein_fzk', 0))
            if keine_feiertag:
                ws.merge_cells(start_row=row, start_column=name_col, end_row=row, end_column=name_col+1)
                value_cell = ws.cell(row=row, column=name_col)
                value_cell.value = "Kein FZK"
                value_cell.alignment = Alignment(horizontal='center', vertical='center')
                set_create_border(
                    min_row=row,
                    max_row=row,
                    min_col=name_col,
                    max_col=name_col+1,
                    side_style=Side(style='thick'),
                    ws=ws
                )
        
def create_fest_summary(ws: Workbook, name, month, year, summary_values: list, summary_start_row: int, name_col: int, worker_type: WorkerTypes, master_db: MasterDataDatabase, db:Database, weekly_hours: float = 0.0):
    stunden = summary_values[0]
    for idx, value in enumerate(summary_values):
        row = summary_start_row + idx
        
        # Standard value writing first (can be overwritten)
        value_cell = ws.cell(row=row, column=name_col)
        value_cell.value = value if value != 0 else ""
        value_cell.alignment = Alignment(horizontal='center', vertical='center')
        if idx == 0:
            value_cell.number_format = "0.00"
        if idx == 1:  # Feiertag
            value_cell = ws.cell(row=row, column=name_col+1)
            value_cell.value = "Tage"
        if idx == 2:  # Urlaubsstunden
            value_cell = ws.cell(row=row, column=name_col+1)
            value_cell.value = "Tage"
            
            value_cell = ws.cell(row=row, column=name_col)
            value_cell.value = get_days_of_urlaub(name, month, year, db)
        if idx == 3:  # Krankstunden
            value_cell = ws.cell(row=row, column=name_col+1)
            value_cell.value = "Tage"
            
            value_cell = ws.cell(row=row, column=name_col)
            value_cell.value = get_days_of_krank(name, month, year, db)
        if idx==5:
            value_cell.number_format = "0.00"
        if idx == 6:  # Mehr-/Minderstd
            # If weekly_hours > 0, we show the calculated value (already in summary_values[6])
            # If NOT weekly_hours > 0, we do the old merge thing
            if stunden <= 0:
                ws.merge_cells(start_row=row, start_column=name_col, end_row=row, end_column=name_col+1)
                value_cell = ws.cell(row=row, column=name_col)
                value_cell.value = worker_type
                value_cell.alignment = Alignment(horizontal='center', vertical='center')
                set_create_border(
                    min_row=row,
                    max_row=row,
                    min_col=name_col,
                    max_col=name_col+1,
                    side_style=Side(style='thick'),
                    ws=ws
                )
            else:
                # Just ensure formatting is correct for the value
                value_cell = ws.cell(row=row, column=name_col)
                value_cell.value = value
                value_cell.number_format = "0.00"
                value_cell.alignment = Alignment(horizontal='center', vertical='center')