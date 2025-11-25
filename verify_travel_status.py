import sqlite3
import os
from database import Database

DB_FILE = "test_travel_status.db"

def setup_db():
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)
    db = Database(DB_FILE)
    return db

def test_travel_status():
    print("Testing Travel Status...")
    db = setup_db()
    
    # 1. Test Basic Insert with Travel Status
    data = {
        "Jahr": 2023, "Monat": 11, "Tag": 10, "Name": "TestUser",
        "Wochentag": "Fr", "Stunden": 8.0, "Urlaub": "", "Krank": "",
        "unter_8h": False, "SKUG": "", "Baustelle": "123",
        "travel_status": "Anreise"
    }
    db.add_or_update_entry(data)
    
    entries = db.get_entries_by_month_and_name(2023, 11, "TestUser")
    assert len(entries) == 1
    assert entries[0]['travel_status'] == "Anreise"
    print("✓ Basic Insert Passed")
    
    # 2. Test Partial Update (Travel Status only, no hours change)
    # First, ensure hours are 8.0
    assert entries[0]['stunden'] == 8.0
    
    # Update only travel status
    update_data = {
        "Jahr": 2023, "Monat": 11, "Tag": 10, "Name": "TestUser",
        "travel_status": "Abreise"
    }
    db.add_or_update_entry(update_data)
    
    entries = db.get_entries_by_month_and_name(2023, 11, "TestUser")
    assert len(entries) == 1
    assert entries[0]['travel_status'] == "Abreise"
    assert entries[0]['stunden'] == 8.0 # Should remain unchanged
    print("✓ Partial Update Passed")
    
    # 3. Test New Entry with Empty Hours (should default to 0.0)
    new_data = {
        "Jahr": 2023, "Monat": 11, "Tag": 11, "Name": "TestUser",
        "Wochentag": "Sa", "Urlaub": "", "Krank": "",
        "unter_8h": False, "SKUG": "", "Baustelle": "123",
        "travel_status": "24h_away"
    }
    # Note: 'Stunden' is missing
    db.add_or_update_entry(new_data)
    
    entries = db.get_entries_by_month_and_name(2023, 11, "TestUser")
    # Should have 2 entries now
    entry_11 = next(e for e in entries if e['tag'] == 11)
    assert entry_11['travel_status'] == "24h_away"
    assert entry_11['stunden'] == 0.0
    print("✓ New Entry Default Hours Passed")

    # Cleanup
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)

if __name__ == "__main__":
    try:
        test_travel_status()
        print("\nAll tests passed!")
    except Exception as e:
        print(f"\nTest Failed: {e}")
        import traceback
        traceback.print_exc()
