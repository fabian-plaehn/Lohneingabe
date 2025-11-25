import sqlite3
import os
from database import Database

# Setup test DB
TEST_DB = "test_bulk_update.db"
if os.path.exists(TEST_DB):
    os.remove(TEST_DB)

db = Database(TEST_DB)

def test_logic():
    print("Starting verification...")
    
    # 1. Create initial entry: 7 hours -> unter_8h should be True
    print("\nTest 1: Initial entry 7h")
    data = {
        "Jahr": 2024, "Monat": 1, "Tag": 1, "Name": "TestUser",
        "Stunden": 7.0, "unter_8h": True, "Wochentag": "Mo"
    }
    db.add_or_update_entry(data)
    
    entry = db.get_entry(2024, 1, 1, "TestUser")
    print(f"Entry: {entry['stunden']}h, unter_8h={entry['unter_8h']}")
    assert entry['unter_8h'] == 1, "Should be unter_8h"

    # 2. Simulate bulk update: Add Breakfast (+0.25h)
    # Logic from GUI: 
    # current_stunden = existing (7.0)
    # verpflegungs_stunden = 7.0 + 0.25 = 7.25
    # unter_8h = 7.25 <= 8.0 -> True
    print("\nTest 2: Bulk update +Breakfast")
    
    # Simulate GUI logic
    existing_stunden = entry['stunden']
    verpflegungs_stunden = existing_stunden + 0.25
    unter_8h = verpflegungs_stunden <= 8.0
    
    update_data = {
        "Jahr": 2024, "Monat": 1, "Tag": 1, "Name": "TestUser",
        "unter_8h": unter_8h
        # Note: In real GUI, we don't send 'Stunden' if it wasn't changed in the input field
        # But we do send the updated unter_8h flag
    }
    db.add_or_update_entry(update_data)
    
    entry = db.get_entry(2024, 1, 1, "TestUser")
    print(f"Entry: {entry['stunden']}h, unter_8h={entry['unter_8h']}")
    assert entry['unter_8h'] == 1, "Should still be unter_8h (7.25h)"

    # 3. Simulate bulk update: Add Breakfast + Lunch (+0.75h)
    # Base is still 7.0h from DB (assuming we didn't change hours in DB, just flags)
    # Wait, the GUI logic uses the hours from DB.
    # If we didn't update hours in DB in step 2 (we didn't send "Stunden" key), then it's still 7.0.
    print("\nTest 3: Bulk update +Breakfast +Lunch")
    
    existing_stunden = entry['stunden'] # 7.0
    verpflegungs_stunden = existing_stunden + 0.25 + 0.5 # 7.75
    unter_8h = verpflegungs_stunden <= 8.0
    
    update_data = {
        "Jahr": 2024, "Monat": 1, "Tag": 1, "Name": "TestUser",
        "unter_8h": unter_8h
    }
    db.add_or_update_entry(update_data)
    
    entry = db.get_entry(2024, 1, 1, "TestUser")
    print(f"Entry: {entry['stunden']}h, unter_8h={entry['unter_8h']}")
    assert entry['unter_8h'] == 1, "Should still be unter_8h (7.75h)"

    # 4. Create entry with 7.5 hours
    print("\nTest 4: Initial entry 7.5h")
    data = {
        "Jahr": 2024, "Monat": 1, "Tag": 2, "Name": "TestUser",
        "Stunden": 7.5, "unter_8h": True, "Wochentag": "Di"
    }
    db.add_or_update_entry(data)
    
    entry = db.get_entry(2024, 1, 2, "TestUser")
    print(f"Entry: {entry['stunden']}h, unter_8h={entry['unter_8h']}")
    assert entry['unter_8h'] == 1, "Should be unter_8h"

    # 5. Bulk update 7.5h + Breakfast + Lunch -> 8.25h -> unter_8h = False
    print("\nTest 5: Bulk update 7.5h +Breakfast +Lunch")
    
    existing_stunden = entry['stunden'] # 7.5
    verpflegungs_stunden = existing_stunden + 0.25 + 0.5 # 8.25
    unter_8h = verpflegungs_stunden <= 8.0
    
    update_data = {
        "Jahr": 2024, "Monat": 1, "Tag": 2, "Name": "TestUser",
        "unter_8h": unter_8h
    }
    db.add_or_update_entry(update_data)
    
    entry = db.get_entry(2024, 1, 2, "TestUser")
    print(f"Entry: {entry['stunden']}h, unter_8h={entry['unter_8h']}")
    assert entry['unter_8h'] == 0, "Should NOT be unter_8h (8.25h)"

    print("\nVerification successful!")

if __name__ == "__main__":
    try:
        test_logic()
    finally:
        if os.path.exists(TEST_DB):
            os.remove(TEST_DB)
