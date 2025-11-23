"""
Script to reset the master_data.db database.
This will delete the existing database and create a new one with the updated schema.
"""
import os

def reset_database():
    """Delete the master_data.db file if it exists."""
    db_file = "master_data.db"

    if os.path.exists(db_file):
        try:
            os.remove(db_file)
            print(f"✓ Database '{db_file}' has been deleted.")
            print("The database will be recreated with the new schema when you start the application.")
        except Exception as e:
            print(f"✗ Error deleting database: {e}")
    else:
        print(f"Database '{db_file}' does not exist. Nothing to delete.")

if __name__ == "__main__":
    confirm = input("This will DELETE the master_data.db database. Are you sure? (yes/no): ")
    if confirm.lower() in ['yes', 'y']:
        reset_database()
    else:
        print("Operation cancelled.")
