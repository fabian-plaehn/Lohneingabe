import pandas as pd
from datetime import datetime

class DataHandler:
    def __init__(self, filename="stundenliste.xlsx"):
        self.filename = filename
        self.entries = []
    
    def add_entry(self, data):
        """Add a new entry to the list."""
        self.entries.append(data)
    
    def save_to_excel(self):
        """Export all entries to Excel."""
        df = pd.DataFrame(self.entries)
        df.to_excel(self.filename, index=False)
    
    def get_all_entries(self):
        """Return all entries."""
        return self.entries