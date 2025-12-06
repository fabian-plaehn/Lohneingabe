
import unittest
from database import Database
from master_data import MasterDataDatabase
import utils

class TestWorkerOverrides(unittest.TestCase):
    def setUp(self):
        # Use existing DBs (or copies if safer, but for now we'll use existing as we'll cleanup)
        # Note: In a real scenario, mock DBs or temp files would be better.
        # Here we add temporary test data and remove it after
        self.db = Database()
        self.master_db = MasterDataDatabase()
        
        # Test Data
        self.test_worker_name = "TEST_WORKER_OVERRIDE"
        self.test_baustelle_nummer = "TEST_BST_999"
        
        # Ensure clean state
        self.cleanup()
        
        # 1. Create Worker
        self.worker_id = self.master_db.add_name(self.test_worker_name)
        
        # 2. Create Baustelle (Default values)
        # Verpflegung: 28.0, Fahrzeit: 1.0 (Roundtrip 2.0)
        self.baustelle_id = self.master_db.add_baustelle(self.test_baustelle_nummer, "Test Baustelle", 28.0, 1.0, 50.0)
        
        # 3. Add hours entry
        self.db.add_or_update_entry({
            'Jahr': 2025,
            'Monat': 1,
            'Tag': 1,
            'Name': self.test_worker_name,
            'Baustelle': f"{self.test_baustelle_nummer} - Test Baustelle",
            'Stunden': 8.0
        })

    def tearDown(self):
        self.cleanup()

    def cleanup(self):
        # Remove Override
        worker_id = self.master_db.get_worker_id_by_name(self.test_worker_name)
        baustelle_id = self.master_db.get_baustelle_id_by_nummer(self.test_baustelle_nummer)
        
        if worker_id and baustelle_id:
            # We don't have a direct delete override by worker/baustelle, but we can get it first
            override = self.master_db.get_override(worker_id, baustelle_id)
            if override:
                self.master_db.delete_override(override['id'])

        # Remove worker
        if worker_id:
             self.master_db.delete_name(worker_id)
             
        # Remove baustelle
        if baustelle_id:
            self.master_db.delete_baustelle(baustelle_id)
            
        # Remove entry
        entries = self.db.get_entries_by_month_and_name(2025, 1, self.test_worker_name)
        for e in entries:
            self.db.delete_entry(e['id'])

    def test_default_values(self):
        # Check without override
        # Fahrzeit = 1.0 * 2 = 2.0
        # Verpflegung = 28.0
        
        fahrstunden = utils.get_fahrstunden_for_name(self.test_worker_name, 1, 2025, self.master_db, self.db)
        verpflegung = utils.get_verpflegungsgeld_for_name(self.test_worker_name, 1, 2025, self.master_db, self.db)
        
        self.assertEqual(fahrstunden, 2.0)
        self.assertEqual(verpflegung, 28.0)

    def test_override_values(self):
        # Add Override
        # New Verpflegung: 50.0, New Fahrzeit: 2.0 (Roundtrip 4.0)
        self.master_db.add_override(self.worker_id, self.baustelle_id, verpflegungsgeld=50.0, fahrzeit=2.0)
        
        fahrstunden = utils.get_fahrstunden_for_name(self.test_worker_name, 1, 2025, self.master_db, self.db)
        verpflegung = utils.get_verpflegungsgeld_for_name(self.test_worker_name, 1, 2025, self.master_db, self.db)
        
        self.assertEqual(fahrstunden, 4.0)
        self.assertEqual(verpflegung, 50.0)

if __name__ == '__main__':
    unittest.main()
