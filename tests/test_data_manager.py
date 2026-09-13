import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch


APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from models.data_manager import DataManager


class DataManagerRegressionTests(unittest.TestCase):
    def make_manager(self, root):
        manager = DataManager()
        manager.JSON_DIR = root
        manager.SETTINGS_FILE = os.path.join(root, "settings.json")
        manager.DATA_ROOT_DIR = os.path.join(root, "data")
        os.makedirs(manager.DATA_ROOT_DIR, exist_ok=True)
        return manager

    def test_legacy_root_files_are_ignored(self):
        with tempfile.TemporaryDirectory() as root:
            manager = self.make_manager(root)
            for name in ("data.json", "data_1.json"):
                with open(os.path.join(root, name), "w", encoding="utf-8") as file:
                    json.dump({"data": {"2026-7-1-1": [["legacy", "999", ""]]}}, file)

            manager.load_data()

            self.assertEqual(manager.data, {})

    def test_csv_export_and_append_import_respect_period(self):
        with tempfile.TemporaryDirectory() as root:
            manager = self.make_manager(root)
            manager.data = {
                "2026-6-1-1": [["outside", "50", ""]],
                "2026-7-2-2": [["shop", "1200", "memo"]],
            }
            csv_path = os.path.join(root, "export.csv")

            count = manager.export_csv(csv_path, (2026, 7), (2026, 7), ["日付", "A", "B"])
            self.assertEqual(count, 1)

            imported = self.make_manager(os.path.join(root, "imported"))
            imported.data = {"2026-7-2-2": [["existing", "300", ""]]}
            count = imported.import_csv(csv_path, (2026, 7), (2026, 7), "append", 3)

            self.assertEqual(count, 1)
            self.assertEqual(len(imported.data["2026-7-2-2"]), 2)
            self.assertEqual(imported.data["2026-7-2-2"][1][0], "shop")

    def test_csv_replace_clears_entire_selected_period(self):
        with tempfile.TemporaryDirectory() as root:
            manager = self.make_manager(root)
            manager.data = {
                "2026-7-1-1": [["old", "100", ""]],
                "2026-8-1-1": [["keep", "200", ""]],
            }
            csv_path = os.path.join(root, "empty.csv")
            with open(csv_path, "w", encoding="utf-8-sig", newline="") as file:
                file.write("年,月,日,列番号,項目,支払先,金額,メモ\n")

            count = manager.import_csv(csv_path, (2026, 7), (2026, 7), "replace", 3)

            self.assertEqual(count, 0)
            self.assertNotIn("2026-7-1-1", manager.data)
            self.assertIn("2026-8-1-1", manager.data)
            with open(manager._get_data_file_path(2026, 7), encoding="utf-8") as file:
                self.assertEqual(json.load(file)["data"], {})

    def test_csv_replace_rolls_memory_back_when_save_fails(self):
        with tempfile.TemporaryDirectory() as root:
            manager = self.make_manager(root)
            manager.data = {'2026-7-1-1': [['old', '100', '']]}
            csv_path = os.path.join(root, 'replace.csv')
            with open(csv_path, 'w', encoding='utf-8-sig', newline='') as file:
                file.write('年,月,日,列番号,項目,支払先,金額,メモ\n2026,7,1,1,A,new,200,\n')
            with patch.object(manager, '_save_complete_month', side_effect=OSError('disk')):
                with self.assertRaises(OSError):
                    manager.import_csv(csv_path, (2026, 7), (2026, 7), 'replace', 3)
            self.assertEqual(manager.data, {'2026-7-1-1': [['old', '100', '']]})

    def test_remove_custom_column_updates_data_and_settings_together(self):
        with tempfile.TemporaryDirectory() as root:
            manager = self.make_manager(root)
            manager.custom_columns = ['A', 'B']
            manager.data = {'2026-7-1-12': [['A', '100', '']],
                            '2026-7-1-13': [['B', '200', '']]}
            manager.remove_custom_column('A')
            self.assertEqual(manager.custom_columns, ['B'])
            self.assertEqual(manager.data, {'2026-7-1-12': [['B', '200', '']]})

if __name__ == "__main__":
    unittest.main()
