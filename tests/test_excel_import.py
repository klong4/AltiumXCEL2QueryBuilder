import unittest
from src.services.excel_importer import ExcelImporter

class TestExcelImporter(unittest.TestCase):

    def setUp(self):
        self.importer = ExcelImporter()

    def test_import_valid_excel_file(self):
        # Assuming we have a valid Excel file for testing
        file_path = 'tests/test_files/valid_data.xlsx'
        data = self.importer.import_excel(file_path)
        self.assertIsNotNone(data)
        self.assertIsInstance(data, list)  # Assuming the data is returned as a list

    def test_import_invalid_excel_file(self):
        file_path = 'tests/test_files/invalid_data.xlsx'
        with self.assertRaises(ValueError):
            self.importer.import_excel(file_path)

    def test_import_empty_excel_file(self):
        file_path = 'tests/test_files/empty_data.xlsx'
        data = self.importer.import_excel(file_path)
        self.assertEqual(data, [])  # Assuming empty file returns an empty list

if __name__ == '__main__':
    unittest.main()