import unittest
from src.reports.excel_generator import MarketStudyExcelReport

class TestExcelReports(unittest.TestCase):
    def setUp(self):
        self.report = MarketStudyExcelReport()
        self.test_dir = "test_reports"
        
    def tearDown(self):
        # Clean up test files
        if os.path.exists(self.test_dir):
            for file in os.listdir(self.test_dir):
                os.remove(os.path.join(self.test_dir, file))
            os.rmdir(self.test_dir)
    
    def test_excel_report_initialization(self):
        # Test initialization
        self.assertIsNotNone(self.report.wb)
        self.assertTrue(os.path.exists(self.report.reports_dir))
        self.assertTrue(self.report.output_filename.endswith(".xlsx"))
        
    def test_excel_report_directory_creation(self):
        # Test directory creation
        self.assertTrue(os.path.exists(self.report.reports_dir))
        
    def test_excel_report_output_file(self):
        # Verify output file path
        self.assertTrue(self.report.output_filename.endswith(".xlsx"))
        self.assertIn("reports", self.report.output_filename)

if __name__ == '__main__':
    unittest.main()
