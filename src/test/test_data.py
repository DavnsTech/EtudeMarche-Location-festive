import unittest
from src.data.competitor_data import CompetitorDataCollector

class TestCompetitorData(unittest.TestCase):
    def setUp(self):
        self.collector = CompetitorDataCollector()
        self.test_file = "test_competitor_research.xlsx"
        
    def tearDown(self):
        # Clean up test files
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_competitor_data_template_creation(self):
        # Test template creation
        file_path = self.collector.create_competitor_template()
        self.assertTrue(os.path.exists(file_path))
        self.assertEqual(file_path, self.test_file)
        
    def test_competitor_data_template_already_exists(self):
        # Create a test file first
        df = pd.DataFrame({
            "Competitor": ["Test"],
            "Website": [""],
            "Services": [""],
            "Pricing Range": [""]
        })
        df.to_excel(self.test_file, index=False)
        
        # Test when template already exists
        result = self.collector.create_competitor_template()
        self.assertEqual(result, self.test_file)
        self.assertTrue(os.path.exists(self.test_file))
        
    def test_competitor_data_template_creation_failure(self):
        # Test failure scenario (invalid directory)
        original_dir = self.collector.data_dir
        self.collector.data_dir = "invalid_dir"
        
        # This should fail
        result = self.collector.create_competitor_template()
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()
