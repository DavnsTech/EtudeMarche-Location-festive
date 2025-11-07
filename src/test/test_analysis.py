import unittest
from src.analysis.competitor_analysis import CompetitorAnalysis

class TestCompetitorAnalysis(unittest.TestCase):
    def setUp(self):
        # Create a test data file
        self.test_file = "test_data.xlsx"
        self.test_data = {
            "Competitor": ["Test Competitor"],
            "Services": ["Test Services"],
            "Pricing Range": ["Test Range"],
            "Market Share (%)": [50]
        }
        
    def tearDown(self):
        # Clean up test files
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_competitor_analysis_initialization(self):
        # Test initialization with valid file
        analysis = CompetitorAnalysis(self.test_file)
        self.assertIsNotNone(analysis)
        self.assertIsNone(analysis.df)  # Initial DataFrame should be None
        
    def test_competitor_analysis_load_data(self):
        # Create a test Excel file
        df = pd.DataFrame(self.test_data)
        df.to_excel(self.test_file, index=False)
        
        # Test loading data
        analysis = CompetitorAnalysis(self.test_file)
        analysis._load_data()
        self.assertIsNotNone(analysis.df)
        self.assertEqual(len(analysis.df), 1)
        self.assertEqual(analysis.df.iloc[0]["Competitor"], "Test Competitor")
        
    def test_competitor_analysis_load_data_missing_file(self):
        # Test loading data when file doesn't exist
        analysis = CompetitorAnalysis("nonexistent_file.xlsx")
        analysis._load_data()
        self.assertIsNone(analysis.df)

if __name__ == '__main__':
    unittest.main()
