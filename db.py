from pymongo import MongoClient
from bson import ObjectId
import json

class AccessibilityDB:
    def __init__(self):
        try:
            self.client = MongoClient('mongodb://localhost:27017/',
                                    serverSelectionTimeoutMS=5000)
            self.client.server_info()
            self.db = self.client['accessibility_tests']
            
            # Separate collections for test runs and page results
            self.test_runs = self.db['test_runs']
            self.page_results = self.db['page_results']
            
            # Create indexes
            self.page_results.create_index([('url', 1), ('test_run_id', 1)])
            self.page_results.create_index('timestamp')
            self.test_runs.create_index('timestamp')
        except Exception as e:
            print(f"Failed to connect to MongoDB: {e}")
            raise

    def get_latest_test_run(self):
        """Get the most recent test run"""
        return self.test_runs.find_one(
            sort=[('timestamp_start', -1)]
        )
    
    def get_all_test_runs(self):
        """Get all test runs"""
        return list(self.test_runs.find(
           sort=[('timestamp_start', -1)]
        ))

    def get_page_results(self, test_run_id):
        """Get all page results for a specific test run"""
        try:
            return list(self.page_results.find(
                {'test_run_id': str(test_run_id)},
                {'_id': 0}
            ))
        except Exception as e:
            print(f"Error getting page results: {e}")
            return []

    def __del__(self):
        if hasattr(self, 'client'):
            self.client.close()