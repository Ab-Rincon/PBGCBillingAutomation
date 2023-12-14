import pandas as pd
import unittest
from utils.clean_data import CleanData, calculate_total_time_difference


class TestCleanDataClass(unittest.TestCase):

    def setUp(self):
        # Set up a mock DataFrame
        data = {'Date': ['01/01/2023'], 'Time Period': ['time in: 09:00 - time out: 17:00']}
        self.df = pd.DataFrame(data)
        self.clean_data_instance = CleanData(self.df)

    def test_calculate_total_time_difference(self):
        in_times = ['09:00', '13:00']
        out_times = ['12:00', '17:00']
        expected_total_hours = 3+4  # Expected total hours
        self.assertEqual(calculate_total_time_difference(in_times, out_times), expected_total_hours)


if __name__ == '__main__':
    unittest.main()
