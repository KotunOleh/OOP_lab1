import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from unittest.mock import Mock, MagicMock
from back.calculator import FormulaCalculator

class TestFormulaCalculator(unittest.TestCase):

    def setUp(self):
        self.calculator = FormulaCalculator()

        self.mock_table = MagicMock()

        self.cell_data = {
            (0, 0): "10",   # A1
            (0, 1): "20",   # B1
            (1, 0): "30",   # A2
            (1, 1): "40",   # B2
            (2, 0): "Text", # A3
        }

        def get_mock_item(row, col):
            value = self.cell_data.get((row, col), "")
            mock_item = Mock()
            mock_item.text.return_value = str(value)
            return mock_item
        self.mock_table.item.side_effect = get_mock_item

        self.mock_table.rowCount.return_value = 5
        self.mock_table.columnCount.return_value = 5

    #  SUM
    def test_sum_range_function(self):
        formula = "=SUM(A1:B2)"
        
        expected_result = "100.0" 
        
        result = self.calculator.parse_and_calculate(formula, self.mock_table)
        
        self.assertEqual(result, expected_result)

    #dependencies
    def test_arithmetic_with_cells(self):
        formula = "=A1 + B1 * 2"
        
        expected_result = "50.0"
        
        result = self.calculator.parse_and_calculate(formula, self.mock_table)
        
        self.assertEqual(result, expected_result)

    # MAX 
    def test_max_function_with_args(self):
        formula = "=MAX(A1, B1, 5)"
        
        expected_result = "20.0"
        
        result = self.calculator.parse_and_calculate(formula, self.mock_table)
        
        self.assertEqual(result, expected_result)

if __name__ == '__main__':
    unittest.main()