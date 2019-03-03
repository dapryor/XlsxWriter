###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2019, John McNamara, jmcnamara@cpan.org
#

import unittest
from datetime import datetime

from ...workbook import Workbook


class TestCalculateSpans(unittest.TestCase):
    """
    Test the _calculate_spans Worksheet method for different cell ranges.

    """

    def setUp(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.add_worksheet()

        self.width_adjustment = 0.17

    def test_autofit_empty_sheet(self):
        """Test autofit on an empty sheet works"""
        exp = 0
        got = self.worksheet.autofit_columns('A:A')
        self.assertEqual(got, exp)

    def test_constant_memory_mode(self):
        """Test autofit in constant memory mode"""
        self.worksheet.constant_memory = 1

        exp = -1
        got = self.worksheet.autofit_columns('A:A')

        self.assertEqual(got, exp)

    def test_max_width_lower_bound(self):
        """Test autofit with max_width set to lower bound"""
        self.worksheet.write_number('A1', 100)
        exp = -2
        got = self.worksheet.autofit_columns('A:A', 0)

        self.assertEqual(got, exp)

    def test_max_width_above_bound(self):
        """Test autofit with max_width set above bound"""
        self.worksheet.write_number('A1', 100)
        exp = -2
        got = self.worksheet.autofit_columns('A:A', 250.1)

        self.assertEqual(got, exp)

    def test_max_width_below_bound(self):
        """Test autofit with max_width set below bound"""
        self.worksheet.write_number('A1', 100)
        exp = -2
        got = self.worksheet.autofit_columns('A:A', -1)

        self.assertEqual(got, exp)

    def test_max_width_upper_bound(self):
        """Test autofit with max_width set to upper bound"""
        self.worksheet.write_number('A1', 100)
        exp = 0
        got = self.worksheet.autofit_columns('A:A', 250)

        self.assertEqual(got, exp)

    def test_columns_out_of_bounds(self):
        """Test autofit with columns out of bounds"""
        self.worksheet.write('A1', '100')

        exp = -3
        got = self.worksheet.autofit_columns(1, self.worksheet.xls_colmax + 1)

        self.assertEqual(got, exp)

    def test_autofit_string(self):
        """Test autofit with string"""
        self.worksheet.write('A1', '100')
        self.worksheet.autofit_columns('A:A')

        exp = 3
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertEqual(got, exp)

    def test_autofit_too_long_string(self):
        """Test autofit with string that is longer than the max_width"""
        self.worksheet.write('A1', '_' * 101)
        self.worksheet.autofit_columns('A:A', 100)

        exp = 100
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertEqual(got, exp)

    def test_autofit_number(self):
        """Test autofit with number"""
        self.worksheet.write('A1', 1000)
        self.worksheet.autofit_columns('A:A')

        exp = 4
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertAlmostEqual(got, exp)

    def test_autofit_long_number_1(self):
        """Test autofit with scientific notation number with leading zeros"""
        self.worksheet.write('A1', 1e+12)
        self.worksheet.autofit_columns('A:A')

        exp = 5
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertAlmostEqual(got, exp)

    def test_autofit_long_number_2(self):
        """Test autofit with scientific notation number"""
        self.worksheet.write('A1', 1.234567e+12)
        self.worksheet.autofit_columns('A:A')

        exp = 11
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertAlmostEqual(got, exp)

    def test_autofit_boolean_1(self):
        """Test autofit with boolean True"""
        self.worksheet.write('A1', True)
        self.worksheet.autofit_columns('A:A')

        exp = 4
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertAlmostEqual(got, exp)

    def test_autofit_boolean_2(self):
        """Test autofit with boolean False"""
        self.worksheet.write('A1', False)
        self.worksheet.autofit_columns('A:A')

        exp = 5
        got = self.worksheet.col_sizes[0] - self.width_adjustment

        self.assertAlmostEqual(got, exp)

    def test_autofit_multicolumn(self):
        """Test autofit across mutli column range"""
        self.worksheet.write('A1', 'False')
        self.worksheet.write('B1', "True")
        self.worksheet.autofit_columns('A:C')

        exp = (5, 4, -self.width_adjustment)
        col_a = self.worksheet.col_sizes.get(0, 0) - self.width_adjustment
        col_b = self.worksheet.col_sizes.get(1, 0) - self.width_adjustment
        col_c = self.worksheet.col_sizes.get(2, 0) - self.width_adjustment
        got = (col_a, col_b, col_c)

        self.assertAlmostEqual(got, exp)

    def test_autofit_datetime(self):
        """Test autofit with a datetime"""
        data_format = self.workbook.add_format({'num_format': 'mmm d yyyy hh:mm AM/PM'})
        self.worksheet.write('A1', datetime.now(), data_format)
        self.worksheet.autofit_columns('A:A')

        exp = 22
        col_a = self.worksheet.col_sizes.get(0, 0) - self.width_adjustment
        got = col_a
        self.assertAlmostEqual(got, exp)

    def test_autofit_formula(self):
        """Test autofit with a formula"""
        self.worksheet.write('A1', '=SUM(1, 2, 3)')
        self.worksheet.autofit_columns('A:A')

        exp = -self.width_adjustment
        col_a = self.worksheet.col_sizes.get(0, 0) - self.width_adjustment
        got = col_a
        self.assertAlmostEqual(got, exp)
