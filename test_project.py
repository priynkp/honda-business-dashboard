"""Regression tests. Run: python -m unittest -v test_project"""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from analysis import analyze, export_analysis, transaction_table
from app import WorkbookStore, normalize_record
from generate_demo_data import generate_records, write_workbooks


class DashboardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.samples = tempfile.TemporaryDirectory()
        cls.base = Path(cls.samples.name)
        write_workbooks(cls.base)
        cls.store = WorkbookStore(cls.base)

    @classmethod
    def tearDownClass(cls):
        cls.samples.cleanup()

    def test_generated_counts_and_reproducibility(self):
        one, two = generate_records(), generate_records()
        self.assertEqual(one, two)
        self.assertEqual({key: len(rows) for key, rows in one.items()},
                         {'Rentals': 600, 'Repairs': 600, 'Sales': 300})
        self.assertNotEqual(one, generate_records(seed=42))

    def test_generator_protects_existing_files(self):
        with self.assertRaises(FileExistsError):
            write_workbooks(self.base)

    def test_validation(self):
        good = ['Aarav Sharma', 'City', '100.25', '2025-02-28']
        self.assertEqual(normalize_record('Rentals', good)[2], 100.25)
        for index, invalid in [(0, ''), (0, '=1+1'), (2, '-1'), (2, 'nan'),
                               (2, 'inf'), (2, '1.001'), (2, '1,000'), (3, '2025-02-30')]:
            row = good.copy()
            row[index] = invalid
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                normalize_record('Rentals', row)
        with self.assertRaises(ValueError):
            normalize_record('Rentals', ['too', 'few'])

    def test_literal_case_insensitive_search(self):
        frame = self.store.filtered('Rentals', 'cItY')
        self.assertGreater(len(frame), 0)
        self.assertTrue(frame['Car Model'].eq('City').all())
        self.assertTrue(self.store.filtered('Rentals', '.*').empty)
        self.assertEqual(len(self.store.filtered('Rentals', '  ')), 600)
        self.assertEqual(len(self.store.frames['Rentals']), 600)

    def test_save_reload_preserves_seed(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            write_workbooks(base)
            seed = (base / 'Rentals.xlsx').read_bytes()
            store = WorkbookStore(base)
            store.add('Rentals', ['Nisha Gupta', 'City', '100.25', '2026-01-01'])
            self.assertEqual(len(WorkbookStore(base).frames['Rentals']), 601)
            self.assertEqual((base / 'Rentals.xlsx').read_bytes(), seed)
            self.assertFalse((base / 'data' / '.Rentals.tmp.xlsx').exists())

    def test_failed_write_does_not_publish(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            write_workbooks(base)
            store = WorkbookStore(base)
            with patch('app.os.replace', side_effect=PermissionError('Locked file')):
                with self.assertRaises(PermissionError):
                    store.add('Rentals', ['Nisha Gupta', 'City', '100', '2025-01-01'])
            self.assertEqual(len(store.frames['Rentals']), 600)
            self.assertFalse((base / 'data' / 'Rentals.xlsx').exists())
            self.assertFalse((base / 'data' / '.Rentals.tmp.xlsx').exists())

    def test_schema_error_is_not_silent(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            write_workbooks(base)
            pd.DataFrame({'wrong': [1]}).to_excel(base / 'Rentals.xlsx', index=False)
            with self.assertRaisesRegex(ValueError, 'expected columns'):
                WorkbookStore(base)

    def test_tidy_table_and_totals(self):
        tidy = transaction_table(self.store.frames)
        result = analyze(self.store.frames)
        self.assertEqual(len(tidy), 1500)
        self.assertEqual(list(tidy.columns), ['date', 'amount', 'model', 'customer', 'category'])
        self.assertEqual(result['kpis'].loc['Sales', 'recorded_amount'], 330781500)
        for category in self.store.frames:
            self.assertAlmostEqual(result['monthly'][category].sum(),
                                   result['kpis'].loc[category, 'recorded_amount'])
        self.assertEqual(len(result['monthly']), 24)

    def test_year_filters_partition_totals(self):
        all_years = analyze(self.store.frames)
        first = analyze(self.store.frames, 2024)
        second = analyze(self.store.frames, 2025)
        self.assertEqual(len(first['monthly']), 12)
        self.assertEqual(first['monthly'].index.year.unique().tolist(), [2024])
        self.assertEqual(int(first['kpis']['records'].sum() + second['kpis']['records'].sum()), 1500)
        pd.testing.assert_series_equal(first['kpis']['recorded_amount'] + second['kpis']['recorded_amount'],
                                       all_years['kpis']['recorded_amount'])

    def test_empty_year_is_zero_filled(self):
        result = analyze(self.store.frames, 2099)
        self.assertEqual(len(result['monthly']), 12)
        self.assertEqual(result['monthly'].sum().sum(), 0)
        self.assertEqual(result['kpis'].sum().sum(), 0)
        self.assertTrue(result['top_customers'].empty)

    def test_rankings_and_repair_average(self):
        result = analyze(self.store.frames)
        self.assertEqual(result['model_sales'].index[0], 'City')
        self.assertEqual(result['model_sales']['sales'].sum(), 300)
        self.assertEqual(len(result['top_customers']), 10)
        self.assertTrue(result['top_customers']['recorded_amount'].is_monotonic_decreasing)
        self.assertAlmostEqual(result['kpis'].loc['Repairs', 'average_amount'],
                               self.store.frames['Repairs']['COST'].mean())
        self.assertEqual(result['repair_summary']['jobs'].sum(), 600)

    def test_missing_months_are_zero_filled(self):
        frames = {category: frame.iloc[:0].copy() for category, frame in self.store.frames.items()}
        frames['Rentals'] = pd.DataFrame([
            ['Riya Verma', 'City', 100, '2025-01-10'],
            ['Karan Patel', 'Jazz', 200, '2025-03-10'],
        ], columns=self.store.frames['Rentals'].columns)
        monthly = analyze(frames)['monthly']
        self.assertEqual(len(monthly), 3)
        self.assertEqual(monthly.loc['2025-02-01', 'Rentals'], 0)
        self.assertEqual(monthly['Rentals'].sum(), 300)

    def test_all_empty_categories(self):
        frames = {category: frame.iloc[:0].copy() for category, frame in self.store.frames.items()}
        result = analyze(frames)
        self.assertEqual(result['kpis'].sum().sum(), 0)
        self.assertTrue(result['monthly'].empty)

    def test_csv_export(self):
        with tempfile.TemporaryDirectory() as directory:
            paths = export_analysis(analyze(self.store.frames, 2025), directory)
            self.assertEqual(len(paths), 5)
            self.assertTrue(all(path.exists() and path.stat().st_size for path in paths))
            monthly = pd.read_csv(Path(directory) / 'monthly.csv')
            self.assertEqual(len(monthly), 12)
            self.assertIn('month', monthly.columns)


if __name__ == '__main__':
    unittest.main()
