"""Regression checks without desktop Excel, model downloads, or credentials."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# Only the Excel decorator is needed at import time. Workbook operations below
# use an in-memory adapter; actual model fitting uses sklearn.
with patch.dict(sys.modules, {'xlwings': types.SimpleNamespace(sub=lambda fn: fn),
                              'tabpfn': types.SimpleNamespace()}):
    forecast = load_module('forecast', ROOT / 'tabpfn_forecast.py')
divvy = load_module('divvy', ROOT / 'divvy_xlsx' / 'divvy_cleanup.py')


class Range:
    def __init__(self, value=None):
        self.value = value

    def expand(self, kind):
        return self


class Sheet:
    def __init__(self, name, values=None):
        self.name = name
        self.values = values or {}

    def range(self, address):
        return self.values.setdefault(address, Range())

    def clear(self):
        self.values.clear()


class Sheets(list):
    def __getitem__(self, key):
        if isinstance(key, str):
            return next(s for s in self if s.name == key)
        return super().__getitem__(key)

    def add(self, name, after=None):
        sheet = Sheet(name)
        self.append(sheet)
        return sheet


class Workbook:
    def __init__(self):
        self.sheets = Sheets([Sheet('Data', {
            (1, 1): Range(['Year', 'Value']),
            (2, 1): Range([[i, 2 * i + 1] for i in range(20)]),
        })])
        self.saved = False

    def save(self):
        self.saved = True


class ForecastTests(unittest.TestCase):
    def run_forecast(self, model_type):
        book = Workbook()
        result = forecast.run_forecast_configured(
            book, ['Data'], 1, 2, 'Value', 'all', [],
            primary_time_index='Year', model_type=model_type,
            n_gb_estimators=10, future_rows=3)
        rows = book.sheets['Forecast'].range('A2').value
        self.assertEqual(len(rows), 3)
        self.assertTrue(np.isfinite(np.asarray(rows, dtype=float)).all())
        self.assertTrue(all(row[-3] >= row[-2] and row[-3] <= row[-1] for row in rows))
        return result

    def test_gradient_boosting_selection_reaches_prediction(self):
        self.assertEqual(self.run_forecast('GradientBoosting')['model_type'], 'GradientBoosting')

    def test_missing_tabpfn_falls_back(self):
        with patch.object(forecast, 'HAS_TABPFN', False):
            self.assertEqual(self.run_forecast('TabPFN')['model_type'], 'GradientBoosting')

    def test_gated_tabpfn_discards_unfitted_model(self):
        class GatedModel:
            def __init__(self, **kwargs):
                pass

            def fit(self, X, y):
                raise RuntimeError('gated model: Authentication required')

        with patch.object(forecast, 'HAS_TABPFN', True), patch.object(forecast, 'TabPFNRegressor', GatedModel):
            self.assertEqual(self.run_forecast('TabPFN')['model_type'], 'GradientBoosting')

    def test_tabpfn_selection_reaches_prediction(self):
        # Exercise successful dispatch independently of external model access.
        class Model:
            def __init__(self, **kwargs):
                pass

            def fit(self, X, y):
                return self

            def predict(self, X):
                return X[:, 0] * 2 + 1

        with patch.object(forecast, 'HAS_TABPFN', True), patch.object(forecast, 'TabPFNRegressor', Model):
            self.assertEqual(self.run_forecast('TabPFN')['model_type'], 'TabPFN')

    def test_removed_models_rejected_before_reading_workbook(self):
        with self.assertRaisesRegex(ValueError, 'Choose model_type'):
            forecast.run_forecast_configured(None, [], 1, 2, 'Value', 'all', [], model_type='OLS')

    def test_shared_runner_writes_diagnostics_and_saves(self):
        book = Workbook()
        cfg = {key: default for key, flags, label, kind, default in forecast.SETTINGS_SCHEMA}
        cfg.update(data_sheets=['Data'], target_column='Value', model_type='GradientBoosting', future_rows=3)
        forecast._run_forecast(book, cfg)
        self.assertTrue(book.saved)
        self.assertIn('Diagnostics', [sheet.name for sheet in book.sheets])
        self.assertEqual(book.sheets['Diagnostics'].range('A1').value[1][1], 'GradientBoosting')

    def test_settings_schema_retains_output_configuration(self):
        defaults = {key: default for key, flags, label, kind, default in forecast.SETTINGS_SCHEMA}
        self.assertEqual(defaults['future_rows'], 10)
        self.assertEqual(defaults['output_sheet'], 'Forecast')
        self.assertNotIn('n_mc_samples', defaults)



class DivvyTests(unittest.TestCase):
    def test_cleanup_writes_sorted_copy_and_preserves_source(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'export.xlsx'
            data = {name: ['newer', 'older'] for name in divvy.KEPT_COLUMNS}
            data[divvy.SORT_COLUMN] = ['2026-10-26 12:00', '2026-10-25 09:00']
            data['Unused'] = [1, 2]
            pd.DataFrame(data).to_excel(source, index=False)
            before = source.read_bytes()
            with patch('builtins.input', return_value=str(source)):
                divvy.main()
            result = pd.read_excel(source.with_name('export_cleaned.xlsx'))
            self.assertEqual(list(result.columns), ['Date'] + divvy.KEPT_COLUMNS[1:])
            self.assertEqual(result['Date'].tolist(), ['2026-10-25', '2026-10-26'])
            self.assertEqual(result['First Name'].tolist(), ['older', 'newer'])
            self.assertEqual(source.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
