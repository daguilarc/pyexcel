#!/usr/bin/env python3
"""
TabPFN Forecast

Two entry points, both routed through run_forecast_configured() via _run_forecast():
- run_forecast() - macro entry point, invoked from VBA via RunPython
- main()         - CLI entry point (python tabpfn_forecast.py <workbook.xlsx> [flags])

See README.md for configuration details (PyexcelSettings.xlsx, 'Forecast' sheet).
"""

import os
import xlwings as xw  # type: ignore
import pandas as pd
import numpy as np
import sys
import argparse
import datetime as dt
import numbers
import re
from sklearn.ensemble import GradientBoostingRegressor

# Model licenses depend on the checkpoint; recent weights such as TabPFN-2.5
# use non-commercial, non-production terms. Check the selected weights' license.
# Download access does not grant commercial usage rights.
# https://github.com/PriorLabs/TabPFN#license
# https://huggingface.co/Prior-Labs/tabpfn_2_5/blob/main/LICENSE
# Prior Labs' hosted API has separate terms: https://priorlabs.ai/pricing
# Optional imports
try:
    from tabpfn import TabPFNRegressor # type: ignore
    HAS_TABPFN = True
except ImportError:
    HAS_TABPFN = False
    TabPFNRegressor = None  # type: ignore



# =============================================================================
# GLOBAL HELPER FUNCTIONS
# =============================================================================

SETTINGS_FILENAME = 'PyexcelSettings.xlsx'
SETTINGS_SHEET_NAME = 'Forecast'

# Single source of truth for every settings key: (key, cli flags, settings-sheet
# label, value kind, built-in default). Drives read_settings, create_default_settings,
# the macro entry's cfg dict, and the CLI's argparse setup + precedence resolution.
SETTINGS_SCHEMA = [
    ('data_sheets', ('--data-sheets',), 'Data Sheets (comma-separated)', 'list', []),
    ('header_row', ('--header-row',), 'Header Row (1-based)', 'int', 1),
    ('data_start_row', ('--data-start-row',), 'Data Start Row (1-based)', 'int', 2),
    ('target_column', ('--target-column',), 'Target Column', 'str', None),
    ('feature_columns', ('--feature-columns',), 'Feature Columns (or "all")', 'str', 'all'),
    ('categorical_columns', ('--categorical-columns',), 'Categorical Columns', 'list', []),
    ('primary_time_index', ('--primary-time-index',), 'Primary Time Index', 'str', None),
    ('secondary_time_index', ('--secondary-time-index',), 'Secondary Time Index', 'str', None),
    ('model_type', ('--model-type',), 'Model Type', 'str', 'TabPFN'),
    ('n_ensemble', ('--ensemble-size',), 'Ensemble Size (TabPFN)', 'int', 50),
    ('n_gb_estimators', ('--gb-estimators',), 'GB Estimators (GradientBoosting)', 'int', 100),
    ('future_rows', ('--future-rows',), 'Future Rows', 'int', 10),
    ('output_sheet', ('--output-sheet',), 'Output Sheet', 'str', 'Forecast'),
]

# CLI help text per key (falls back to a generated string from the label if absent).
_SETTINGS_HELP = {
    'data_sheets': 'Comma-separated sheet names to read',
    'header_row': 'Row containing column headers (1-based, default: 1)',
    'data_start_row': 'Row where data starts (1-based, default: 2)',
    'target_column': 'Name of target variable',
    'feature_columns': 'Historical numeric covariates: all, none, or comma-separated names (default: all)',
    'categorical_columns': 'Comma-separated identity columns defining independent series',
    'primary_time_index': 'Primary time column name (e.g., "Year")',
    'secondary_time_index': 'Secondary time column name (e.g., "Date", "Quarter")',
    'model_type': 'Model to use: TabPFN, GradientBoosting (default: TabPFN)',
    'n_ensemble': 'TabPFN ensemble size (default: 50)',
    'n_gb_estimators': 'GradientBoosting n_estimators (default: 100)',
    'future_rows': 'Number of subsequent observations per series (default: 10)',
    'output_sheet': 'Output sheet name (default: Forecast)',
}


def _get_or_create_sheet(wb, sheet_name: str, wb_sheet_names: set):
    """Return sheet_name, adding it after the last sheet if absent, else clearing and reusing it."""
    if sheet_name not in wb_sheet_names:
        return wb.sheets.add(sheet_name, after=wb.sheets[-1])
    sheet = wb.sheets[sheet_name]
    sheet.clear()
    return sheet


# =============================================================================
# SETTINGS
# =============================================================================

def get_settings_path() -> str:
    """Return the path to PyexcelSettings.xlsx beside this script."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), SETTINGS_FILENAME)


def settings_sheet_exists(path: str) -> bool:
    """Return True if `path` exists and contains the SETTINGS_SHEET_NAME sheet."""
    if not os.path.exists(path):
        return False
    settings_wb = xw.Book(path)
    try:
        return SETTINGS_SHEET_NAME in {s.name for s in settings_wb.sheets}
    finally:
        settings_wb.close()


def read_settings(path: str) -> dict:
    """
    Read the 'Forecast' settings sheet (label/value pairs in columns A/B).
    Returns only keys that are present and non-blank; absent or blank keys are
    omitted so callers fall through to built-in defaults.
    """
    settings_wb = xw.Book(path)
    try:
        sheet = settings_wb.sheets[SETTINGS_SHEET_NAME]
        used = sheet.used_range
        rows = used.value if used is not None else []
        if rows and not isinstance(rows[0], (list, tuple)):
            rows = [rows]
        labels = {}
        for row in rows or []:
            if not row or row[0] is None:
                continue
            label = str(row[0]).strip()
            labels[label] = row[1] if len(row) > 1 else None
    finally:
        settings_wb.close()

    def present(label):
        if label not in labels:
            return False
        value = labels[label]
        return not (value is None or (isinstance(value, str) and not value.strip()))

    settings: dict = {}
    for key, _flags, label, kind, _default in SETTINGS_SCHEMA:
        if not present(label):
            continue
        raw_value = labels[label]
        if kind == 'list':
            settings[key] = [s.strip() for s in str(raw_value).split(',') if s.strip()]
        elif kind == 'int':
            if isinstance(raw_value, (bool, np.bool_)):
                raise ValueError(f'{label} must be a positive integer, not a boolean')
            settings[key] = float(raw_value)
        elif kind == 'float':
            settings[key] = float(raw_value)
        else:  # 'str'
            settings[key] = str(raw_value).strip()
    return settings


def create_default_settings(path: str) -> None:
    """
    Create PyexcelSettings.xlsx with a default 'Forecast' settings sheet.
    If the workbook already exists, open it and add/replace the 'Forecast' sheet rather than
    overwriting the file.
    """
    default_rows = [
        [label, '' if (kind == 'list' or default is None) else default]
        for _key, _flags, label, kind, default in SETTINGS_SCHEMA
    ]

    file_exists = os.path.exists(path)
    settings_wb = xw.Book(path) if file_exists else xw.Book()
    try:
        if SETTINGS_SHEET_NAME in {s.name for s in settings_wb.sheets}:
            sheet = settings_wb.sheets[SETTINGS_SHEET_NAME]
            sheet.clear()
        elif file_exists:
            sheet = settings_wb.sheets.add(SETTINGS_SHEET_NAME, after=settings_wb.sheets[-1])
        else:
            sheet = settings_wb.sheets[0]
            sheet.name = SETTINGS_SHEET_NAME
        sheet.range('A1').value = default_rows
        sheet.range('A:A').column_width = 34
        sheet.range('B:B').column_width = 26
        settings_wb.save(path)
    finally:
        settings_wb.close()


# =============================================================================
# SHARED RUN FUNCTION (both entry points call this)
# =============================================================================


def _check_settings_boundary(data_path, settings_path):
    """Protect the open data workbook from settings helpers that close books."""
    same = os.path.realpath(data_path) == os.path.realpath(settings_path)
    if os.path.exists(data_path) and os.path.exists(settings_path):
        same = same or os.path.samefile(data_path, settings_path)
    if same:
        raise ValueError('Use a separate settings workbook; data and settings paths are identical')


def _read_frame(sheet, header_row, data_start_row):
    """Read explicit rectangular extents, preserving singleton dimensions."""
    last_col = sheet.used_range.last_cell.column
    last_row = sheet.used_range.last_cell.row
    headers = sheet.range((header_row, 1), (header_row, last_col)).options(ndim=2).value[0]
    if any(not isinstance(h, str) or not h.strip() for h in headers):
        raise ValueError(f'{sheet.name}: headers must be nonblank text')
    headers = [h.strip() for h in headers]
    if len(set(headers)) != len(headers):
        raise ValueError(f'{sheet.name}: duplicate headers')
    values = [] if last_row < data_start_row else sheet.range(
        (data_start_row, 1), (last_row, last_col)).options(ndim=2).value
    rows, locations = [], []
    for row_number, row in enumerate(values, data_start_row):
        if all(v is None or isinstance(v, str) and not v.strip() for v in row):
            continue
        rows.append(row)
        locations.append((sheet.name, row_number))
    frame = pd.DataFrame(rows, columns=headers, dtype=object)
    frame.attrs['provenance'] = locations
    return frame


def _prepare_series(history, cfg):
    """Normalize required cells while retaining cell labels and reader provenance."""
    target, primary, secondary = (cfg[k] for k in ('target_column', 'primary_time_index', 'secondary_time_index'))
    categories = cfg['categorical_columns']
    if not target:
        raise ValueError('Provide Target Column/--target-column')
    if secondary and not primary:
        raise ValueError('Secondary Time Index requires Primary Time Index')
    roles = [target] + [c for c in (primary, secondary) if c] + categories
    if len(set(roles)) != len(roles) or any(not isinstance(c, str) or not c.strip() for c in roles):
        raise ValueError('Target, time and categorical roles must be unique nonblank names')
    selection = cfg['feature_columns'].strip()
    if selection.lower() == 'all':
        features = [c for c in history.columns if c not in roles]
    elif selection.lower() == 'none':
        features = []
    else:
        features = [c.strip() for c in selection.split(',')]
        if any(not c for c in features) or len(set(features)) != len(features):
            raise ValueError('Feature Columns must contain unique nonblank names')
        if set(features) & set(roles):
            raise ValueError('Feature Columns conflicts with target/time/series identity roles')
    required = roles + features
    if any(c not in history.columns for c in required):
        raise ValueError(f'Missing required columns: {[c for c in required if c not in history.columns]}')
    provenance = history.attrs.get('provenance')
    if provenance is None or len(provenance) != len(history):
        raise ValueError('History requires explicit source sheet/row provenance')
    groups, kind = {}, None
    for position, values in enumerate(history.itertuples(index=False, name=None)):
        row = dict(zip(history.columns, values))
        sheet, source_row = provenance[position]
        location = f'{sheet} row {source_row}'
        normalized, identity = {}, []
        for col in required:
            value = row[col]
            try:
                if value is None or isinstance(value, str) and not value.strip() or isinstance(value, (numbers.Real, dt.date, dt.datetime)) and pd.isna(value):
                    raise ValueError('missing required value')
                if col in categories:
                    if isinstance(value, str):
                        key = ('str', value)
                    elif isinstance(value, (bool, np.bool_)):
                        key = ('bool', bool(value))
                    elif isinstance(value, numbers.Real) and np.isfinite(value):
                        key = ('num', float(value))
                    else:
                        raise ValueError('identity must be string, boolean or finite numeric scalar')
                    identity.append(key)
                    normalized[col] = value
                elif col == primary:
                    if isinstance(value, (bool, np.bool_)):
                        raise ValueError('boolean time is invalid')
                    numeric = isinstance(value, numbers.Real) or isinstance(value, str) and re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', value.strip())
                    this_kind = 'numeric' if numeric else 'datetime'
                    if kind is not None and kind != this_kind:
                        raise ValueError(f'expected {kind} primary time')
                    kind = this_kind
                    if numeric:
                        value = float(value)
                        if not np.isfinite(value):
                            raise ValueError('nonfinite primary time')
                    else:
                        if isinstance(value, str):
                            if not re.fullmatch(r'\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?', value):
                                raise ValueError('use unambiguous naive ISO date/datetime')
                            value = dt.datetime.fromisoformat(value)
                        if not isinstance(value, (dt.date, dt.datetime)):
                            raise ValueError('expected numeric or naive date/datetime')
                        value = pd.Timestamp(value)
                        if value.tzinfo is not None:
                            raise ValueError('timezone-aware time is invalid')
                    normalized[col] = value
                else:
                    if isinstance(value, (bool, np.bool_)):
                        raise ValueError('boolean numeric value is invalid')
                    value = float(value)
                    if not np.isfinite(value):
                        raise ValueError('nonfinite numeric value')
                    normalized[col] = value
            except (ValueError, TypeError, OverflowError) as exc:
                raise ValueError(f'{location}, column {col}: {exc}') from exc
        key = tuple(identity)
        if key not in groups:
            groups[key] = dict(key=key, labels=[row[c] for c in categories], records=[])
        groups[key]['records'].append(dict(values=normalized, original={c: row[c] for c in (primary, secondary) if c}, sheet=sheet, row=source_row))
    if not groups:
        raise ValueError('History contains no valid records')
    for group in groups.values():
        records = group['records']
        if primary:
            records.sort(key=lambda r: tuple(r['values'][c] for c in (primary, secondary) if c))
            keys = [tuple(r['values'][c] for c in (primary, secondary) if c) for r in records]
            if len(set(keys)) != len(keys):
                raise ValueError(f'Series {group["key"]}: duplicate complete time keys; identify separate series or resolve observations')
        group['y'] = np.array([r['values'][target] for r in records], dtype=float)
        group['x'] = np.array([[r['values'][c] for c in features] for r in records], dtype=float).reshape(len(records), len(features))
    return list(groups.values()), features, kind


def _feature_row(y, x, context):
    """One available-information feature definition for training and recursion."""
    window = y[-3:]
    mean = sum(float(v) / len(window) for v in window)
    row = np.array([y[-1], mean, len(y), *x], dtype=float)
    if not np.isfinite(row).all():
        raise ValueError(f'{context}: nonfinite generated features')
    return row


def _new_estimator(family, cfg):
    return (TabPFNRegressor(n_estimators=cfg['n_ensemble'], random_state=42)
            if family == 'TabPFN' else GradientBoostingRegressor(n_estimators=cfg['n_gb_estimators'], random_state=42))


def _rollout(y, x, horizon, family, cfg, context):
    if len(y) < 3:
        return np.repeat(y[-1], horizon), 'Last-value baseline'
    matrix = np.array([_feature_row(y[:t], x[t-1], context) for t in range(1, len(y))])
    with np.errstate(over='ignore', invalid='ignore'):
        changes = np.diff(y)
    if not np.isfinite(changes).all():
        raise ValueError(f'{context}: nonfinite derived target changes in source prefix')
    model = _new_estimator(family, cfg)
    model.fit(matrix, changes)
    buffer, predictions = list(y), []
    for _ in range(horizon):
        change = float(np.asarray(model.predict(_feature_row(buffer, x[-1], context).reshape(1, -1))).reshape(-1)[0])
        level = float(buffer[-1]) + change
        if not np.isfinite(change) or not np.isfinite(level):
            raise ValueError(f'{context}: nonfinite predicted change or reconstructed target')
        buffer.append(level)
        predictions.append(level)
    return np.array(predictions), family


def _origin(record):
    return '; '.join([f'{k}={v}' for k, v in record['original'].items()] + [f'{record["sheet"]} row {record["row"]}'])


def _future_labels(group, primary, secondary, kind, horizon):
    unknown = 'Next-observation timing unknown'
    if not primary or secondary:
        return [None] * horizon, unknown
    times = [r['values'][primary] for r in group['records']]
    try:
        if kind == 'numeric' and len(times) >= 2:
            with np.errstate(over='ignore', invalid='ignore'):
                differences = np.diff(times)
            if np.isfinite(differences).all() and differences[0] > 0 and np.allclose(differences, differences[0], rtol=1e-9, atol=1e-12):
                values = [times[-1] + h * differences[0] for h in range(1, horizon+1)]
                if np.isfinite(values).all():
                    return values, 'Regular numeric step'
        elif kind == 'datetime' and len(times) >= 3:
            frequency = pd.infer_freq(pd.DatetimeIndex(times))
            if frequency:
                return [v.to_pydatetime() for v in pd.date_range(times[-1], periods=horizon+1, freq=frequency)[1:]], f'Inferred calendar frequency {frequency}'
    except (ValueError, OverflowError, pd.errors.OutOfBoundsDatetime):
        pass
    return [None] * horizon, unknown


def _metrics(errors, context):
    if not errors:
        return None, None
    values = np.asarray(errors, dtype=float)
    with np.errstate(over='ignore', invalid='ignore'):
        mae = float(np.mean(np.abs(values)))
        rmse = float(np.sqrt(np.mean(values**2)))
    if not np.isfinite(values).all() or not np.isfinite([mae, rmse]).all():
        raise ValueError(f'{context}: backtest error/aggregate overflow')
    return mae, rmse


def forecast_frames(history, cfg):
    """History-only per-series recursive prediction and bounded temporal backtest."""
    groups, features, kind = _prepare_series(history, cfg)
    horizon = cfg['future_rows']
    if horizon * len(groups) + 1 > 1048576 or len(cfg['categorical_columns']) + 5 > 16384:
        raise ValueError('Output exceeds Excel row/column capacity')
    family = cfg['model_type'].lower()
    if family not in ('tabpfn', 'gradientboosting'):
        raise ValueError('Choose Model Type TabPFN or GradientBoosting')
    family = 'TabPFN' if family == 'tabpfn' else 'GradientBoosting'
    fallback = None
    if family == 'TabPFN' and not HAS_TABPFN:
        family, fallback = 'GradientBoosting', 'tabpfn package unavailable'
    while True:
        try:
            rows, ledger, summaries = [], [], []
            for group in groups:
                y, x, records = group['y'], group['x'], group['records']
                n = len(y)
                block = min(horizon, max(1, n // 4))
                origins = [p for p in (n-2*block, n-block) if p >= 3]
                start = len(ledger)
                for prefix in origins:
                    context = f'Series {group["key"]}, origin {_origin(records[prefix-1])}, prefix from {_origin(records[0])}'
                    predicted, _ = _rollout(y[:prefix], x[:prefix], block, family, cfg, context)
                    for lead, prediction in enumerate(predicted, 1):
                        actual = float(y[prefix+lead-1])
                        error = float(prediction) - actual
                        if not np.isfinite(error):
                            raise ValueError(f'{context}: nonfinite backtest error at {_origin(records[prefix+lead-1])}')
                        ledger.append(dict(series=repr(group['key']), origin=_origin(records[prefix-1]), training_records=prefix, lead=lead, actual=actual, prediction=float(prediction), error=error))
                context = f'Series {group["key"]}, final origin {_origin(records[-1])}, prefix from {_origin(records[0])}'
                predictions, method = _rollout(y, x, horizon, family, cfg, context)
                labels, timing = _future_labels(group, cfg['primary_time_index'], cfg['secondary_time_index'], kind, horizon)
                for lead, (prediction, label) in enumerate(zip(predictions, labels), 1):
                    rows.append(group['labels'] + [_origin(records[-1]), lead] + ([label] if cfg['primary_time_index'] else []) + [float(prediction), method])
                group_ledger = ledger[start:]
                mae, rmse = _metrics([r['error'] for r in group_ledger], context)
                per_lead = []
                for lead in range(1, horizon+1):
                    errors = [r['error'] for r in group_ledger if r['lead'] == lead]
                    a, b = _metrics(errors, f'{context}, lead {lead}')
                    per_lead.append(dict(lead=lead, count=len(errors), mae=a, rmse=b, status='Measured' if errors else 'Unmeasured'))
                summaries.append(dict(series=repr(group['key']), identity_labels=repr(group['labels']), n=n, H=horizon, B=block, origins=len(origins), boundaries=[_origin(records[p-1]) for p in origins], origin=_origin(records[-1]), method=method, timing=timing, frozen_covariates=dict(zip(features, x[-1].tolist())), covariate_policy='Unused: insufficient history for selected model' if n < 3 else 'Fixed at last prefix observation during each rollout', mae=mae, rmse=rmse, count=len(group_ledger), unavailable_reason=None if origins else 'No candidate origin has at least three training records', per_lead=per_lead))
            break
        except RuntimeError as exc:
            message = str(exc).lower()
            if family != 'TabPFN' or not (re.search(r'\bauthentication\b', message) or
                    re.search(r'\bgated\b', message) and re.search(r'\b(?:access|repo|repository)\b', message)):
                raise
            family, fallback = 'GradientBoosting', str(exc)
    columns = list(cfg['categorical_columns'])
    for label in ['Forecast origin', 'Step'] + ([cfg['primary_time_index']] if cfg['primary_time_index'] else []) + ['Prediction', 'Method']:
        candidate, suffix = label, 2
        while candidate in columns:
            candidate = f'{label} ({suffix})'
            suffix += 1
        columns.append(candidate)
    mae, rmse = _metrics([r['error'] for r in ledger], 'Global backtest ledger')
    results = dict(model_type=family, fallback_reason=fallback, n_samples=sum(len(g['y']) for g in groups), n_features=3+len(features), n_predictions=len(rows), backtest_mae=mae, backtest_rmse=rmse, backtest_count=len(ledger), unevaluated_series=sum(not s['origins'] for s in summaries), selected_features=features, generated_features=['target:lag1', 'target:mean3', 'observation:index'] + [f'covariate:previous:{c}' for c in features], series=summaries, ledger=ledger)
    return pd.DataFrame(rows, columns=columns), results


def _diagnostic_payload(results, cfg):
    rows = [['Forecast diagnostics', 'Value'], ['Effective family', results['model_type']], ['Fallback reason', results['fallback_reason']], ['Target units', f'Original units of {cfg["target_column"]}'], ['Ordering', 'Within-series configured time keys' if cfg['primary_time_index'] else 'Configured sheet order then source row order'], ['Selected covariates', repr(results['selected_features'])], ['Generated features', repr(results['generated_features'])], ['Prediction interpretation', 'Point predictions only; no uncertainty intervals'], ['Backtest interpretation', 'Earlier smaller models under fixed-X assumptions; not measured accuracy of final fit or guaranteed future error'], ['Global ledger MAE', results['backtest_mae']], ['Global ledger RMSE', results['backtest_rmse']], ['Global ledger count', results['backtest_count']], ['Unevaluated series', results['unevaluated_series']]]
    for summary in results['series']:
        rows.append(['Series', summary['series']])
        rows.extend([[k, repr(v) if isinstance(v, (dict, list)) else v] for k, v in summary.items() if k != 'per_lead'])
        rows.append(['Lead', 'Count', 'MAE', 'RMSE', 'Status'])
        rows.extend([[p['lead'], p['count'], p['mae'], p['rmse'], p['status']] for p in summary['per_lead']])
    rows.append(['Series', 'Origin', 'Training records', 'Lead', 'Actual', 'Prediction', 'Signed error'])
    rows.extend([[r[k] for k in ('series', 'origin', 'training_records', 'lead', 'actual', 'prediction', 'error')] for r in results['ledger']])
    width = max(map(len, rows))
    return [row + [None] * (width-len(row)) for row in rows]


def run_forecast_configured(wb, **cfg):
    """Prevalidate every payload before clearing destinations; save once."""
    for key in ('header_row', 'data_start_row', 'n_ensemble', 'n_gb_estimators', 'future_rows'):
        value = cfg[key]
        if isinstance(value, bool) or not isinstance(value, numbers.Real) or not np.isfinite(value) or value < 1 or int(value) != value:
            raise ValueError(f'{key} must be a positive integer')
        cfg[key] = int(value)
    if cfg['data_start_row'] <= cfg['header_row'] or max(cfg['header_row'], cfg['data_start_row']) > 1048576:
        raise ValueError('Data Start Row must follow Header Row within Excel capacity')
    sources, output_name = cfg['data_sheets'], cfg['output_sheet']
    if not sources or any(not isinstance(n, str) or not n.strip() for n in sources):
        raise ValueError('Provide nonblank Data Sheets/--data-sheets')
    folded = [n.casefold() for n in sources]
    if len(set(folded)) != len(folded):
        raise ValueError('Historical sheets must be unique')
    if not isinstance(output_name, str) or not output_name.strip() or len(output_name) > 31 or re.search(r'[\\/*?:\[\]]', output_name) or output_name.startswith("'") or output_name.endswith("'"):
        raise ValueError('Output Sheet must be a legal Excel name (1–31 characters)')
    if output_name.casefold() == 'diagnostics' or set(folded) & {output_name.casefold(), 'diagnostics'}:
        raise ValueError('Input sheets, Output Sheet and Diagnostics must be distinct')
    actual = {s.name.casefold(): s.name for s in wb.sheets}
    if any(n not in actual for n in folded):
        raise ValueError(f'Missing source sheets: {[n for n in sources if n.casefold() not in actual]}')
    histories = [_read_frame(wb.sheets[actual[n.casefold()]], cfg['header_row'], cfg['data_start_row']) for n in sources]
    if any(list(f.columns) != list(histories[0].columns) for f in histories[1:]):
        raise ValueError('Historical sheets must have identical headers')
    history = pd.concat(histories, ignore_index=True)
    history.attrs['provenance'] = [p for f in histories for p in f.attrs['provenance']]
    output, results = forecast_frames(history, cfg)
    payload = [output.columns.tolist()] + output.values.tolist()
    diagnostics = _diagnostic_payload(results, cfg)
    if any(len(content) > 1048576 or max(map(len, content)) > 16384 for content in (payload, diagnostics)):
        raise ValueError('Output exceeds Excel row/column capacity')
    for name, content in ((output_name, payload), ('Diagnostics', diagnostics)):
        sheet = _get_or_create_sheet(wb, actual.get(name.casefold(), name), set(actual.values()))
        sheet.range('A1').value = content
    wb.save()
    return results


def _run_forecast(wb, cfg):
    return run_forecast_configured(wb, **cfg)


# =============================================================================
# MACRO ENTRY POINT (call via Alt+F8 / RunPython)
# =============================================================================

def run_forecast():
    """Macro entry point. Invoked from VBA via RunPython."""
    try:
        wb = xw.Book.caller()
        settings_path = get_settings_path()
        _check_settings_boundary(wb.fullname, settings_path)

        if not settings_sheet_exists(settings_path):
            create_default_settings(settings_path)
            wb.app.alert(
                f"Created '{SETTINGS_FILENAME}' with default configuration.\n\n"
                f"Location: {settings_path}\n\n"
                f"Please configure the '{SETTINGS_SHEET_NAME}' sheet, then run this macro again."
            )
            return

        raw = read_settings(settings_path)
        cfg = {key: raw.get(key, default) for key, _, _, _, default in SETTINGS_SCHEMA}

        results = _run_forecast(wb, cfg)

        wb.sheets[cfg['output_sheet']].activate()

        wb.app.alert(
            f"Forecast Complete!\n\n"
            f"Sheets processed: {len(cfg['data_sheets'])}\n"
            f"Target: {cfg['target_column']}\n"
            f"Features: {results['n_features']}\n"
            f"Samples: {results['n_samples']}\n"
            f"Backtest MAE: {results['backtest_mae']}\n"
            f"Backtest RMSE: {results['backtest_rmse']}\n"
            f"Backtest rows: {results['backtest_count']}; unevaluated series: {results['unevaluated_series']}\n\n"
            f"See 'Diagnostics' sheet for full details."
        )
    except Exception as e:
        xw.Book.caller().app.alert(f"Error: {e}")


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    """
    CLI entry point. python tabpfn_forecast.py <workbook.xlsx> [flags]

    Configuration comes from CLI flags, falling back to the settings workbook
    (default: PyexcelSettings.xlsx beside this script, 'Forecast' sheet),
    falling back to built-in defaults.
    """
    parser = argparse.ArgumentParser(description='TabPFN / GradientBoosting Forecast (CLI mode)')
    parser.add_argument('excel_path', help='Path to Excel workbook')
    for key, flags, label, kind, default in SETTINGS_SCHEMA:
        help_text = _SETTINGS_HELP.get(key, f'{label} (default: {default})')
        if kind == 'bool':
            parser.add_argument(*flags, dest=key, action='store_true', help=help_text)
        elif kind == 'int':
            parser.add_argument(*flags, dest=key, type=int, default=None, help=help_text)
        elif kind == 'float':
            parser.add_argument(*flags, dest=key, type=float, default=None, help=help_text)
        else:  # 'str' or 'list'
            parser.add_argument(*flags, dest=key, default=None, help=help_text)
    parser.add_argument('--settings', default=None,
                         help=f'Path to settings workbook (default: {SETTINGS_FILENAME} beside this script)')

    args = parser.parse_args()

    try:
        _check_settings_boundary(args.excel_path, args.settings or get_settings_path())
        if args.settings:
            if not os.path.exists(args.settings):
                print(f"ERROR: settings file not found: {args.settings}", file=sys.stderr)
                return 1
            if not settings_sheet_exists(args.settings):
                print(f"ERROR: settings file '{args.settings}' does not contain the '{SETTINGS_SHEET_NAME}' sheet", file=sys.stderr)
                return 1
            settings = read_settings(args.settings)
        else:
            default_settings_path = get_settings_path()
            settings = read_settings(default_settings_path) if settings_sheet_exists(default_settings_path) else {}

        cfg = {}
        reported = []
        for key, _flags, _label, kind, default in SETTINGS_SCHEMA:
            flag_value = getattr(args, key)
            if kind == 'bool':
                if flag_value:
                    value, from_settings = True, False
                elif key in settings:
                    value, from_settings = settings[key], True
                else:
                    value, from_settings = default, False
            elif kind == 'list':
                if flag_value is not None:
                    value = [s.strip() for s in flag_value.split(',') if s.strip()]
                    from_settings = False
                elif key in settings:
                    value = settings[key]
                    from_settings = True
                else:
                    value = default
                    from_settings = False
            else:
                if flag_value is not None:
                    value, from_settings = flag_value, False
                elif key in settings:
                    value, from_settings = settings[key], True
                else:
                    value, from_settings = default, False
            cfg[key] = value
            if from_settings:
                reported.append((key, value))

        for name, value in reported:
            print(f"settings: {name}={value}")

        wb = xw.Book(args.excel_path)
        results = _run_forecast(wb, cfg)

        print(f"\n{'='*50}")
        print(f"Forecast Complete!")
        print(f"Target: {cfg['target_column']}")
        print(f"Features: {results['n_features']}")
        print(f"Samples: {results['n_samples']}")
        print(f"Backtest MAE: {results['backtest_mae']}")
        print(f"Backtest RMSE: {results['backtest_rmse']}")
        print(f"Backtest rows: {results['backtest_count']}; unevaluated series: {results['unevaluated_series']}")
        print(f"Predictions: {results['n_predictions']}")
        print(f"See '{cfg['output_sheet']}' sheet for details.")

        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
