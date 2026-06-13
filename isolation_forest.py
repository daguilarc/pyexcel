#!/usr/bin/env python3
"""
Isolation Forest Anomaly Detection - Excel UDF Tool

Excel formulas:
- =ISOLATION_FOREST(...) - returns continuous anomaly scores (optional: ascending, epsilon)
- =ANOMALY_BINARY(...) - returns binary 0/1 predictions (with optional Bayesian sequential updating)

Workflow:
1. Run Isolation Forest on the provided data
2. Score the same data points for anomaly detection
3. Output variable array of scores/predictions

Monte Carlo Resampling:
- Uses stationary bootstrap (Politis & Romano, 1994) when sample size >= expected block length
- Falls back to i.i.d. bootstrap for small samples (< expected block length)
- Preserves time structure better than i.i.d. bootstrap for time series data
- Expected block length = 1/p (default p=0.1 gives threshold=10)

Uses xlwings UDF wrapper - no Settings sheet required.
"""

import xlwings as xw  # type: ignore
import pandas as pd
import numpy as np
import sys
import argparse
import time
from typing import Dict, List, Optional, Tuple, Union
from sklearn.ensemble import IsolationForest
from scipy.stats import beta


# =============================================================================
# GLOBAL HELPER FUNCTIONS
# =============================================================================

# Constants
RANDOM_SEED = 42  # Base seed for reproducibility

BOOL_SLOTS_ISOLATION = ('epsilon', 'ascending', 'binary', 'bayes_freq')
BOOL_SLOTS_ANOMALY = ('bayes_freq',)


def _is_contamination_value(arg: float) -> bool:
    return 0 < arg <= 0.5


def _is_mc_samples_value(arg: float) -> bool:
    return arg == 0 or (100 <= arg <= 10000)


def _is_array_index_value(arg: float) -> bool:
    return arg >= 0 and not _is_contamination_value(arg) and not _is_mc_samples_value(arg)


def _parse_udf_args(
    args: tuple,
    bool_slot_names: Tuple[str, ...],
) -> Tuple[List, Dict[str, bool], Optional[float], int, int]:
    """
    Parse Excel UDF *args: data arrays first, then optional params by type.
    Booleans use None sentinels so FALSE advances to the next slot correctly.
    """
    data_arrays: List = []
    bools: Dict[str, Optional[bool]] = {name: None for name in bool_slot_names}
    contamination: Optional[float] = None
    monte_carlo_samples: Optional[int] = None
    array_index: Optional[int] = None

    for arg in args:
        if isinstance(arg, bool):
            for name in bool_slot_names:
                if bools[name] is None:
                    bools[name] = arg
                    break
        elif isinstance(arg, (int, float)):
            if contamination is None and _is_contamination_value(arg):
                contamination = float(arg)
            elif monte_carlo_samples is None and _is_mc_samples_value(arg):
                monte_carlo_samples = int(arg)
            elif array_index is None and _is_array_index_value(arg):
                array_index = int(arg)
        else:
            data_arrays.append(arg)

    resolved_bools = {name: (bools[name] if bools[name] is not None else False) for name in bool_slot_names}
    return (
        data_arrays,
        resolved_bools,
        contamination,
        monte_carlo_samples if monte_carlo_samples is not None else 0,
        array_index if array_index is not None else 0,
    )


def _arrays_to_numpy(data_arrays: List) -> Tuple[List[np.ndarray], List[int], np.ndarray, np.ndarray]:
    arrays: List[np.ndarray] = []
    array_sizes: List[int] = []
    for arr in data_arrays:
        arr_np = np.asarray(arr, dtype=np.float64)
        if arr_np.ndim == 1:
            arr_np = arr_np.reshape(1, -1)
        elif arr_np.ndim != 2:
            raise ValueError("Data arrays must be 1D or 2D")
        arrays.append(arr_np)
        array_sizes.append(arr_np.shape[0])

    n_cols = arrays[0].shape[1]
    for i, arr in enumerate(arrays[1:], 1):
        if arr.shape[1] != n_cols:
            raise ValueError(
                f"All arrays must have same number of columns. Array 1: {n_cols}, Array {i+1}: {arr.shape[1]}"
            )

    X = np.vstack(arrays) if len(arrays) > 1 else arrays[0]
    valid_mask = ~np.isnan(X).any(axis=1)
    return arrays, array_sizes, X, valid_mask


def _apply_score_transforms(scores: np.ndarray, ascending: bool, epsilon: bool) -> np.ndarray:
    if ascending:
        scores = -scores
    if epsilon:
        percentile_10 = np.percentile(scores, 0.1)
        epsilon_val = max(1e-6, percentile_10 - scores.min() + 1e-6)
        scores = scores - percentile_10 + epsilon_val
    return scores


def _scores_to_binary(scores: np.ndarray, contamination: float, ascending: bool) -> np.ndarray:
    pct = (100 - contamination * 100) if ascending else (contamination * 100)
    threshold_val = np.percentile(scores, pct)
    if ascending:
        return (scores >= threshold_val).astype(int)
    return (scores <= threshold_val).astype(int)


def _fit_score_array(
    arr_clean: np.ndarray,
    contamination: float,
    monte_carlo_samples: int,
    X_cumulative: np.ndarray,
) -> np.ndarray:
    model = IsolationForest(contamination=contamination, random_state=RANDOM_SEED)
    model.fit(X_cumulative)
    if monte_carlo_samples > 0:
        return monte_carlo_score_samples(arr_clean, contamination, monte_carlo_samples)
    return model.score_samples(arr_clean)


def _bayesian_cumulative_binary(
    arrays: List[np.ndarray],
    array_sizes: List[int],
    prior_contamination: float,
    monte_carlo_samples: int,
    ascending: bool,
    epsilon: bool,
    n_arrays: int,
) -> Tuple[List[np.ndarray], np.ndarray]:
    alpha = prior_contamination * 100
    beta_param = (1 - prior_contamination) * 100
    all_binary: List[np.ndarray] = []
    cumulative_arrays: List[np.ndarray] = []

    for arr_idx, arr in enumerate(arrays):
        arr_valid_mask = ~np.isnan(arr).any(axis=1)
        arr_clean = arr[arr_valid_mask]

        if len(arr_clean) < 2:
            all_binary.append(np.full(array_sizes[arr_idx], np.nan))
            continue

        cumulative_arrays.append(arr_clean)
        current_contamination = min(beta.mean(alpha, beta_param), 0.5)
        X_cumulative = np.vstack(cumulative_arrays) if len(cumulative_arrays) > 1 else cumulative_arrays[0]

        arr_scores = _fit_score_array(arr_clean, current_contamination, monte_carlo_samples, X_cumulative)
        arr_scores = _apply_score_transforms(arr_scores, ascending, epsilon)
        arr_binary = _scores_to_binary(arr_scores, current_contamination, ascending)

        if arr_idx < n_arrays - 1:
            alpha += np.sum(arr_binary)
            beta_param += len(arr_clean) - np.sum(arr_binary)

        arr_result = np.full(array_sizes[arr_idx], np.nan)
        arr_result[arr_valid_mask] = arr_binary
        all_binary.append(arr_result)

    return all_binary, np.concatenate(all_binary)


def _bayesian_cumulative_predictions(
    arrays: List[np.ndarray],
    array_sizes: List[int],
    prior_contamination: float,
    monte_carlo_samples: int,
    n_arrays: int,
) -> Tuple[List[np.ndarray], List[np.ndarray], np.ndarray]:
    """ANOMALY_BINARY Bayesian mode: descending scores only (no ascending/epsilon)."""
    alpha = prior_contamination * 100
    beta_param = (1 - prior_contamination) * 100
    all_predictions: List[np.ndarray] = []
    all_valid_masks: List[np.ndarray] = []
    cumulative_arrays: List[np.ndarray] = []

    for arr_idx, arr in enumerate(arrays):
        arr_valid_mask = ~np.isnan(arr).any(axis=1)
        arr_clean = arr[arr_valid_mask]

        if len(arr_clean) < 2:
            all_predictions.append(np.full(array_sizes[arr_idx], np.nan))
            all_valid_masks.append(arr_valid_mask)
            continue

        cumulative_arrays.append(arr_clean)
        current_contamination = min(beta.mean(alpha, beta_param), 0.5)
        X_cumulative = np.vstack(cumulative_arrays) if len(cumulative_arrays) > 1 else cumulative_arrays[0]

        scores = _fit_score_array(arr_clean, current_contamination, monte_carlo_samples, X_cumulative)
        arr_binary = (scores <= np.percentile(scores, current_contamination * 100)).astype(int)

        if arr_idx < n_arrays - 1:
            alpha += np.sum(arr_binary)
            beta_param += len(arr_clean) - np.sum(arr_binary)

        arr_result = np.full(array_sizes[arr_idx], np.nan)
        arr_result[arr_valid_mask] = arr_binary
        all_predictions.append(arr_result)
        all_valid_masks.append(arr_valid_mask)

    return all_predictions, all_valid_masks, np.concatenate(all_predictions)


def _build_array_column(
    arr_size: int,
    arr_valid_mask: np.ndarray,
    source: np.ndarray,
    source_start: int,
    source_count: int,
    pad_to: int,
) -> np.ndarray:
    col = np.full(arr_size, np.nan)
    col[arr_valid_mask] = source[source_start:source_start + source_count]
    if arr_size < pad_to:
        col = np.pad(col, (0, pad_to - arr_size), constant_values=np.nan)
    return col


def _build_excel_output(
    array_sizes: List[int],
    valid_mask: np.ndarray,
    scores: np.ndarray,
    n_arrays: int,
    input_single_row: bool,
    array_index: int,
    binary: bool,
    binary_predictions: Optional[np.ndarray],
    all_binary: Optional[List[np.ndarray]],
    bayes_freq: bool,
) -> Union[float, List, List[List]]:
    if input_single_row:
        if binary and binary_predictions is not None:
            return [float(scores[0]), int(binary_predictions[0])]
        return float(scores[0])

    if n_arrays == 1:
        if binary and binary_predictions is not None:
            return np.column_stack([scores, binary_predictions]).tolist()
        return scores.reshape(-1, 1).tolist()

    array_cumsum = np.cumsum([0] + array_sizes)
    valid_cumsum = np.concatenate([[0], np.cumsum(valid_mask)])
    max_rows = max(array_sizes)

    if array_index > 0:
        target_idx = array_index - 1
        array_start_idx = array_cumsum[target_idx]
        arr_size = array_sizes[target_idx]
        arr_valid_mask = valid_mask[array_start_idx:array_start_idx + arr_size]
        arr_valid_count = int(np.sum(arr_valid_mask))
        score_idx = valid_cumsum[array_start_idx]

        arr_score = np.full(arr_size, np.nan)
        arr_score[arr_valid_mask] = scores[score_idx:score_idx + arr_valid_count]

        if binary and binary_predictions is not None:
            if bayes_freq and all_binary is not None:
                arr_binary = all_binary[target_idx]
            else:
                arr_binary = np.full(arr_size, np.nan)
                arr_binary[arr_valid_mask] = binary_predictions[score_idx:score_idx + arr_valid_count]
            return np.column_stack([arr_score, arr_binary]).tolist()
        return arr_score.reshape(-1, 1).tolist()

    columns: List[np.ndarray] = []
    for arr_idx, arr_size in enumerate(array_sizes):
        array_start_idx = array_cumsum[arr_idx]
        arr_valid_mask = valid_mask[array_start_idx:array_start_idx + arr_size]
        arr_valid_count = int(np.sum(arr_valid_mask))
        score_idx = valid_cumsum[array_start_idx]

        arr_score = _build_array_column(
            arr_size, arr_valid_mask, scores, score_idx, arr_valid_count, max_rows
        )

        if binary and binary_predictions is not None:
            if bayes_freq and all_binary is not None:
                arr_binary = all_binary[arr_idx]
                if arr_size < max_rows:
                    arr_binary = np.pad(arr_binary, (0, max_rows - arr_size), constant_values=np.nan)
            else:
                arr_binary = _build_array_column(
                    arr_size, arr_valid_mask, binary_predictions, score_idx, arr_valid_count, max_rows
                )
            columns.append(arr_score)
            columns.append(arr_binary)
        else:
            columns.append(arr_score)

    return np.array(columns).T.tolist()


def monte_carlo_score_samples(X_clean, contamination, monte_carlo_samples, stationary_bootstrap_p=0.1):
    """
    Monte Carlo resampling for Isolation Forest scoring.
    Uses stationary bootstrap (Politis & Romano, 1994) with fallback to i.i.d. bootstrap for small samples.
    
    Stationary bootstrap algorithm:
    - Starts at random position
    - With probability p, continues to next observation (wraps around at end)
    - With probability (1-p), jumps to random position
    - Expected block length = 1/p (default p=0.1 gives expected block length of 10)
    - Preserves time structure better than i.i.d. bootstrap for time series data
    - Falls back to i.i.d. bootstrap if sample size < expected block length (1/p)
    
    Args:
        X_clean: Clean data array (no NaN rows)
        contamination: Contamination rate (0-0.5)
        monte_carlo_samples: Number of Monte Carlo iterations (100-10000)
        stationary_bootstrap_p: Probability parameter for stationary bootstrap (default 0.1)
            Controls expected block length = 1/p. Lower p = longer blocks.
            If sample size < 1/p, falls back to i.i.d. bootstrap.
    
    Returns:
        Aggregated scores (mean across all Monte Carlo iterations)
    """
    throttle_interval = 0.005  # 100 samples per 0.5 seconds (0.5/100)
    start_time = time.time()
    n_samples = len(X_clean)
    scores_buf = np.empty((monte_carlo_samples, n_samples), dtype=np.float64)
    
    # Calculate minimum sample threshold from expected block length
    expected_block_length = 1.0 / stationary_bootstrap_p
    min_samples_threshold = int(np.ceil(expected_block_length))
    
    # Use i.i.d. bootstrap if sample size is too small for stationary bootstrap
    use_stationary = n_samples >= min_samples_threshold
    
    # Create seed sequence for proper parallelization support
    child_ss = np.random.SeedSequence(RANDOM_SEED).spawn(monte_carlo_samples)
    
    # Pre-allocate indices array for stationary bootstrap (O(1) optimization)
    indices = np.empty(n_samples, dtype=np.int64) if use_stationary else None
    
    for sample in range(monte_carlo_samples):
        rng = np.random.default_rng(child_ss[sample])
        
        # Generate bootstrap indices
        if use_stationary:
            # Stationary bootstrap algorithm (Politis & Romano, 1994)
            current_pos = rng.integers(0, n_samples)
            for i in range(n_samples):
                indices[i] = current_pos
                # With probability p, continue to next (wrap around); with probability (1-p), jump to random
                if rng.random() < stationary_bootstrap_p:
                    current_pos = (current_pos + 1) % n_samples
                else:
                    current_pos = rng.integers(0, n_samples)
            X_boot = X_clean[indices]
        else:
            # i.i.d. bootstrap fallback for small samples
            X_boot = X_clean[rng.integers(0, n_samples, size=n_samples)]
        
        scores_buf[sample] = IsolationForest(
            contamination=contamination,
            random_state=rng.integers(0, 2**42)
        ).fit(X_boot).score_samples(X_clean)

        sleep_time = (sample + 1) * throttle_interval - (time.time() - start_time)
        if sleep_time > 0:
            time.sleep(sleep_time)

    return np.mean(scores_buf, axis=0)


# =============================================================================
# EXCEL UDF FORMULA (use as: =ISOLATION_FOREST(...))
# =============================================================================

@xw.func
def ISOLATION_FOREST(
    *args
) -> Union[float, List[List]]:
    """
    Excel UDF for Isolation Forest anomaly scoring - returns one score per data point.
    
    Isolation Forest is an unsupervised algorithm that detects anomalies by learning patterns
    from the data itself. It trains on the provided data and scores each data point for anomaly detection.
    
    Parameters (in order):
        *args: Input data arrays followed by optional parameters in this specific order:
              - One or more data arrays (rows x columns format)
              - Each row = one data point, each column = one feature
              - Multiple arrays are combined (all must have same number of columns)
              
              Optional parameters (can be specified in any order after data arrays):
              Parameters are identified by type/pattern, so you can specify them in any order.
              - contamination: 0-0.5 decimal (e.g., 0.1 for 10%)
              - monte_carlo_samples: 0 or 100-10000 integer
              - ascending: TRUE/FALSE boolean
              - epsilon: TRUE/FALSE boolean (requires ascending=TRUE)
              - binary: TRUE/FALSE boolean (adds binary column next to scores)
              - bayes_freq: TRUE/FALSE boolean (only used when binary=TRUE, enables Bayesian updating)
              - array_index: 0 or 1-indexed integer (for multiple arrays)
              
              - contamination: Expected proportion of anomalies as decimal (0-0.5)
                     Default: 0.1 (10%)
                     Examples: 0.05 = 5%, 0.1 = 10%, 0.5 = 50% (max)
                     Maximum is 0.5 (50%) based on the fundamental assumption of Isolation Forest
                     (Liu et al., 2008) that anomalies are "few and different" - they should be
                     a minority of the data. Beyond 0.5, the distinction between "normal" and
                     "anomalous" becomes meaningless, as you'd be assuming more than half your
                     data is anomalous, which contradicts the algorithm's core premise.
                     Reference: Liu, F. T., Ting, K. M., & Zhou, Z.-H. (2008). Isolation Forest.
                     In 2008 Eighth IEEE International Conference on Data Mining (pp. 413-422).
              - monte_carlo_samples: Number of Monte Carlo resampling iterations (0 or 100-10000)
                     Default: 0 (off)
                     Must be 0 (disabled) or between 100-10000
                     Throttled to 100 samples per 0.5 seconds - only change if your computer can handle it
              - ascending: Boolean, if TRUE scores are ascending (high score = highly anomalous)
                     Default: FALSE (descending: lower score = more anomalous)
              - epsilon: Boolean, if TRUE shifts scores using 0.1th percentile threshold to ensure all values > 0
                     Default: FALSE
                     Only works when ascending=TRUE (requires ascending to be enabled)
                     This is for gamma distribution modeling of the scores (gamma GLM requires strictly positive values)
                     Uses 0.1th percentile as shift threshold (robust to outliers) with epsilon to guarantee positivity
              - binary: Boolean, if TRUE adds binary column (0/1) next to scores
                     Default: FALSE (scores only)
                     When TRUE: Returns 2 columns per array (score, binary)
                     Binary uses percentile threshold from contamination rate
              - bayes_freq: Boolean, if TRUE uses Bayesian updating for binary column (only when binary=TRUE)
                     Default: FALSE (frequentist binary from main scores)
                     TRUE: Fits models cumulatively (arrays 1, then 1+2, then 1+2+3, etc.) for binary
                     Scores each array separately but fits on cumulative data
                     Updates contamination using Beta-Binomial conjugate prior
                     Only works with multiple arrays
                     
                     IMPORTANT LIMITATION: The procedure updates a Beta prior using counts that are
                     defined by the parameter being updated (threshold = posterior mean percentile),
                     so the posterior reflects the stability of the cutoff, not evidence about the
                     true population anomaly rate. This is a self-consistency measure, not Bayesian
                     inference about an unknown parameter from independent observations.
              - array_index: Integer, 1-indexed array number to return (0 = return all arrays)
                     Default: 0 (return all columns)
                     Examples: 1 = first array only, 2 = second array only
                     Only works with multiple arrays. Returns error if index exceeds number of arrays.
    
    Returns:
        Anomaly scores (by default descending: lower = more anomalous, unless ascending=TRUE):
        - Single value if input is single row (1 data point), or [score, binary] if binary=TRUE
        - Column array (n_rows x 1) if single array input, or (n_rows x 2) if binary=TRUE
        - Multiple columns (one per input array) if multiple arrays, or (n_arrays × 2) if binary=TRUE
        When binary=TRUE: Columns are interleaved (score1, binary1, score2, binary2, ...)
    
    Usage (parameters can be in any order after data arrays):
        =ISOLATION_FOREST(A2:E100)  # Single array (default 10% contamination)
        =ISOLATION_FOREST(A2:E100, 0.05)  # 5% contamination
        =ISOLATION_FOREST(A2:E100, 1000)  # 1000 Monte Carlo samples (contamination=default)
        =ISOLATION_FOREST(A2:E100, TRUE)  # ascending=TRUE (contamination=default)
        =ISOLATION_FOREST(A2:E100, 0.05, 1000)  # 5% contamination + 1000 Monte Carlo samples
        =ISOLATION_FOREST(A2:E100, 1000, 0.05)  # Same as above (order doesn't matter)
        =ISOLATION_FOREST(A2:E100, 0.05, 1000, TRUE, TRUE)  # All params in any order
        =ISOLATION_FOREST(A2:E50, G2:J50)  # Multiple arrays (combined)
        =ISOLATION_FOREST(A2:E50, G2:J50, 2)  # Multiple arrays, return only 2nd array
        =ISOLATION_FOREST(A2:E50, G2:J50, 0.1, TRUE, 2)  # Multiple arrays with params in any order
    
    Note: All input arrays must have the same number of columns (features).
          Contamination must be between 0 and 0.5 (exclusive of 0, inclusive of 0.5).
          Monte Carlo samples must be 0 (off) or between 100-10000 (no values like 3 allowed).
          
    Practical Limits:
    - Excel formula length: Maximum 8,192 characters (theoretical limit ~700-1000 arrays)
    - Memory: All arrays are loaded and combined upfront (O(n) memory where n = total rows).
              Memory usage ≈ total_rows × features × 8 bytes × 2 (data + scores arrays + overhead).
              With 16 GB RAM (Excel/OS use ~4-6 GB, leaving ~10-12 GB available), memory becomes
              noticeable around: 200-300 arrays with 10,000 rows × 10 features ≈ 400-600 MB data
              (800 MB-1.2 GB total with overhead). Significant memory pressure occurs around:
              500-800 arrays with 10,000 rows × 10 features ≈ 1-1.6 GB data (2-3.2 GB total).
              Beyond 1,000+ arrays, you risk memory errors or Excel crashes.
    - Processing time: Frequentist mode is O(n log n) where n = total rows (combines all arrays).
                      Bayesian mode is O(n_arrays × n log n) - processes each array separately.
                      With 16 GB RAM, recommended practical limit: 10-30 arrays for frequentist,
                      5-15 arrays for Bayesian mode (due to sequential processing).
                      Beyond these limits, you will experience noticeable processing lag, especially
                      in Bayesian mode or with Monte Carlo sampling enabled.
    """
    try:
        if not args:
            return "#ERROR: At least one data array is required"

        data_arrays, bools, contamination_raw, monte_carlo_samples, array_index = _parse_udf_args(
            args, BOOL_SLOTS_ISOLATION
        )
        epsilon = bools['epsilon']
        ascending = bools['ascending']
        binary = bools['binary']
        bayes_freq = bools['bayes_freq']

        if not data_arrays:
            return "#ERROR: At least one data array is required"

        contamination = 0.1 if contamination_raw is None else contamination_raw
        if contamination <= 0 or contamination > 0.5:
            return "#ERROR: contamination must be > 0 and <= 0.5 (0 < contamination <= 0.5)"

        if monte_carlo_samples != 0 and (monte_carlo_samples < 100 or monte_carlo_samples > 10000):
            return "#ERROR: monte_carlo_samples must be 0 (off) or between 100-10000"

        if epsilon and not ascending:
            return "#ERROR: epsilon requires ascending=TRUE. Epsilon only works with ascending scores."

        if bayes_freq and not binary:
            return "#ERROR: bayes_freq only works when binary=TRUE"

        try:
            arrays, array_sizes, X, valid_mask = _arrays_to_numpy(data_arrays)
        except ValueError as e:
            return f"#ERROR: {e}"

        n_arrays = len(arrays)
        input_single_row = X.shape[0] == 1
        X_clean = X[valid_mask]

        if len(X_clean) < 2:
            return "#ERROR: Need at least 2 valid data points"

        model = IsolationForest(contamination=contamination, random_state=RANDOM_SEED)
        model.fit(X_clean)

        if monte_carlo_samples > 0:
            scores = monte_carlo_score_samples(X_clean, contamination, monte_carlo_samples)
        else:
            scores = model.score_samples(X_clean)

        scores = _apply_score_transforms(scores, ascending, epsilon)

        binary_predictions = None
        all_binary = None
        if binary:
            if bayes_freq and n_arrays > 1:
                all_binary, binary_predictions = _bayesian_cumulative_binary(
                    arrays, array_sizes, contamination, monte_carlo_samples,
                    ascending, epsilon, n_arrays,
                )
            else:
                binary_predictions = _scores_to_binary(scores, contamination, ascending)

        if array_index > 0:
            if array_index > n_arrays:
                return f"#ERROR: array_index {array_index} exceeds number of arrays ({n_arrays})"
            if n_arrays == 1:
                return "#ERROR: array_index can only be used with multiple arrays"

        return _build_excel_output(
            array_sizes, valid_mask, scores, n_arrays, input_single_row, array_index,
            binary, binary_predictions, all_binary, bayes_freq,
        )

    except Exception as e:
        return f"#ERROR: {e}"


@xw.func
def ANOMALY_BINARY(
    *args
) -> Union[int, List[List]]:
    """
    Excel UDF for Isolation Forest binary anomaly detection - returns 0 or 1 per data point.
    
    Returns binary predictions: 1 = anomaly, 0 = normal.
    
    Parameters (in order):
        *args: Input data arrays followed by optional parameters in this specific order:
              - One or more data arrays (rows x columns format)
              - Each row = one data point, each column = one feature
              - Multiple arrays are combined (all must have same number of columns)
              
              Optional parameters (can be specified in any order after data arrays):
              Parameters are identified by type/pattern, so you can specify them in any order.
              - prior_contamination: 0-0.5 decimal (e.g., 0.1 for 10%)
              - monte_carlo_samples: 0 or 100-10000 integer (Monte Carlo resampling for scores)
              - bayes_freq: TRUE/FALSE boolean
              - array_index: 0 or 1-indexed integer (for multiple arrays)
              
              - prior_contamination: Prior contamination rate for first array (0-0.5 as decimal)
                     Default: 0.1 (10%)
                     Examples: 0.05 = 5%, 0.1 = 10%, 0.5 = 50% (max)
                     Maximum is 0.5 (50%) based on the fundamental assumption of Isolation Forest
                     (Liu et al., 2008) that anomalies are "few and different" - i.e., they should be
                     a minority of the data. Beyond 0.5, the distinction between "normal" and
                     "anomalous" becomes meaningless, as you'd be assuming more than half your
                     data is anomalous, which contradicts the algorithm's core premise.
                     Reference: Liu, F. T., Ting, K. M., & Zhou, Z.-H. (2008). Isolation Forest.
                     In 2008 Eighth IEEE International Conference on Data Mining (pp. 413-422).

                     When bayes_freq=TRUE, this is the prior for the first array
                     When bayes_freq=FALSE (default), this is the contamination for all arrays
                     Note: Higher prior_contamination sets a higher percentile threshold, which means
                     more points are counted as anomalies initially. This leads to higher variance
                     (uncertainty) in the Beta posterior distribution itself, meaning the contamination
                     rate estimate has more uncertainty. The threshold recalibrates based on the posterior
                     mean, which with higher variance can swing more, but update magnitude also depends
                     on the actual data observed.

                     Obviously, this doesn't do anything if you have only one array.

                     A contamination rate of 0% will always return 0 for all points and never update.
              - monte_carlo_samples: Number of Monte Carlo resampling iterations (0 or 100-10000)
                     Default: 0 (off)
                     Must be 0 (disabled) or between 100-10000
                     Throttled to 100 samples per 0.5 seconds - only change if your computer can handle it
                     Uses bootstrap resampling to aggregate scores across iterations, providing more
                     robust anomaly detection. When enabled, scores are computed via Monte Carlo method
                     and then thresholded to produce binary predictions.
              - bayes_freq: Boolean, if TRUE uses Bayesian updating of contamination across arrays
                     Default: FALSE (frequentist, combines all arrays, fits once)
                     TRUE (BAYES): Fits each array sequentially, updates contamination using
                     Beta-Binomial conjugate prior. Uses continuous anomaly scores from 
                     score_samples() to determine anomalies based on actual data structure, then
                     updates contamination rate based on score distribution. 
                     
                     Bayes theorem recap: Beta(α, β) prior where α/(α+β) = prior_contamination,
                     then updates: α += anomalies_found, β += normal_points_found for each array.
                     
                     IMPORTANT LIMITATION: The procedure updates a Beta prior using counts that are
                     defined by the parameter being updated (threshold = posterior mean percentile),
                     so the posterior reflects the stability of the cutoff, not evidence about the
                     true population anomaly rate. This is a self-consistency measure, not Bayesian
                     inference about an unknown parameter from independent observations.

              - array_index: Integer, 1-indexed array number to return (0 = return all arrays)
                     Default: 0 (return all columns)
                     Examples: 1 = first array only, 2 = second array only
                     Only works with multiple arrays. Returns error if index exceeds number of arrays.
    
    Returns:
        Binary predictions (1 = anomaly, 0 = normal):
        - Single value if input is single row (1 data point)
        - Column array (n_rows x 1) if single array input
        - Multiple columns (one per input array) if multiple arrays, padded to max length
    
    Usage (parameters can be in any order after data arrays):
        =ANOMALY_BINARY(A2:E100)  # Single array (default 10% contamination)
        =ANOMALY_BINARY(A2:E100, 0.05)  # 5% contamination
        =ANOMALY_BINARY(A2:E100, 1000)  # 1000 Monte Carlo samples (contamination=default)
        =ANOMALY_BINARY(A2:E50, G2:J50)  # Multiple arrays (combined, no Bayesian updating)
        =ANOMALY_BINARY(A2:E50, G2:J50, TRUE)  # Multiple arrays with Bayesian updating (prior=default)
        =ANOMALY_BINARY(A2:E50, G2:J50, 0.1, TRUE)  # Multiple arrays with 10% prior + Bayesian updating
        =ANOMALY_BINARY(A2:E50, G2:J50, 0.1, 1000, TRUE)  # With Monte Carlo + Bayesian updating
        =ANOMALY_BINARY(A2:E50, G2:J50, 2)  # Multiple arrays, return only 2nd array
        =ANOMALY_BINARY(A2:E50, G2:J50, 0.1, TRUE, 2)  # Multiple arrays with params in any order
    
    Note: All input arrays must have the same number of columns (features).
          Contamination must be between 0 and 0.5 (exclusive of 0, inclusive of 0.5).
          
    Practical Limits:
    - Excel formula length: Maximum 8,192 characters (theoretical limit ~700-1000 arrays)
    - Memory: Isolation Forest is O(n) in memory where n = total rows across all arrays.
              For 16 GB RAM, typical limits: ~10-50 arrays with 1,000-10,000 rows each.
              Memory usage ≈ (total_rows × features × 8 bytes) + model overhead.
              Example: 50 arrays × 5,000 rows × 10 features ≈ 20 MB (negligible for 16 GB).
    - Processing time: Frequentist mode is O(n log n) where n = total rows (combines all arrays).
                      Bayesian mode is O(n_arrays × n log n) - processes each array separately.
                      With 16 GB RAM, recommended practical limit: 10-30 arrays for frequentist,
                      5-15 arrays for Bayesian mode (due to sequential processing).
                      Beyond this, Excel may become slow or unresponsive.
    """
    try:
        if not args:
            return "#ERROR: At least one data array is required"

        data_arrays, bools, prior_raw, monte_carlo_samples, array_index = _parse_udf_args(
            args, BOOL_SLOTS_ANOMALY
        )
        bayes_freq = bools['bayes_freq']

        if not data_arrays:
            return "#ERROR: At least one data array is required"

        prior_contamination = 0.1 if prior_raw is None else prior_raw
        if prior_contamination <= 0 or prior_contamination > 0.5:
            return "#ERROR: prior_contamination must be > 0 and <= 0.5 (0 < prior_contamination <= 0.5)"

        if monte_carlo_samples != 0 and (monte_carlo_samples < 100 or monte_carlo_samples > 10000):
            return "#ERROR: monte_carlo_samples must be 0 (off) or between 100-10000"

        try:
            arrays, array_sizes, X, valid_mask = _arrays_to_numpy(data_arrays)
        except ValueError as e:
            return f"#ERROR: {e}"

        n_arrays = len(arrays)
        input_single_row = X.shape[0] == 1

        if bayes_freq and n_arrays > 1:
            _, all_valid_masks, predictions = _bayesian_cumulative_predictions(
                arrays, array_sizes, prior_contamination, monte_carlo_samples, n_arrays,
            )
            valid_mask = np.concatenate(all_valid_masks)
        else:
            valid_mask = ~np.isnan(X).any(axis=1)
            X_clean = X[valid_mask]
            if len(X_clean) < 2:
                return "#ERROR: Need at least 2 valid data points"
            model = IsolationForest(contamination=prior_contamination, random_state=RANDOM_SEED)
            model.fit(X_clean)
            predictions = (model.predict(X_clean) == -1).astype(int)

        if array_index > 0:
            if array_index > n_arrays:
                return f"#ERROR: array_index {array_index} exceeds number of arrays ({n_arrays})"
            if n_arrays == 1:
                return "#ERROR: array_index can only be used with multiple arrays"

        if input_single_row:
            return int(predictions[0])
        if n_arrays == 1:
            return predictions.reshape(-1, 1).tolist()

        return _build_excel_output(
            array_sizes, valid_mask, predictions, n_arrays, False, array_index,
            False, None, None, False,
        )

    except Exception as e:
        return f"#ERROR: {e}"


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    """
    CLI entry point - processes first sheet in Excel file.
    
    Note: Assumes row 1 contains column headers. Data starts at row 2.
    All numeric columns are used as features. Rows with any NaN values are excluded.
    """
    parser = argparse.ArgumentParser(description='Isolation Forest Anomaly Detection (CLI mode)')
    parser.add_argument('excel_path', help='Path to Excel workbook')
    parser.add_argument('--sheet', help='Sheet name to process (default: first sheet)')
    parser.add_argument('--contamination', type=float, default=0.1, help='Expected proportion of anomalies (default: 0.1)')
    parser.add_argument('--monte-carlo-samples', type=int, default=0, help='Monte Carlo resampling iterations (0=off, 100-10000)')
    parser.add_argument('--ascending', action='store_true', help='Use ascending scores (high score = highly anomalous)')
    parser.add_argument('--epsilon', action='store_true', help='Shift scores to ensure all values > 0 (requires --ascending)')
    parser.add_argument('--output-sheet', default='AnomalyScores', help='Output sheet name (default: AnomalyScores)')
    
    args = parser.parse_args()
    
    try:
        wb = xw.Book(args.excel_path)
        
        # Get sheet to process
        if args.sheet:
            sheet = wb.sheets[args.sheet]
        else:
            sheet = wb.sheets[0]
        
        # Read data
        # Note: header=1 assumes row 1 contains column headers, data starts at row 2
        if (data_range := sheet.used_range) is None:
            print("ERROR: No data found in sheet!", file=sys.stderr)
            return 1
        
        raw_data = data_range.options(pd.DataFrame, header=1).value
        if raw_data is None or len(raw_data) == 0:
            print("ERROR: No data found in sheet!", file=sys.stderr)
            return 1
        
        # Remove entirely NaN rows
        df = pd.DataFrame(raw_data).dropna(how='all').reset_index(drop=True)
        
        # Use all numeric columns (cache result)
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        n_features = len(numeric_cols)  # Cache for later use
        if n_features == 0:
            print("ERROR: No numeric columns found!", file=sys.stderr)
            return 1
        
        # Remove NaN rows (vectorized)
        valid_mask = ~np.isnan(X := df[numeric_cols].values.astype(np.float64)).any(axis=1)
        X_clean = X[valid_mask]
        n_samples_clean = len(X_clean)  # Cache for later use
        
        if n_samples_clean < 2:
            print("ERROR: Need at least 2 valid data points!", file=sys.stderr)
            return 1
        
        # Validate contamination
        if args.contamination <= 0 or args.contamination > 0.5:
            print("ERROR: contamination must be > 0 and <= 0.5", file=sys.stderr)
            return 1
        
        # Validate Monte Carlo samples
        if args.monte_carlo_samples != 0 and (args.monte_carlo_samples < 100 or args.monte_carlo_samples > 10000):
            print("ERROR: monte_carlo_samples must be 0 (off) or between 100-10000", file=sys.stderr)
            return 1
        
        # Validate epsilon requires ascending
        if args.epsilon and not args.ascending:
            print("ERROR: epsilon requires --ascending. Epsilon only works with ascending scores.", file=sys.stderr)
            return 1
        
        # Fit and score
        # n_estimators defaults to 100 in sklearn (optimal: path lengths converge)
        # max_samples="auto" (sklearn default) automatically adapts to data size
        model = IsolationForest(
            contamination=args.contamination,
            random_state=RANDOM_SEED
        )
        model.fit(X_clean)
        
        # Score data (one score per row) - O(n)
        if args.monte_carlo_samples > 0:
            anomaly_scores = monte_carlo_score_samples(X_clean, args.contamination, args.monte_carlo_samples)
        else:
            anomaly_scores = model.score_samples(X_clean)
        
        anomaly_scores = _apply_score_transforms(anomaly_scores, args.ascending, args.epsilon)
        binary_predictions = _scores_to_binary(anomaly_scores, args.contamination, args.ascending)
        
        # Write output (cache sheet names for O(1) lookup)
        if args.output_sheet not in {s.name for s in wb.sheets}:
            out_sht = wb.sheets.add(args.output_sheet, after=wb.sheets[-1])
        else:
            out_sht = wb.sheets[args.output_sheet]
            out_sht.clear()
        
        # Build data rows efficiently (vectorized where possible)
        # Pre-compute anomaly labels from binary predictions
        anomaly_labels = np.where(binary_predictions == 1, 'Anomaly', 'Normal')
        data_rows = []
        valid_idx = 0
        for orig_idx in range(len(df)):
            if valid_mask[orig_idx]:
                data_rows.append([orig_idx + 1] + X_clean[valid_idx].tolist() + [
                    anomaly_scores[valid_idx], 
                    anomaly_labels[valid_idx]
                ])
                valid_idx += 1
        
        out_sht.range('A1').value = [['Index'] + list(numeric_cols) + ['Anomaly Score', 'Is Anomaly']]
        out_sht.range('A2').value = data_rows
        
        wb.save()
        
        n_anomalies = np.sum(binary_predictions)
        print(f"\n{'='*50}")
        print(f"Anomaly Detection Complete!")
        print(f"Features: {n_features}")
        print(f"Samples: {n_samples_clean}")
        print(f"Anomalies: {n_anomalies} ({n_anomalies/n_samples_clean:.2%})")
        print(f"See '{args.output_sheet}' sheet for details.")
        
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())

