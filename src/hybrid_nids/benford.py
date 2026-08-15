from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(slots=True)
class BenfordMetrics:
    count: int
    mad: float
    ks: float
    chi_square: float
    euclidean: float
    entropy: float
    entropy_gap: float


def benford_probabilities(n_digits: int) -> np.ndarray:
    start = 10 ** (n_digits - 1)
    stop = 10**n_digits
    digits = np.arange(start, stop, dtype=float)
    return np.log10(1.0 + 1.0 / digits)


def significant_digits(values: np.ndarray, n_digits: int) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    arr = np.abs(arr)
    arr = arr[np.isfinite(arr) & (arr > 0)]
    if arr.size == 0:
        return np.array([], dtype=int)

    magnitudes = np.floor(np.log10(arr))
    scaled = arr / np.power(10.0, magnitudes)
    digits = np.floor(scaled * (10 ** (n_digits - 1))).astype(int)

    lower = 10 ** (n_digits - 1)
    upper = 10**n_digits
    mask = (digits >= lower) & (digits < upper)
    return digits[mask]


def observed_distribution(digits: np.ndarray, n_digits: int) -> np.ndarray:
    lower = 10 ** (n_digits - 1)
    upper = 10**n_digits
    counts = np.bincount(digits - lower, minlength=upper - lower).astype(float)
    total = counts.sum()
    if total == 0:
        return counts
    return counts / total


def shannon_entropy(probabilities: np.ndarray) -> float:
    probs = np.asarray(probabilities, dtype=float)
    probs = probs[probs > 0]
    if probs.size == 0:
        return 0.0
    return float(-(probs * np.log2(probs)).sum())


def evaluate_benford(
    values: np.ndarray,
    n_digits: int,
    min_count: int,
) -> BenfordMetrics:
    digits = significant_digits(values, n_digits=n_digits)
    if digits.size < min_count:
        return BenfordMetrics(
            count=int(digits.size),
            mad=np.nan,
            ks=np.nan,
            chi_square=np.nan,
            euclidean=np.nan,
            entropy=np.nan,
            entropy_gap=np.nan,
        )

    observed = observed_distribution(digits, n_digits=n_digits)
    expected = benford_probabilities(n_digits=n_digits)
    observed_cdf = np.cumsum(observed)
    expected_cdf = np.cumsum(expected)
    expected_counts = expected * digits.size
    observed_counts = observed * digits.size

    with np.errstate(divide="ignore", invalid="ignore"):
        chi_square = np.nansum((observed_counts - expected_counts) ** 2 / expected_counts)

    return BenfordMetrics(
        count=int(digits.size),
        mad=float(np.mean(np.abs(observed - expected))),
        ks=float(np.max(np.abs(observed_cdf - expected_cdf))),
        chi_square=float(chi_square),
        euclidean=float(np.linalg.norm(observed - expected)),
        entropy=shannon_entropy(observed),
        entropy_gap=abs(shannon_entropy(observed) - shannon_entropy(expected)),
    )


def flatten_benford_metrics(prefix: str, metrics: BenfordMetrics) -> dict[str, float]:
    return {
        f"{prefix}_count": metrics.count,
        f"{prefix}_mad": metrics.mad,
        f"{prefix}_ks": metrics.ks,
        f"{prefix}_chi_square": metrics.chi_square,
        f"{prefix}_euclidean": metrics.euclidean,
        f"{prefix}_entropy": metrics.entropy,
        f"{prefix}_entropy_gap": metrics.entropy_gap,
    }
