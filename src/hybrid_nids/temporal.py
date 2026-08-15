from __future__ import annotations

import numpy as np
import pandas as pd


def compute_ewma(values: pd.Series, alpha: float) -> pd.Series:
    return values.ewm(alpha=alpha, adjust=False).mean()


def compute_cusum(values: pd.Series, mean: float, std: float, k: float) -> pd.DataFrame:
    safe_std = std if std > 1e-9 else 1.0
    z_scores = (values - mean) / safe_std
    pos = np.zeros(len(z_scores))
    neg = np.zeros(len(z_scores))

    for idx, value in enumerate(z_scores.to_numpy()):
        if idx == 0:
            pos[idx] = max(0.0, value - k)
            neg[idx] = min(0.0, value + k)
            continue
        pos[idx] = max(0.0, pos[idx - 1] + value - k)
        neg[idx] = min(0.0, neg[idx - 1] + value + k)

    return pd.DataFrame(
        {
            "cusum_z": z_scores,
            "cusum_pos": pos,
            "cusum_neg": neg,
            "cusum_score": np.maximum(pos, np.abs(neg)),
        },
        index=values.index,
    )
