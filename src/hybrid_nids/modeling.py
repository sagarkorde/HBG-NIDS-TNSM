from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest


@dataclass(slots=True)
class IsolationForestArtifact:
    model: IsolationForest
    feature_columns: list[str]
    fill_values: pd.Series


def fit_isolation_forest(
    train_df: pd.DataFrame,
    feature_columns: list[str],
    contamination: float,
    random_state: int,
) -> IsolationForestArtifact:
    usable_columns = [col for col in feature_columns if train_df[col].notna().any()]
    X = train_df[usable_columns].replace([np.inf, -np.inf], np.nan)
    fill_values = X.median()
    X = X.fillna(fill_values)

    model = IsolationForest(
        n_estimators=300,
        contamination=contamination,
        random_state=random_state,
        n_jobs=-1,
    )
    model.fit(X)
    return IsolationForestArtifact(
        model=model,
        feature_columns=usable_columns,
        fill_values=fill_values,
    )


def score_isolation_forest(
    artifact: IsolationForestArtifact,
    df: pd.DataFrame,
) -> pd.Series:
    X = df[artifact.feature_columns].replace([np.inf, -np.inf], np.nan)
    X = X.fillna(artifact.fill_values)
    return pd.Series(-artifact.model.score_samples(X), index=df.index, name="iforest_raw_score")
