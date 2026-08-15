from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd

from .config import PipelineConfig


def slugify(value: str) -> str:
    lowered = value.strip().lower()
    lowered = re.sub(r"[^a-z0-9]+", "_", lowered)
    return lowered.strip("_")


def list_dataset_files(dataset_dir: Path) -> list[Path]:
    files = sorted(dataset_dir.glob("*.csv"))
    return sorted(files, key=lambda path: ("monday" not in path.name.lower(), path.name.lower()))


def read_flow_csv(path: Path, config: PipelineConfig) -> pd.DataFrame:
    df = pd.read_csv(
        path,
        low_memory=False,
        skipinitialspace=True,
        encoding="latin1",
        na_values=["Infinity", "inf", "-inf", "NaN", "nan"],
        nrows=config.max_rows_per_file,
    )
    df.columns = [str(col).strip() for col in df.columns]
    df["Label"] = df["Label"].astype(str).str.strip()
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp", "Source IP", "Destination IP"]).copy()

    numeric_columns = [col for col in df.columns if col not in {"Flow ID", "Source IP", "Destination IP", "Label", "Timestamp"}]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    df["total_bytes"] = (
        df["Total Length of Fwd Packets"].fillna(0) + df["Total Length of Bwd Packets"].fillna(0)
    )
    df["total_packets"] = (
        df["Total Fwd Packets"].fillna(0) + df["Total Backward Packets"].fillna(0)
    )
    df["is_attack"] = df["Label"].str.upper().ne("BENIGN")
    df = df.sort_values("Timestamp").reset_index(drop=True)
    df["window_start"] = df["Timestamp"].dt.floor(config.window_rule)
    return df
