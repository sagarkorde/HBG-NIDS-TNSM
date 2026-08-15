from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd
from sklearn.metrics import f1_score, precision_score, recall_score, roc_auc_score

from .benford import evaluate_benford, flatten_benford_metrics
from .config import PipelineConfig
from .data import list_dataset_files, read_flow_csv, slugify
from .graphing import GRAPH_FEATURE_COLUMNS, extract_window_graph_artifacts
from .modeling import fit_isolation_forest, score_isolation_forest
from .temporal import compute_cusum, compute_ewma


EPSILON = 1e-9
STAT_METRIC_SUFFIXES = ("mad", "ks", "chi_square", "euclidean", "entropy_gap")

matplotlib.use("Agg")
import matplotlib.pyplot as plt


def positive_zscore(values: pd.Series | pd.DataFrame, mean: pd.Series | float, std: pd.Series | float) -> pd.Series | pd.DataFrame:
    safe_std = std.copy() if isinstance(std, pd.Series) else std
    if isinstance(safe_std, pd.Series):
        safe_std = safe_std.replace(0, np.nan).fillna(1.0)
    elif safe_std == 0:
        safe_std = 1.0
    z_values = (values - mean) / safe_std
    return z_values.clip(lower=0)


class HybridBenfordPipeline:
    def __init__(self, config: PipelineConfig) -> None:
        self.config = config.resolve_paths()
        self.config.output_dir.mkdir(parents=True, exist_ok=True)

    def run(self) -> dict[str, object]:
        window_df, host_df, edge_df = self._build_window_feature_table()
        training_mask = window_df["is_training_window"]
        train_df = window_df[training_mask].copy()

        stat_feature_columns = self._select_stat_feature_columns(window_df)
        graph_feature_columns = [
            column for column in GRAPH_FEATURE_COLUMNS if train_df[column].notna().any()
        ]

        train_stat_means = train_df[stat_feature_columns].mean()
        train_stat_stds = train_df[stat_feature_columns].std(ddof=0).replace(0, np.nan).fillna(1.0)
        stat_matrix = positive_zscore(window_df[stat_feature_columns], train_stat_means, train_stat_stds)
        window_df["stat_score"] = stat_matrix.mean(axis=1, skipna=True).fillna(0.0)

        train_df = window_df[training_mask].copy()
        stat_mean = float(train_df["stat_score"].mean())
        stat_std = float(train_df["stat_score"].std(ddof=0) or 1.0)
        window_df["ewma_score"] = compute_ewma(window_df["stat_score"], alpha=self.config.ewma_alpha)
        cusum_df = compute_cusum(
            window_df["stat_score"],
            mean=stat_mean,
            std=stat_std,
            k=self.config.cusum_k,
        )
        window_df = pd.concat([window_df, cusum_df], axis=1)
        window_df["temporal_score"] = (
            positive_zscore(
                window_df["ewma_score"],
                window_df.loc[training_mask, "ewma_score"].mean(),
                window_df.loc[training_mask, "ewma_score"].std(ddof=0) or 1.0,
            )
            + positive_zscore(
                window_df["cusum_score"],
                window_df.loc[training_mask, "cusum_score"].mean(),
                window_df.loc[training_mask, "cusum_score"].std(ddof=0) or 1.0,
            )
        ) / 2.0

        iforest_artifact = fit_isolation_forest(
            train_df=train_df,
            feature_columns=stat_feature_columns,
            contamination=self.config.isolation_contamination,
            random_state=self.config.random_state,
        )
        window_df["iforest_score"] = score_isolation_forest(iforest_artifact, window_df)

        graph_means = train_df[graph_feature_columns].mean()
        graph_stds = train_df[graph_feature_columns].std(ddof=0).replace(0, np.nan).fillna(1.0)
        graph_matrix = positive_zscore(window_df[graph_feature_columns], graph_means, graph_stds)
        window_df["graph_score"] = graph_matrix.mean(axis=1, skipna=True).fillna(0.0)

        train_df = window_df[training_mask].copy()
        train_if_mean = float(train_df["iforest_score"].mean())
        train_if_std = float(train_df["iforest_score"].std(ddof=0) or 1.0)
        train_temporal_mean = float(train_df["temporal_score"].mean())
        train_temporal_std = float(train_df["temporal_score"].std(ddof=0) or 1.0)
        train_graph_mean = float(train_df["graph_score"].mean())
        train_graph_std = float(train_df["graph_score"].std(ddof=0) or 1.0)

        window_df["stat_component"] = positive_zscore(window_df["stat_score"], stat_mean, stat_std)
        window_df["temporal_component"] = positive_zscore(
            window_df["temporal_score"], train_temporal_mean, train_temporal_std
        )
        window_df["iforest_component"] = positive_zscore(
            window_df["iforest_score"], train_if_mean, train_if_std
        )
        window_df["graph_component"] = positive_zscore(
            window_df["graph_score"], train_graph_mean, train_graph_std
        )
        window_df["dominant_component"] = (
            window_df[["stat_component", "temporal_component", "iforest_component", "graph_component"]]
            .idxmax(axis=1)
            .str.replace("_component", "", regex=False)
        )

        weights = self.config.final_score_weights
        window_df["final_score"] = (
            weights["stat"] * window_df["stat_component"]
            + weights["temporal"] * window_df["temporal_component"]
            + weights["iforest"] * window_df["iforest_component"]
            + weights["graph"] * window_df["graph_component"]
        )

        train_df = window_df[training_mask].copy()
        threshold = float(train_df["final_score"].quantile(self.config.alert_quantile))
        window_df["alert"] = window_df["final_score"] >= threshold
        window_df["true_attack_window"] = window_df["attack_flow_fraction"] > 0
        window_df["alert_context"] = (
            "dominant="
            + window_df["dominant_component"].astype(str)
            + "; host="
            + window_df["graph_top_host"].astype(str)
            + "; edge="
            + window_df["graph_top_edge"].astype(str)
        )

        merge_columns = [
            "source_file",
            "window_start",
            "alert",
            "final_score",
            "graph_component",
            "dominant_component",
            "alert_context",
            "window_label",
            "true_attack_window",
        ]
        if not host_df.empty:
            host_df = host_df.merge(window_df[merge_columns], on=["source_file", "window_start"], how="left")
        if not edge_df.empty:
            edge_df = edge_df.merge(window_df[merge_columns], on=["source_file", "window_start"], how="left")

        alerts_df = window_df[window_df["alert"]].copy()
        summary = self._evaluate(window_df, host_df, edge_df, threshold)

        self._save_outputs(window_df, alerts_df, host_df, edge_df, summary)
        return {
            "window_scores": window_df,
            "alerts": alerts_df,
            "host_scores": host_df,
            "edge_scores": edge_df,
            "summary": summary,
            "config": asdict(self.config),
        }

    def _build_window_feature_table(self) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        records: list[dict[str, object]] = []
        host_records: list[dict[str, object]] = []
        edge_records: list[dict[str, object]] = []
        dataset_files = list_dataset_files(self.config.dataset_dir)
        if not dataset_files:
            raise FileNotFoundError(f"No CSV files found in {self.config.dataset_dir}")

        for csv_path in dataset_files:
            day_df = read_flow_csv(csv_path, self.config)
            grouped = day_df.groupby("window_start", sort=True)
            for window_start, window_df in grouped:
                record: dict[str, object] = {
                    "source_file": csv_path.name,
                    "window_start": window_start,
                    "flow_count": int(len(window_df)),
                    "attack_flow_fraction": float(window_df["is_attack"].mean()),
                    "attack_label_count": int(window_df["is_attack"].sum()),
                    "distinct_sources": int(window_df["Source IP"].nunique()),
                    "distinct_destinations": int(window_df["Destination IP"].nunique()),
                    "window_label": "ATTACK" if window_df["is_attack"].any() else "BENIGN",
                    "is_training_window": "monday" in csv_path.name.lower(),
                    "timestamp_min": window_df["Timestamp"].min(),
                    "timestamp_max": window_df["Timestamp"].max(),
                }

                for feature in self.config.benford_features:
                    prefix = slugify(feature)
                    values = window_df[feature].to_numpy(dtype=float, copy=False)
                    for digits in (1, 2):
                        metrics = evaluate_benford(
                            values,
                            n_digits=digits,
                            min_count=self.config.min_values_per_benford_test,
                        )
                        record.update(flatten_benford_metrics(f"{prefix}_f{digits}", metrics))

                graph_artifacts = extract_window_graph_artifacts(
                    window_df=window_df,
                    source_file=csv_path.name,
                    window_start=window_start,
                )
                record.update(graph_artifacts.window_metrics)
                records.append(record)
                host_records.extend(graph_artifacts.host_rows)
                edge_records.extend(graph_artifacts.edge_rows)

        result = pd.DataFrame.from_records(records).sort_values("window_start").reset_index(drop=True)
        host_df = pd.DataFrame.from_records(host_records).sort_values(
            ["window_start", "node_score"], ascending=[True, False]
        ).reset_index(drop=True) if host_records else pd.DataFrame()
        edge_df = pd.DataFrame.from_records(edge_records).sort_values(
            ["window_start", "edge_score"], ascending=[True, False]
        ).reset_index(drop=True) if edge_records else pd.DataFrame()
        return result, host_df, edge_df

    def _select_stat_feature_columns(self, window_df: pd.DataFrame) -> list[str]:
        candidates = [
            column
            for column in window_df.columns
            if column.endswith(STAT_METRIC_SUFFIXES)
        ]
        return [column for column in candidates if window_df[column].notna().any()]

    def _evaluate(
        self,
        window_df: pd.DataFrame,
        host_df: pd.DataFrame,
        edge_df: pd.DataFrame,
        threshold: float,
    ) -> dict[str, object]:
        y_true = window_df["true_attack_window"].astype(int)
        y_pred = window_df["alert"].astype(int)
        summary: dict[str, object] = {
            "threshold": threshold,
            "window_count": int(len(window_df)),
            "alert_count": int(y_pred.sum()),
            "attack_window_count": int(y_true.sum()),
            "precision": float(precision_score(y_true, y_pred, zero_division=0)),
            "recall": float(recall_score(y_true, y_pred, zero_division=0)),
            "f1": float(f1_score(y_true, y_pred, zero_division=0)),
        }
        try:
            summary["roc_auc"] = float(roc_auc_score(y_true, window_df["final_score"]))
        except ValueError:
            summary["roc_auc"] = None

        per_file = (
            window_df.groupby("source_file")
            .agg(
                windows=("window_start", "size"),
                alerts=("alert", "sum"),
                attack_windows=("true_attack_window", "sum"),
                avg_final_score=("final_score", "mean"),
            )
            .reset_index()
        )
        summary["per_file"] = per_file.to_dict(orient="records")

        if not host_df.empty:
            top_hosts = (
                host_df[host_df["alert"]]
                .groupby("host")
                .agg(
                    alert_windows=("window_start", "nunique"),
                    mean_node_score=("node_score", "mean"),
                    max_node_score=("node_score", "max"),
                )
                .sort_values(["alert_windows", "max_node_score"], ascending=[False, False])
                .head(10)
                .reset_index()
            )
            summary["top_alert_hosts"] = top_hosts.to_dict(orient="records")
        else:
            summary["top_alert_hosts"] = []

        if not edge_df.empty:
            top_edges = (
                edge_df[edge_df["alert"]]
                .groupby(["source_ip", "destination_ip"])
                .agg(
                    alert_windows=("window_start", "nunique"),
                    mean_edge_score=("edge_score", "mean"),
                    max_edge_score=("edge_score", "max"),
                )
                .sort_values(["alert_windows", "max_edge_score"], ascending=[False, False])
                .head(10)
                .reset_index()
            )
            summary["top_alert_edges"] = top_edges.to_dict(orient="records")
        else:
            summary["top_alert_edges"] = []
        return summary

    def _save_outputs(
        self,
        window_df: pd.DataFrame,
        alerts_df: pd.DataFrame,
        host_df: pd.DataFrame,
        edge_df: pd.DataFrame,
        summary: dict[str, object],
    ) -> None:
        window_csv = self.config.output_dir / "window_scores.csv"
        alerts_csv = self.config.output_dir / "alerts.csv"
        host_csv = self.config.output_dir / "host_scores.csv"
        edge_csv = self.config.output_dir / "edge_scores.csv"
        summary_json = self.config.output_dir / "summary.json"
        figure_path = self.config.output_dir / "final_scores.png"

        window_df.to_csv(window_csv, index=False)
        alerts_df.to_csv(alerts_csv, index=False)
        if not host_df.empty:
            host_df.to_csv(host_csv, index=False)
        if not edge_df.empty:
            edge_df.to_csv(edge_csv, index=False)
        summary_json.write_text(json.dumps(summary, indent=2), encoding="utf-8")

        plt.figure(figsize=(14, 6))
        plt.plot(window_df["window_start"], window_df["final_score"], marker="o", linewidth=1.5)
        plt.axhline(summary["threshold"], color="red", linestyle="--", label="Alert threshold")
        plt.title("Hybrid Benford NIDS Final Score Over Time")
        plt.xlabel("Window Start")
        plt.ylabel("Final Score")
        plt.xticks(rotation=45, ha="right")
        plt.legend()
        plt.tight_layout()
        plt.savefig(figure_path, dpi=180)
        plt.close()
