from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


DEFAULT_BENFORD_FEATURES = (
    "Flow Duration",
    "Flow IAT Mean",
    "Fwd IAT Mean",
    "Bwd IAT Mean",
    "Total Length of Fwd Packets",
    "Total Length of Bwd Packets",
    "Packet Length Mean",
    "Average Packet Size",
    "total_bytes",
    "total_packets",
)


@dataclass(slots=True)
class PipelineConfig:
    dataset_dir: Path = Path("data/raw/cicids2017/GeneratedLabelledFlows/TrafficLabelling")
    output_dir: Path = Path("outputs/cicids2017")
    window_minutes: int = 15
    min_values_per_benford_test: int = 50
    benford_features: tuple[str, ...] = DEFAULT_BENFORD_FEATURES
    isolation_contamination: float = 0.05
    alert_quantile: float = 0.95
    random_state: int = 42
    max_rows_per_file: int | None = None
    ewma_alpha: float = 0.30
    cusum_k: float = 0.50
    final_score_weights: dict[str, float] = field(
        default_factory=lambda: {
            "stat": 0.35,
            "temporal": 0.20,
            "iforest": 0.25,
            "graph": 0.20,
        }
    )

    @property
    def window_rule(self) -> str:
        return f"{self.window_minutes}min"

    def resolve_paths(self) -> "PipelineConfig":
        self.dataset_dir = self.dataset_dir.resolve()
        self.output_dir = self.output_dir.resolve()
        return self
