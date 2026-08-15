# HBG-NIDS: Hybrid Benford-Graph Network Intrusion Detection

Reproducibility code for the paper *"Robust Unsupervised Anomaly Detection in
Encrypted Network Flows via Multi-Layer Score Fusion and Adaptive Extreme Value
Thresholding."* The framework fuses Benford's Law divergence, EWMA/CUSUM temporal
drift, and Isolation Forest into a single unsupervised detection score, with a
post-hoc graph module for victim-host identification.

This repository intentionally contains only the source code behind the paper's
results — no drafts, no build scripts, no generated outputs.

---

## Repository Structure

```
HBG-NIDS-TNSM/
├── src/hybrid_nids/          ← Importable package: the full detection pipeline
│   ├── __init__.py           ← Exposes HybridBenfordPipeline and PipelineConfig
│   ├── config.py             ← PipelineConfig: window size, feature list, weights
│   ├── data.py                ← CICFlowMeter CSV loading and windowing
│   ├── benford.py             ← Benford's Law divergence metrics (MAD, KS, chi-square, entropy)
│   ├── temporal.py            ← EWMA and CUSUM temporal drift statistics
│   ├── modeling.py            ← Isolation Forest fitting and scoring
│   ├── graphing.py            ← Per-window host/edge graph feature extraction (PageRank, betweenness)
│   ├── pipeline.py            ← HybridBenfordPipeline: orchestrates all layers, fuses scores, evaluates
│   └── visualization.py       ← Post-hoc graph rendering (NetworkX/PyVis/Gravis) and ranked-alert charts
├── scripts/                   ← Standalone modules verifying individual paper findings
│   ├── rolling_baseline.py    ← WeekdayBaselineModel + simulation for Finding F2 (baseline drift)
│   ├── evt_thresholding.py    ← SPOT/DSPOT adaptive thresholding for Finding F6 (threshold fragility)
│   └── feature_decorrelation.py ← PCA-based Benford decorrelation for Finding F1 (feature redundancy)
├── LICENSE
└── README.md
```

---

## What Each Module Does

**`src/hybrid_nids/` (the detection pipeline)**

- `config.py` — `PipelineConfig` dataclass: dataset paths, 15-minute window rule,
  the 10 candidate Benford features, Isolation Forest contamination, and the
  layer fusion weights.
- `data.py` — Reads raw CICFlowMeter CSVs, coerces types, derives `total_bytes`/
  `total_packets`/`is_attack`, and floors timestamps into fixed-width windows.
- `benford.py` — Computes first- and second-digit Benford conformity per window
  (MAD, Kolmogorov–Smirnov, chi-square, Euclidean distance, Shannon entropy gap).
- `temporal.py` — EWMA and two-sided CUSUM statistics over the windowed Benford
  divergence score, used to flag temporal drift.
- `modeling.py` — Fits an `IsolationForest` on the benign training partition and
  scores every window.
- `graphing.py` — Builds a per-window directed host graph from flow records and
  extracts topology features (density, PageRank, betweenness, top host/edge).
- `pipeline.py` — `HybridBenfordPipeline`: runs the full five-phase procedure
  (feature screening → baseline training → weighted score fusion → thresholded
  alerting → per-window CSV/JSON/figure export), and computes precision/recall/
  F1/ROC-AUC against window-level labels.
- `visualization.py` — Loads pipeline outputs and renders the post-hoc graph
  evidence (static NetworkX PNG/SVG, interactive PyVis/Gravis HTML) plus ranked
  top-alerted-host/edge bar charts.

**`scripts/` (finding-verification modules)**

- `rolling_baseline.py` — Implements the `WeekdayBaselineModel` described in the
  paper and a Monday-vs-weekday-adaptive drift simulation, reproducing the
  Thursday false-positive reduction reported for Finding F2.
- `evt_thresholding.py` — Implements `SPOTDetector` and `DSPOTDetector`
  (Streaming/Drift Peak-Over-Threshold, Generalized Pareto Distribution fitting)
  used to derive the adaptive alert threshold discussed under Finding F6.
- `feature_decorrelation.py` — Implements `DecorrelatedBenfordDetector`, a
  PCA-based projection that decorrelates the Benford-conformant feature set,
  reproducing the correlation-matrix results behind Finding F1.

---

## Running the Pipeline

```python
from pathlib import Path
from hybrid_nids import HybridBenfordPipeline, PipelineConfig

config = PipelineConfig(
    dataset_dir=Path("data/raw/cicids2017/GeneratedLabelledFlows/TrafficLabelling"),
    output_dir=Path("outputs/cicids2017"),
)
pipeline = HybridBenfordPipeline(config)
results = pipeline.run()
print(results["summary"])
```

This expects the CICIDS2017 CSVs (from the Canadian Institute for Cybersecurity:
https://www.unb.ca/cic/datasets/ids-2017.html) laid out under `dataset_dir`.
`pipeline.run()` writes `window_scores.csv`, `alerts.csv`, `host_scores.csv`,
`edge_scores.csv`, `summary.json`, and a final-score time series figure to
`output_dir`.

To render the post-hoc graph evidence and ranked-alert charts after a run:

```python
from pathlib import Path
from hybrid_nids.visualization import render_visualization_bundle

render_visualization_bundle(output_dir=Path("outputs/cicids2017"))
```

To reproduce the standalone finding simulations:

```bash
python scripts/rolling_baseline.py        # Finding F2
python scripts/evt_thresholding.py        # Finding F6 (import SPOTDetector/DSPOTDetector directly)
python scripts/feature_decorrelation.py   # Finding F1
```

---

## Requirements

```
numpy
pandas
scikit-learn
scipy
networkx
matplotlib
pyvis
gravis
```

Python 3.10+ recommended (uses `from __future__ import annotations` and
`dataclass(slots=True)`).

---

## Dataset

CICIDS2017 is available from the Canadian Institute for Cybersecurity:
https://www.unb.ca/cic/datasets/ids-2017.html — subject to their usage terms.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
