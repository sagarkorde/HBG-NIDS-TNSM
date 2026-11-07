# Revised evaluation (ICMLDE 2026, Paper ID 1388)

Everything reported in the revised manuscript is produced by the scripts in `code/` and stored in `results/`.

```
revision/
├── code/        experiment harness (see code/README.md for the exact commands)
├── results/
│   ├── cic2017/     corrected CIC-IDS-2017 (full)
│   ├── cic2018/     corrected CSE-CIC-IDS-2018 (5% uniform flow sample, seed 0)
│   └── sampling/    sampling-sensitivity runs on CIC-IDS-2017 (50/20/10/5%)
└── figures/     figures used in the paper
```

Key result files per dataset:

| File | Paper element |
|---|---|
| `flow_all_runs.csv`, `flow_summary.csv` | Table 2 (flow-level P/R/F1/AUC, 5 seeds) |
| `flow_significance.csv`, `flow_block_bootstrap.csv` | Table 3 (paired ΔF1; flow-, window- and day-block bootstrap) |
| `window_{1,15}min_all_runs.csv` | Table 4 (window-level ablation) |
| `flow_dayheldout.csv` | Section 5.3 (supervised, day-held-out) |
| `explain_*.csv`, `flow_per_attack_recall.csv` | Section 5.4 / Fig. 2 (layer contributions, TreeSHAP, per-attack recall) |
| `flow_thresholds.csv`, `flow_prevalence.csv` | Table 5 (operating points, precision at 1/0.1/0.01% prevalence) |
| `../sampling/*` | Table 6 (sensitivity to flow sampling) |
| `benford_screening_*.csv` | Section 3.1 (Benford feature screening) |

Data: corrected CSVs from https://intrusion-detection.distrinet-research.be/CNS2022/Dataset_Download.html
("attempted" flows are relabelled benign). Raw data and score dumps are not included.
