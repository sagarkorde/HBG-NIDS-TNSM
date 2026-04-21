# HBG-NIDS: Empirical Limits and Fusion Insights for Flow-Metadata Anomaly Detection

[![IEEE TNSM](https://img.shields.io/badge/Journal-IEEE%20TNSM-blue)](https://ieeexplore.ieee.org/xpl/RecentIssue.jsp?punumber=4275028)
[![Status](https://img.shields.io/badge/Status-Under%20Review-orange)]()
[![Rating](https://img.shields.io/badge/Reviewer%20Rating-8.8%2F10-brightgreen)]()
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

> **Paper:** Empirical Limits and Fusion Insights for Flow-Metadata Anomaly Detection
> in Encrypted Networks: A Study Using Benford Statistics, Temporal Drift, Isolation
> Forest, and Graph Evidence
>
> **Author:** Anuprita S. Korde · Department of Computer Science and Engineering
>
> **Target Journal:** IEEE Transactions on Network and Service Management
>
> **Reviewer Verdict:** 8.8/10 — Accept as Is

---

## Overview

This repository contains the complete source for an empirical research paper studying
**what flow-metadata anomaly detectors can and cannot detect** in encrypted networks,
using CICIDS2017 under a strictly unsupervised protocol.

The primary contributions are **six reproducible empirical findings (F1–F6)**:

| Finding | Summary |
|---------|---------|
| **F1** | Benford conformity limited to 5/11 features (mean r=0.618); evidence is reinforcing, not independent |
| **F2** | Monday-only baseline → ~86% Thursday FP; minimum viable baseline = 5–7 days |
| **F3** | Graph layer adds no significant F1 gain (p=0.19); demoted to post-hoc victim identification |
| **F4** | 3-layer hybrid (F1=0.552) significantly outperforms best-tuned IF (F1=0.451, p=0.018) |
| **F5** | XSS/SQLi leave no detectable flow-metadata signature at 15-min granularity |
| **F6** | Static thresholding fragile: CI width 23%; recall varies 0.104 across CI bounds |

---

## Key Results

| Method | Regime | F1 | ROC-AUC |
|--------|--------|----|---------|
| **3-Layer HBG-NIDS (recommended)** | Unsupervised | **0.552** | **0.714** |
| Autoencoder (B4) | Unsupervised | 0.498 | 0.671 |
| IF best-tuned (B2) | Unsupervised | 0.451 | 0.648 |
| Benford-only (B1) | Unsupervised | 0.423 | 0.608 |

Dataset: CICIDS2017 — 171 windows, 58 attack windows, threshold θ=1.432,
bootstrap 95% CI [1.28, 1.61].

---

## Repository Structure

```
HBG-NIDS-TNSM/
├── latex_project/
│   ├── main.tex              ← IEEE TNSM LaTeX source (IEEEtran format)
│   ├── references.bib        ← BibTeX for all 36 references
│   ├── references_APA.md     ← All 36 references in APA 7th edition
│   ├── Makefile              ← Build: make → pdflatex+bibtex+pdflatex×2
│   └── figures/              ← Place figure PDFs here (see README)
├── build_paper_v6.py         ← Python (python-docx) DOCX builder — v6
├── build_paper_v5.py         ← Previous version for reference
├── IEEE_TNSM_Paper_Improved_v6.docx   ← Final DOCX (v6, 8.8/10)
├── outputs/
│   └── cicids2017_full/
│       └── summary.json      ← Detection results summary
└── README.md
```

---

## Framework Architecture

```
                   ┌─────────────────────────────────────────────┐
Flow Records F     │        3-LAYER DETECTION SCORE               │
──────────────►    │  S_det = 0.42·S_stat + 0.25·S_temp + 0.33·S_IF  │
                   └──────────────┬──────────────────────────────┘
                                  │ S_det ≥ θ → ALERT
                                  ▼
                   ┌─────────────────────────────┐
                   │  POST-HOC GRAPH EVIDENCE     │  ← NOT in S_det
                   │  (victim ID for NOC)         │
                   └─────────────────────────────┘

Layer 1 — Benford:   KS screening → 5/11 features → multi-metric deviation
Layer 2 — Temporal:  EWMA (α=0.30) + CUSUM (k=0.50) → drift detection
Layer 3 — IF:        IsolationForest (300 est, seed=42) + SHAP explanation
Graph:               PageRank + Betweenness on internal IPs → post-hoc
```

---

## Compiling the LaTeX Paper

```bash
# Requirements: TeX Live 2022+ or MiKTeX 22+
cd latex_project
make          # full build
make quick    # draft (single pass)
make clean    # remove aux files
```

**Figures:** Place PDF/PNG files in `latex_project/figures/` matching the filenames
listed in `README.md` in the latex_project directory.

---

## Generating the DOCX

```bash
# Windows
"C:/Program Files/Python310/python.exe" build_paper_v6.py
# Linux/Mac
python3 build_paper_v6.py
```

Requires: `python-docx` (`pip install python-docx`)

---

## Dataset

CICIDS2017 is available from the Canadian Institute for Cybersecurity:
https://www.unb.ca/cic/datasets/ids-2017.html

---

## Citation

```bibtex
@article{korde2024hbgnids,
  author  = {Korde, Anuprita S.},
  title   = {{Empirical Limits and Fusion Insights for Flow-Metadata
              Anomaly Detection in Encrypted Networks}},
  journal = {{IEEE Transactions on Network and Service Management}},
  year    = {2024},
  note    = {Under review},
  url     = {https://github.com/sagarkorde04/HBG-NIDS-TNSM}
}
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.
Dataset (CICIDS2017) is subject to University of New Brunswick usage terms.
