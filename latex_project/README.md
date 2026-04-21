# HBG-NIDS: Empirical Limits and Fusion Insights for Flow-Metadata Anomaly Detection

**Paper:** Empirical Limits and Fusion Insights for Flow-Metadata Anomaly Detection
in Encrypted Networks: A Study Using Benford Statistics, Temporal Drift, Isolation
Forest, and Graph Evidence

**Author:** Anuprita S. Korde, Member, IEEE
**Affiliation:** Department of Computer Science and Engineering
**Target:** IEEE Transactions on Network and Service Management (TNSM)
**Status:** Minor Revision — v6 (8.8/10 reviewer rating)

---

## Repository Structure

```
HBG-NIDS-TNSM/
├── latex_project/
│   ├── main.tex              ← IEEE TNSM LaTeX source (IEEEtran format)
│   ├── references.bib        ← BibTeX for all 36 references
│   ├── references_APA.md     ← All references in APA 7th edition format
│   ├── Makefile              ← Build automation
│   └── figures/              ← Place figure files here (see below)
├── build_paper_v6.py         ← Python script to generate the DOCX (v6)
├── IEEE_TNSM_Paper_Improved_v6.docx   ← Final DOCX output
└── outputs/                  ← CICIDS2017 experimental results
    └── cicids2017_full/
        ├── summary.json
        └── visualizations/
```

---

## Compiling the LaTeX Paper

### Requirements
- **TeX distribution:** TeX Live 2022+ or MiKTeX 22+
  (both include `IEEEtran.cls` and all required packages)
- **Required packages:** `cite`, `amsmath`, `graphicx`, `booktabs`, `multirow`,
  `algorithm`, `algpseudocode`, `hyperref`, `microtype`, `balance`

### Build
```bash
cd latex_project
make          # full build: pdflatex → bibtex → pdflatex × 2
make quick    # single pass (fast draft)
make clean    # remove auxiliary files
```

### Manual Compilation (Windows / without make)
```
pdflatex main
bibtex main
pdflatex main
pdflatex main
```

### Figure Files
Place the following figure files in `latex_project/figures/`:

| Filename              | Description |
|-----------------------|-------------|
| `final_scores.pdf`    | Anomaly score distribution with threshold θ=1.432 |
| `confusion_matrix.pdf`| 3-layer confusion matrix (37 TP, 39 FP, 21 FN, 74 TN) |
| `per_file_alerts.pdf` | Per-file alerts vs. true attack windows |
| `avg_score_by_day.pdf`| Average 3-layer score per day |
| `top_alert_hosts.pdf` | Top hosts by alert frequency (post-hoc graph) |
| `top_alert_edges.pdf` | Top edges by alert frequency |
| `portscan_networkx.pdf`| Internal-IP graph for highest-scoring PortScan window |

Generate these from `outputs/cicids2017_full/` and convert PNG → PDF with:
```bash
for f in outputs/cicids2017_full/visualizations/*.png; do
  convert "$f" "latex_project/figures/$(basename ${f%.png}).pdf"
done
```

---

## Six Empirical Findings (F1–F6)

| Finding | Summary |
|---------|---------|
| **F1** | Benford conformity limited to 5/11 features; mean r=0.618; reinforcing, not independent |
| **F2** | Monday-only baseline → ~86% Thursday FP; minimum viable baseline = 5–7 days |
| **F3** | Graph layer adds no significant F1 gain (p=0.19); demoted to post-hoc victim ID |
| **F4** | 3-layer hybrid (F1=0.552) significantly outperforms best-tuned IF (0.451, p=0.018) |
| **F5** | XSS/SQLi leave no detectable flow-metadata signature at 15-min granularity |
| **F6** | Static thresholding fragile: CI width 23%; recall varies 0.104 across bounds |

---

## Key Results

| Method | Regime | F1 | ROC-AUC |
|--------|--------|----|---------|
| 3-Layer HBG-NIDS (recommended) | Unsupervised | **0.552** | **0.714** |
| Autoencoder (B4) | Unsupervised | 0.498 | 0.671 |
| IF best-tuned (B2) | Unsupervised | 0.451 | 0.648 |
| Benford-only (B1) | Unsupervised | 0.423 | 0.608 |

Dataset: CICIDS2017 — 171 windows, 58 attack, θ=1.432, bootstrap CI [1.28, 1.61].

---

## Generating the DOCX

```bash
/c/Program\ Files/Python310/python.exe build_paper_v6.py
```

Output: `IEEE_TNSM_Paper_Improved_v6.docx`

---

## Citation

```bibtex
@article{korde2024hbgnids,
  author  = {Korde, Anuprita S.},
  title   = {{Empirical Limits and Fusion Insights for Flow-Metadata
              Anomaly Detection in Encrypted Networks}},
  journal = {{IEEE Transactions on Network and Service Management}},
  year    = {2024},
  note    = {Under review}
}
```

---

## License

Code released under the MIT License. See `LICENSE` for details.
Dataset (CICIDS2017) is subject to the University of New Brunswick usage terms.
