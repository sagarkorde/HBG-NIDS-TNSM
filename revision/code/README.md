# Revision experiments (Paper ID 1388)

Reproduces every number in `1388_revised_version`.

    python prepare_data.py                      # splits the corrected CNS2022 zips into parts
    python run.py --kind cic2017 --data <CICIDS2017_improved parts> --out results/cic2017 --minutes 1 15 --seeds 5
    python run.py --kind cic2018 --data <CSECICIDS2018_improved parts> --out results/cic2018 --minutes 1 15 --seeds 5 --train-days 1
    python thresholds.py results/cic2017 results/cic2018
    python explain.py cic2017 <data> results/cic2017   (same for cic2018)
    python figures.py

Requirements: numpy, pandas, scikit-learn, scipy, networkx, pyod, xgboost, shap, matplotlib.
Protocol: benign-only training on the first capture day (blocked 80/20 train/validation split), test = later days;
CSE-CIC-IDS-2018 uses a 5% uniform flow sample (seeded). Labels: corrected releases of Liu et al., CNS 2022; "attempted" flows -> benign.

Added after internal review (7 Nov):
    python stats_extra.py results/cic2017 results/cic2018       # paired McNemar + bootstrap ΔF1, precision at 1/0.1/0.01% prevalence
    python sampling.py <CICIDS2017 parts> results/sampling      # window-score sensitivity to uniform flow sampling
    python run.py --kind cic2017 --data <parts> --out results/sampling/flow_0.2 --minutes 0 --seeds 1 --only IF AE OCSVM KitNET --sample-frac 0.2
    python blockboot.py cic2017 <parts> results/cic2017      # window- and day-block bootstrap of ΔF1 (same for cic2018)
