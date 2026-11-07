"""Post-hoc operating-point analysis on saved seed-0 scores: empirical percentile vs SPOT (EVT) thresholds,
all fitted on benign validation scores only. Reports realised test FPR (calibration under drift) and F1."""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.stats import genpareto

sys.path.insert(0, str(Path(__file__).parent))
from stats import metrics  # noqa: E402


def spot_threshold(s, risk, init_q=0.98):
    """SPOT initial-batch threshold (Siffer et al., KDD 2017): GPD fitted to excesses over the init_q quantile."""
    s = np.asarray(s, float)
    t = np.quantile(s, init_q)
    peaks = s[s > t] - t
    n, nt = len(s), len(peaks)
    if nt < 10:
        return float(np.quantile(s, 1 - risk))
    g, _, sig = genpareto.fit(peaks, floc=0)
    r = risk * n / nt
    if abs(g) < 1e-6:
        return float(t - sig * np.log(r))
    return float(t + sig / g * (r ** (-g) - 1))


def analyse(res_dir: Path, methods):
    z = np.load(res_dir / "flow_scores_seed0.npz", allow_pickle=False)
    y = z["yB"]
    rows = []
    for m in methods:
        if "te__" + m not in z or "va__" + m not in z:
            continue
        sv, st = z["va__" + m], z["te__" + m]
        for name, th in [("P95", np.quantile(sv, 0.95)), ("P99", np.quantile(sv, 0.99)),
                         ("P99.9", np.quantile(sv, 0.999)),
                         ("SPOT(1e-2)", spot_threshold(sv, 1e-2)), ("SPOT(1e-3)", spot_threshold(sv, 1e-3))]:
            r = metrics(y, st, st >= th)
            rows.append(dict(method=m, threshold_rule=name, target_fpr=dict(P95=.05, P99=.01).get(name, .001 if "1e-3" in name or "99.9" in name else .01),
                             test_fpr=r["fpr"], precision=r["precision"], recall=r["recall"], f1=r["f1"]))
    return pd.DataFrame(rows)


if __name__ == "__main__":
    for ds in sys.argv[1:]:
        d = Path(ds)
        out = analyse(d, ["IF", "IF+B", "AE", "AE+B", "KitNET", "KitNET+B", "OCSVM", "OCSVM+B", "LOF", "LOF+B", "HBOS+B"])
        out.to_csv(d / "flow_thresholds.csv", index=False)
        print(d.name)
        print(out.round(3).to_string())
