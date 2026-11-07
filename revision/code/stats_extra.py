"""Paired significance (detector vs detector+Benford) and precision under realistic attack prevalence.
Uses the saved seed-0 scores; thresholds = P95 / P99 of benign validation scores (as in the paper)."""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.stats import binomtest
from sklearn.metrics import f1_score

DETS = ["IF", "KitNET", "AE", "OCSVM", "LOF", "HBOS"]


def paired(y, pa, pb, B=1000, seed=0):
    rs = np.random.default_rng(seed)
    n = len(y)
    d = []
    for _ in range(B):
        i = rs.integers(0, n, n)
        d.append(f1_score(y[i], pb[i]) - f1_score(y[i], pa[i]))
    ca, cb = pa == y, pb == y
    b, c = int((ca & ~cb).sum()), int((~ca & cb).sum())
    p = float(binomtest(b, b + c, 0.5).pvalue) if b + c else 1.0
    return float(np.percentile(d, 2.5)), float(np.percentile(d, 97.5)), b, c, p


def main(res):
    z = np.load(Path(res) / "flow_scores_seed0.npz")
    y = z["yB"].astype(int)
    sig, prev = [], []
    for m in DETS:
        P = {}
        for k in (m, m + "+B"):
            sv, st = z["va__" + k], z["te__" + k]
            for q in (0.95, 0.99):
                P[(k, q)] = st >= np.quantile(sv, q)
        lo, hi, b, c, p = paired(y, P[(m, 0.95)], P[(m + "+B", 0.95)])
        sig.append(dict(detector=m, f1_alone=f1_score(y, P[(m, 0.95)]), f1_plus_B=f1_score(y, P[(m + "+B", 0.95)]),
                        delta_f1=f1_score(y, P[(m + "+B", 0.95)]) - f1_score(y, P[(m, 0.95)]), ci_low=lo, ci_high=hi,
                        only_alone_correct=b, only_plusB_correct=c, mcnemar_p=p))
        for k in (m, m + "+B"):
            for q in (0.95, 0.99):
                pr = P[(k, q)]
                tpr, fpr = pr[y == 1].mean(), pr[y == 0].mean()
                row = dict(method=k, threshold=f"P{int(q * 100)}", tpr=tpr, fpr=fpr)
                for pi in (0.01, 0.001, 0.0001):
                    row[f"precision_at_{pi:g}"] = tpr * pi / (tpr * pi + fpr * (1 - pi)) if tpr + fpr > 0 else np.nan
                prev.append(row)
    pd.DataFrame(sig).to_csv(Path(res) / "flow_significance.csv", index=False)
    pd.DataFrame(prev).to_csv(Path(res) / "flow_prevalence.csv", index=False)
    print(pd.DataFrame(sig).round(4).to_string())
    print(pd.DataFrame(prev).round(4).to_string())


if __name__ == "__main__":
    for r_ in sys.argv[1:]:
        print("==", r_)
        main(r_)
