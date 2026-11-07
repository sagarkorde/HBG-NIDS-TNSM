"""Metrics, bootstrap confidence intervals and exact McNemar tests."""
import numpy as np
from scipy.stats import binomtest
from sklearn.metrics import average_precision_score, f1_score, precision_score, recall_score, roc_auc_score


def metrics(y, score, pred):
    y, pred = np.asarray(y).astype(int), np.asarray(pred).astype(int)
    out = dict(precision=precision_score(y, pred, zero_division=0), recall=recall_score(y, pred, zero_division=0),
               f1=f1_score(y, pred, zero_division=0), fpr=float(pred[y == 0].mean()) if (y == 0).any() else np.nan,
               tp=int(((y == 1) & (pred == 1)).sum()), fp=int(((y == 0) & (pred == 1)).sum()),
               fn=int(((y == 1) & (pred == 0)).sum()), tn=int(((y == 0) & (pred == 0)).sum()))
    try:
        out["roc_auc"] = roc_auc_score(y, score)
        out["pr_auc"] = average_precision_score(y, score)
    except ValueError:
        out["roc_auc"] = out["pr_auc"] = np.nan
    return out


def bootstrap_ci(y, score, pred, B=1000, seed=0, keys=("f1", "precision", "recall", "roc_auc")):
    rs = np.random.default_rng(seed)
    y, score, pred = map(np.asarray, (y, score, pred))
    n = len(y)
    vals = {k: [] for k in keys}
    for _ in range(B):
        i = rs.integers(0, n, n)
        if y[i].min() == y[i].max():
            continue
        m = metrics(y[i], score[i], pred[i])
        for k in keys:
            vals[k].append(m[k])
    return {k: (float(np.nanpercentile(v, 2.5)), float(np.nanpercentile(v, 97.5))) for k, v in vals.items()}


def mcnemar(y, pred_a, pred_b):
    y = np.asarray(y).astype(int)
    ca, cb = np.asarray(pred_a) == y, np.asarray(pred_b) == y
    b, c = int((ca & ~cb).sum()), int((~ca & cb).sum())
    nd = b + c
    if nd == 0:
        return dict(b=b, c=c, n_d=0, p=1.0)
    return dict(b=b, c=c, n_d=nd, p=float(binomtest(b, nd, 0.5).pvalue))
