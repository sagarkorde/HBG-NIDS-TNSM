"""Cluster (block) bootstrap of ΔF1 = F1(D+B) − F1(D): resample whole 1-min windows, and whole capture days.
Rebuilds the evaluation-half flow indices exactly as run.flow_experiment (verified against the saved labels)."""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

sys.path.insert(0, str(Path(__file__).parent))
from loaders import load_dir  # noqa
from run import split_masks  # noqa
from windows import build_windows  # noqa

DETS = ["IF", "KitNET", "AE", "OCSVM", "LOF", "HBOS"]


def eval_index(kind, data, train_days=1, n_test=300_000, n_train=200_000, n_val=50_000):
    df = load_dir(Path(data), kind)
    W, codes = build_windows(df, 1, with_graph=False)
    trw, vaw, tew = split_masks(W, kind, train_days)
    rs = np.random.default_rng(0)
    fi_tr = np.flatnonzero(np.isin(codes, np.flatnonzero(trw)) & (df.label.values == 0))
    fi_va = np.flatnonzero(np.isin(codes, np.flatnonzero(vaw)) & (df.label.values == 0))
    fi_te = np.flatnonzero(np.isin(codes, np.flatnonzero(tew)))
    rs.choice(fi_tr, min(n_train, len(fi_tr)), replace=False)  # consume RNG exactly as run.py
    rs.choice(fi_va, min(n_val, len(fi_va)), replace=False)
    te_df = df.iloc[fi_te]
    keep = []
    cap = n_test // max(te_df.attack.nunique(), 1)
    for a, g in te_df.groupby("attack").groups.items():
        g = np.asarray(g)
        keep.append(rs.choice(g, min(len(g), cap if a != "BENIGN" else n_test // 2), replace=False))
    fi_te = np.sort(np.concatenate(keep))
    yte = df.label.values[fi_te]
    A, B = train_test_split(np.arange(len(fi_te)), test_size=0.5, stratify=yte, random_state=0)
    ix = fi_te[B]
    return df.label.values[ix], codes[ix], df.dayfile.values[ix]


def counts(groups, y, p, G):
    tp = np.bincount(groups, weights=(y == 1) & p, minlength=G)
    fp = np.bincount(groups, weights=(y == 0) & p, minlength=G)
    fn = np.bincount(groups, weights=(y == 1) & ~p, minlength=G)
    return tp, fp, fn


def f1w(c, w):
    tp, fp, fn = (w @ x for x in c)
    return 2 * tp / np.maximum(2 * tp + fp + fn, 1e-12)


def main(kind, data, res, B=2000):
    y, win, day = eval_index(kind, data)
    z = np.load(Path(res) / "flow_scores_seed0.npz")
    assert np.array_equal(y.astype(int), z["yB"].astype(int)), "index reconstruction mismatch"
    rows = []
    rs = np.random.default_rng(0)
    for level, g in [("window", win), ("day", day)]:
        gid, gidx = np.unique(g, return_inverse=True)
        G = len(gid)
        Wts = rs.multinomial(G, np.ones(G) / G, size=B).astype(float)  # bootstrap weights over clusters
        for m in DETS:
            pa = z["te__" + m] >= np.quantile(z["va__" + m], 0.95)
            pb = z["te__" + m + "+B"] >= np.quantile(z["va__" + m + "+B"], 0.95)
            ca, cb = counts(gidx, y, pa, G), counts(gidx, y, pb, G)
            d = f1w(cb, Wts) - f1w(ca, Wts)
            point = f1w(cb, np.ones((1, G)))[0] - f1w(ca, np.ones((1, G)))[0]
            rows.append(dict(level=level, clusters=G, detector=m, delta_f1=point, ci_low=np.percentile(d, 2.5),
                             ci_high=np.percentile(d, 97.5), share_le_0=float((d <= 0).mean()) if point > 0 else float((d >= 0).mean())))
    out = pd.DataFrame(rows)
    out.to_csv(Path(res) / "flow_block_bootstrap.csv", index=False)
    print(out.round(4).to_string())


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
