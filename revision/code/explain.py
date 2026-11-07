"""Per-layer explanation of IF+Benford alerts (flow level, seed 0), replicating run.flow_experiment's sampling.
Outputs: explain_layers.csv (mean layer contribution and dominant-layer share per attack class),
         explain_shap_if.csv (mean |TreeSHAP| of top flow features for IF-dominant alerts),
         explain_benford.csv (Benford per-feature contribution for Benford-dominant alerts)."""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import shap
from sklearn.ensemble import IsolationForest
from sklearn.model_selection import train_test_split

sys.path.insert(0, str(Path(__file__).parent))
from detectors import FBTI, screen_benford  # noqa
from loaders import load_dir  # noqa
from run import split_masks, sym_log, threshold  # noqa
from windows import BENFORD_CANDIDATES, build_windows  # noqa


def main(kind, data, out, train_days=1, q=0.95, n_train=200_000, n_val=50_000, n_test=300_000):
    out = Path(out)
    df = load_dir(Path(data), kind)
    W, codes = build_windows(df, 1, with_graph=False)
    trw, vaw, tew = split_masks(W, kind, train_days)
    feats = screen_benford(df[np.isin(codes, np.flatnonzero(trw))], BENFORD_CANDIDATES)
    feats = feats.loc[feats.passed, "feature"].tolist() or ["duration", "total_bytes", "flow_bytes_s"]
    rs = np.random.default_rng(0)
    fi_tr = np.flatnonzero(np.isin(codes, np.flatnonzero(trw)) & (df.label.values == 0))
    fi_va = np.flatnonzero(np.isin(codes, np.flatnonzero(vaw)) & (df.label.values == 0))
    fi_te = np.flatnonzero(np.isin(codes, np.flatnonzero(tew)))
    fi_tr = np.sort(rs.choice(fi_tr, min(n_train, len(fi_tr)), replace=False))
    fi_va = np.sort(rs.choice(fi_va, min(n_val, len(fi_va)), replace=False))
    te_df = df.iloc[fi_te]
    keep = []
    cap = n_test // max(te_df.attack.nunique(), 1)
    for a, g in te_df.groupby("attack").groups.items():
        g = np.asarray(g)
        keep.append(rs.choice(g, min(len(g), cap if a != "BENIGN" else n_test // 2), replace=False))
    fi_te = np.sort(np.concatenate(keep))
    fcols = [c for c in df.columns if c.startswith("f_")]
    allix = np.concatenate([fi_tr, fi_va, fi_te])
    Xs = np.nan_to_num(sym_log(df[fcols].values[allix].astype(np.float64)), nan=0.0, posinf=0.0, neginf=0.0)
    keepc = Xs[: len(fi_tr)].std(0) > 0
    Xs = Xs[:, keepc]; names = [c[2:] for c, k in zip(fcols, keepc) if k]
    Xtr, Xva, Xte = Xs[: len(fi_tr)], Xs[len(fi_tr): len(fi_tr) + len(fi_va)], Xs[len(fi_tr) + len(fi_va):]
    yte = df.label.values[fi_te]
    A, B = train_test_split(np.arange(len(fi_te)), test_size=0.5, stratify=yte, random_state=0)
    f0 = FBTI(feats, seed=0).fit(W, trw)
    Lw = f0.layers(W)
    ifm = IsolationForest(n_estimators=300, random_state=0, n_jobs=-1).fit(Xtr)
    sel = np.random.default_rng(0).choice(len(fi_tr), min(len(fi_tr), 200_000), replace=False)
    s_tr = -ifm.score_samples(Xtr[sel[:20000]])
    mu, sd = s_tr.mean(), s_tr.std() or 1
    Lv = pd.DataFrame({"B": Lw.B.values[codes[fi_va]], "D": np.clip((-ifm.score_samples(Xva) - mu) / sd, 0, None)})
    Lt = pd.DataFrame({"B": Lw.B.values[codes[fi_te[B]]], "D": np.clip((-ifm.score_samples(Xte[B]) - mu) / sd, 0, None)})
    w = f0.optimise_weights(Lv, q, ("B", "D"))
    sv, st = f0.fuse(Lv, w), f0.fuse(Lt, w)
    alert = st >= threshold(sv, q)
    att = df.attack.values[fi_te][B]
    C = pd.DataFrame({"attack": att, "alert": alert, "c_Benford": w["B"] * Lt.B.values, "c_IF": w["D"] * Lt.D.values})
    C["dominant"] = np.where(C.c_Benford >= C.c_IF, "Benford", "IF")
    al = C[C.alert]
    lay = al.groupby("attack").agg(alerts=("alert", "size"), mean_c_Benford=("c_Benford", "mean"), mean_c_IF=("c_IF", "mean"),
                                   share_Benford_dominant=("dominant", lambda s: (s == "Benford").mean()))
    lay["recall"] = C.groupby("attack").alert.mean()
    lay.to_csv(out / "explain_layers.csv")
    print("weights", w); print(lay.round(3).to_string())
    # TreeSHAP for IF-dominant alerts (sample up to 300 per attack class)
    rows = []
    ex = shap.TreeExplainer(ifm)
    XB = Xte[B]
    for a, g in al[al.dominant == "IF"].groupby("attack"):
        idx = g.index.values[:300]
        sv_ = -np.asarray(ex.shap_values(XB[idx]))
        m = pd.Series(np.abs(sv_).mean(0), index=names).sort_values(ascending=False).head(5)
        for k, (fn, v) in enumerate(m.items()):
            rows.append(dict(attack=a, rank=k + 1, feature=fn, mean_abs_shap=v, n=len(idx)))
    pd.DataFrame(rows).to_csv(out / "explain_shap_if.csv", index=False)
    # Benford per-feature attribution for Benford-dominant alerts (window level)
    bc = f0.benford_contrib(W)
    rows = []
    for a, g in al[al.dominant == "Benford"].groupby("attack"):
        wins = codes[fi_te[B][g.index.values]]
        m = bc.iloc[wins].mean().sort_values(ascending=False)
        rows.append(dict(attack=a, n=len(g), **m.to_dict()))
    pd.DataFrame(rows).to_csv(out / "explain_benford.csv", index=False)
    print("explain done")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
