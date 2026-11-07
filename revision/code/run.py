"""Full experimental protocol for the revision.

python run.py --kind cic2017 --data DIR --out results/cic2017 [--minutes 1 15] [--seeds 5]
kinds: cic2017 (improved CNS2022 CSVs), cic2017_orig (original MachineLearningCVE/GeneratedLabelledFlows), cic2018, unsw
"""
from __future__ import annotations

import argparse
import itertools
import json
import time
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.model_selection import StratifiedKFold, train_test_split

from detectors import FBTI, screen_benford, supervised_zoo, unsupervised_zoo
from loaders import load_dir
from stats import bootstrap_ci, mcnemar, metrics
from windows import BENFORD_CANDIDATES, build_windows

SLOW = {"OCSVM": 20_000, "LOF": 20_000, "ECOD": 20_000, "COPOD": 20_000, "KitNET": 100_000}


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


def split_masks(W: pd.DataFrame, kind: str, train_days: int):
    """Benign-only training (novelty-detection protocol). CICIDS2017: Monday. Others: benign windows of the
    first `train_days` capture days. Blocked 80/20 train/validation split; test = all later days."""
    days = list(W.groupby("day")["window_start"].min().sort_values().index)
    first = days[:train_days]
    pool = W["day"].isin(first).values & (W["y"].values == 0)
    idx = np.flatnonzero(pool)
    # blocked split: ~50 blocks of consecutive benign windows, every 5th block -> validation (covers the whole day)
    is_val = np.zeros(len(W), bool)
    blk = max(1, len(idx) // 50)
    is_val[idx[(np.arange(len(idx)) // blk) % 5 == 4]] = True
    train = pool & ~is_val
    val = pool & is_val
    test = ~W["day"].isin(first).values
    return train, val, test


def threshold(score_val, q):
    return float(np.quantile(score_val, q))


def sym_log(X):
    return np.sign(X) * np.log1p(np.abs(X))


# ============================================================================ window level
def window_experiment(df, kind, minutes, seeds, q, train_days, out: Path):
    log(f"[window {minutes}min] building windows")
    W, codes = build_windows(df, minutes, with_graph=True)
    tr, va, te = split_masks(W, kind, train_days)
    log(f"windows={len(W)} train={tr.sum()} val={va.sum()} test={te.sum()} test_attack={W.y[te].mean():.3f}")
    train_flows = df[np.isin(codes, np.flatnonzero(tr))]
    scr = screen_benford(train_flows, BENFORD_CANDIDATES)
    scr.to_csv(out / f"benford_screening_{minutes}min.csv", index=False)
    feats = scr.loc[scr.passed, "feature"].tolist() or scr.sort_values("mad").feature.head(3).tolist()
    log("Benford-conformant features:", feats)
    yte = W.y.values[te]
    rows, preds, scores = [], {}, {}
    subsets = [s for r in (1, 2, 3) for s in itertools.combinations("BTI", r)] + [("B", "T", "I", "G")]
    layer_store = None
    for seed in range(seeds):
        f = FBTI(feats, seed=seed).fit(W, tr)
        L = f.layers(W)
        if seed == 0:
            layer_store, fbti0 = L, f
        for sub in subsets:
            for wmode in ("opt", "equal"):
                if wmode == "equal" and len(sub) == 1:
                    continue
                w = f.optimise_weights(L[va], q, sub) if wmode == "opt" else {k: 1 / len(sub) for k in sub}
                s = f.fuse(L, w)
                th = threshold(s[va], q)
                p = s[te] >= th
                name = "FBTI[" + "+".join(sub) + "]" + ("" if wmode == "opt" else "(eq)")
                m = metrics(yte, s[te], p)
                rows.append(dict(method=name, seed=seed, weights=json.dumps(w), threshold=th, **m))
                if seed == 0:
                    preds[name], scores[name] = p, s[te]
    # --- unsupervised baselines on the full window feature vector
    X = W[[c for c in W.columns if c.startswith(("a_", "bf_", "g_")) and c != "g_top_host"]].astype(float)
    X = X.fillna(X[tr].median()).fillna(0).values
    for name, make in unsupervised_zoo().items():
        for seed in range(seeds if name in ("IF", "AE", "KitNET") else 1):
            mdl = unsupervised_zoo(seed)[name]().fit(X[tr])
            s = mdl.score(X)
            th = threshold(s[va], q)
            p = s[te] >= th
            rows.append(dict(method=name, seed=seed, threshold=th, **metrics(yte, s[te], p)))
            if seed == 0:
                preds[name], scores[name] = p, s[te]
        log(" ", name, "done")
    # --- supervised: 5-fold stratified CV on the test windows (out-of-fold predictions)
    Xte = X[te]
    for name in supervised_zoo():
        for seed in range(seeds):
            oof = np.zeros(len(yte))
            for a, b in StratifiedKFold(5, shuffle=True, random_state=seed).split(Xte, yte):
                m_ = supervised_zoo(seed)[name]().fit(Xte[a], yte[a])
                oof[b] = m_.predict_proba(Xte[b])[:, 1]
            p = oof >= 0.5
            rows.append(dict(method=name + " (sup.)", seed=seed, threshold=0.5, **metrics(yte, oof, p)))
            if seed == 0:
                preds[name + " (sup.)"], scores[name + " (sup.)"] = p, oof
    R = pd.DataFrame(rows)
    R.to_csv(out / f"window_{minutes}min_all_runs.csv", index=False)
    summ = R.groupby("method").agg(**{f"{k}_mean": (k, "mean") for k in ["precision", "recall", "f1", "roc_auc", "pr_auc", "fpr"]},
                                   **{f"{k}_std": (k, "std") for k in ["f1", "roc_auc"]}).reset_index()
    ref = "FBTI[B+T+I]"
    ci, mc = {}, {}
    for name in preds:
        ci[name] = bootstrap_ci(yte, scores[name], preds[name])
        mc[name] = mcnemar(yte, preds[ref], preds[name])
    summ["f1_ci"] = summ.method.map(lambda n: ci.get(n, {}).get("f1"))
    summ["auc_ci"] = summ.method.map(lambda n: ci.get(n, {}).get("roc_auc"))
    summ["mcnemar_vs_FBTI_p"] = summ.method.map(lambda n: mc.get(n, {}).get("p"))
    summ["mcnemar_nd"] = summ.method.map(lambda n: mc.get(n, {}).get("n_d"))
    summ.sort_values("f1_mean", ascending=False).to_csv(out / f"window_{minutes}min_summary.csv", index=False)
    # --- per-attack recall + layer attribution for the 3-layer model (seed 0)
    Wt = W[te].copy()
    Wt["alert"] = preds[ref]
    Wt["score"] = scores[ref]
    w0 = json.loads(R[(R.method == ref) & (R.seed == 0)].weights.iloc[0])
    Lt = layer_store[te]
    for k, v in w0.items():
        Wt[f"contrib_{k}"] = v * Lt[k].values
    Wt["dominant_layer"] = Wt[[f"contrib_{k}" for k in w0]].idxmax(1).str[-1]
    per_att = Wt[Wt.y == 1].groupby("dominant_attack").agg(windows=("alert", "size"), recall=("alert", "mean"),
                                                         median_attack_frac=("attack_frac", "median")).reset_index()
    per_att.to_csv(out / f"window_{minutes}min_per_attack.csv", index=False)
    al = Wt[Wt.alert]
    attrib = al.groupby(["y", "dominant_attack"])[[f"contrib_{k}" for k in w0]].mean()
    attrib["n_alerts"] = al.groupby(["y", "dominant_attack"]).size()
    attrib["dominant_layer_mode"] = al.groupby(["y", "dominant_attack"])["dominant_layer"].agg(lambda s: s.value_counts().index[0])
    attrib.to_csv(out / f"window_{minutes}min_layer_attribution.csv")
    # SHAP for IF layer + exact Benford attribution
    try:
        import shap
        Xa = fbti0.sc.transform(W.loc[te, fbti0.ac].fillna(0))
        ex = shap.TreeExplainer(fbti0.iforest)
        sv = ex.shap_values(Xa[Wt.alert.values])
        S = pd.DataFrame(-sv, columns=fbti0.ac)  # negative path length contribution => anomaly
        S["dominant_attack"] = Wt.loc[Wt.alert, "dominant_attack"].values
        S["y"] = Wt.loc[Wt.alert, "y"].values
        S.groupby(["y", "dominant_attack"]).mean().T.to_csv(out / f"window_{minutes}min_shap_IF.csv")
        Bc = fbti0.benford_contrib(W[te])[Wt.alert.values]
        Bc["dominant_attack"] = Wt.loc[Wt.alert, "dominant_attack"].values
        Bc.groupby("dominant_attack").mean().T.to_csv(out / f"window_{minutes}min_benford_attr.csv")
    except Exception as e:  # pragma: no cover
        log("SHAP failed:", e)
    Wt.drop(columns=[c for c in Wt.columns if c.startswith(("bf_",))]).to_csv(out / f"window_{minutes}min_test_windows.csv", index=False)
    log(summ.sort_values("f1_mean", ascending=False)[["method", "precision_mean", "recall_mean", "f1_mean", "roc_auc_mean", "pr_auc_mean"]].head(25).to_string())
    return W, codes, layer_store, fbti0, tr, va, te


# ============================================================================ flow level
def flow_experiment(df, kind, seeds, q, train_days, out: Path, only=None, ctx_minutes=1, n_train=200_000, n_val=50_000,
                    n_test=300_000):
    log("[flow] building 1-min context windows")
    W, codes = build_windows(df, ctx_minutes, with_graph=False)
    trw, vaw, tew = split_masks(W, kind, train_days)
    feats = screen_benford(df[np.isin(codes, np.flatnonzero(trw))], BENFORD_CANDIDATES)
    feats = feats.loc[feats.passed, "feature"].tolist() or ["duration", "total_bytes", "flow_bytes_s"]
    rs = np.random.default_rng(0)
    fi_tr = np.flatnonzero(np.isin(codes, np.flatnonzero(trw)) & (df.label.values == 0))
    fi_va = np.flatnonzero(np.isin(codes, np.flatnonzero(vaw)) & (df.label.values == 0))
    fi_te = np.flatnonzero(np.isin(codes, np.flatnonzero(tew)))
    fi_tr = np.sort(rs.choice(fi_tr, min(n_train, len(fi_tr)), replace=False))
    fi_va = np.sort(rs.choice(fi_va, min(n_val, len(fi_va)), replace=False))
    # stratified test sample that keeps every attack class (cap per class)
    te_df = df.iloc[fi_te]
    keep = []
    cap = n_test // max(te_df.attack.nunique(), 1)
    for a, g in te_df.groupby("attack").groups.items():
        g = np.asarray(g)
        keep.append(rs.choice(g, min(len(g), cap if a != "BENIGN" else n_test // 2), replace=False))
    fi_te = np.sort(np.concatenate(keep))
    fcols = [c for c in df.columns if c.startswith("f_")]
    allix = np.concatenate([fi_tr, fi_va, fi_te])
    Xs = sym_log(df[fcols].values[allix].astype(np.float64))
    Xs = np.nan_to_num(Xs, nan=0.0, posinf=0.0, neginf=0.0)
    keepc = Xs[: len(fi_tr)].std(0) > 0
    Xs = Xs[:, keepc]
    # remap global flow index -> row in Xs (X[...] below is always indexed by fi_tr / fi_va / fi_te)
    pos = {"tr": np.arange(len(fi_tr)), "va": len(fi_tr) + np.arange(len(fi_va)),
           "te": len(fi_tr) + len(fi_va) + np.arange(len(fi_te))}
    class _X:  # tiny adapter so that X[fi_*] keeps working
        def __getitem__(self, ix):
            if ix is fi_tr: return Xs[pos["tr"]]
            if ix is fi_va: return Xs[pos["va"]]
            if ix is fi_te: return Xs[pos["te"]]
            raise KeyError
        shape = Xs.shape
    X = _X()
    y = df.label.values
    yte = y[fi_te]
    # supervised split of the test sample: A (train sup.) / B (evaluate everyone)
    A, B = train_test_split(np.arange(len(fi_te)), test_size=0.5, stratify=yte, random_state=0)
    yB = yte[B]
    log(f"flows train={len(fi_tr)} val={len(fi_va)} test={len(fi_te)} (eval half {len(B)}, attack share {yB.mean():.3f}) features={X.shape[1]}")
    rows, preds, scores = [], {}, {}
    vals = {}
    # FBTI flow variant: window Benford + temporal context + flow-level IF
    for seed in range(seeds):
        f = FBTI(feats, seed=seed).fit(W, trw)
        Lw = f.layers(W)
        from sklearn.ensemble import IsolationForest
        ifm = IsolationForest(n_estimators=300, random_state=seed, n_jobs=-1).fit(X[fi_tr])
        def lay(ix):
            L = pd.DataFrame({"B": Lw.B.values[codes[ix]], "T": Lw["T"].values[codes[ix]]})
            L["I"] = -ifm.score_samples(X[ix])
            return L
        Ltr = lay(fi_tr); mu, sd = Ltr.I.mean(), Ltr.I.std() or 1
        Lva, Lte = lay(fi_va), lay(fi_te)
        for L_ in (Lva, Lte):
            L_["I"] = np.clip((L_.I - mu) / sd, 0, None)
        for sub in [("I",), ("B", "I"), ("T", "I"), ("B", "T", "I")]:
            w = f.optimise_weights(Lva, q, sub)
            s_va, s_te = f.fuse(Lva, w), f.fuse(Lte, w)
            th = threshold(s_va, q)
            p = s_te[B] >= th
            name = "FBTI-flow[" + "+".join(sub) + "]"
            rows.append(dict(method=name, seed=seed, weights=json.dumps(w), **metrics(yB, s_te[B], p)))
            if seed == 0:
                preds[name], scores[name] = p, s_te[B]
                vals[name] = np.asarray(s_va)
        log(" FBTI-flow seed", seed)
    f0 = FBTI(feats, seed=0).fit(W, trw)
    Lw0 = f0.layers(W)
    ctx = lambda ix: pd.DataFrame({"B": Lw0.B.values[codes[ix]], "T": Lw0["T"].values[codes[ix]]})
    Cva, CteB = ctx(fi_va), ctx(fi_te[B])
    for name in unsupervised_zoo():
        if only and name not in only:
            continue
        for seed in range(seeds if name in ("IF", "AE", "KitNET") else 1):
            ntr = SLOW.get(name, len(fi_tr))
            sel = np.random.default_rng(seed).choice(len(fi_tr), min(ntr, len(fi_tr)), replace=False)
            mdl = unsupervised_zoo(seed)[name]().fit(X[fi_tr][sel])
            s_va, s_te = mdl.score(X[fi_va]), mdl.score(X[fi_te][B])
            p = s_te >= threshold(s_va, q)
            rows.append(dict(method=name, seed=seed, **metrics(yB, s_te, p)))
            if seed == 0:
                preds[name], scores[name] = p, s_te
                vals[name] = np.asarray(s_va)
            # plug-in context (pre-specified): z-normalise detector score on training flows, fuse with window layers
            s_tr = mdl.score(X[fi_tr][sel[:20000]])
            mu, sd = float(np.mean(s_tr)), float(np.std(s_tr) or 1)
            for sub in [("B", "D"), ("B", "T", "D")]:
                Lv, Lt = Cva.copy(), CteB.copy()
                Lv["D"], Lt["D"] = np.clip((s_va - mu) / sd, 0, None), np.clip((s_te - mu) / sd, 0, None)
                w = f0.optimise_weights(Lv, q, sub)
                sv, st = f0.fuse(Lv, w), f0.fuse(Lt, w)
                pp = st >= threshold(sv, q)
                nm = name + "+" + "+".join(k for k in sub if k != "D")
                rows.append(dict(method=nm, seed=seed, weights=json.dumps(w), **metrics(yB, st, pp)))
                if seed == 0:
                    preds[nm], scores[nm] = pp, st
                    vals[nm] = np.asarray(sv)
        log(" ", name, "done")
    # day-held-out supervised: train on test flows from the earlier half of the test days, evaluate on the later half
    order = list(df.groupby("dayfile")["ts"].min().sort_values().index)
    tdays = df["dayfile"].values[fi_te]
    ud = [d for d in order if d in set(tdays)]
    late = np.isin(tdays, ud[len(ud) // 2:])
    XA_d, yA_d = X[fi_te][~late], yte[~late]
    Bl = np.asarray(B)[late[B]]
    late_mask_B = late[B]
    dh_rows = []
    for name in supervised_zoo():
        m_ = supervised_zoo(0)[name]().fit(XA_d, yA_d)
        s = m_.predict_proba(X[fi_te][Bl])[:, 1]
        dh_rows.append(dict(method=name + " (sup., day-held-out)", **metrics(yte[Bl], s, s >= 0.5)))
    for n_, p_ in preds.items():
        dh_rows.append(dict(method=n_, **metrics(yB[late_mask_B], scores[n_][late_mask_B], p_[late_mask_B])))
    pd.DataFrame(dh_rows).sort_values("f1", ascending=False).to_csv(out / "flow_dayheldout.csv", index=False)
    log("day-held-out train days:", [str(d)[:10] for d in ud[: len(ud) // 2]], "eval days:", [str(d)[:10] for d in ud[len(ud) // 2:]])
    XA, XB = X[fi_te][A], X[fi_te][B]
    for name in supervised_zoo():
        for seed in range(seeds):
            m_ = supervised_zoo(seed)[name]().fit(XA, yte[A])
            s = m_.predict_proba(XB)[:, 1]
            rows.append(dict(method=name + " (sup.)", seed=seed, **metrics(yB, s, s >= 0.5)))
            if seed == 0:
                preds[name + " (sup.)"], scores[name + " (sup.)"] = s >= 0.5, s
        log(" ", name, "done")
    R = pd.DataFrame(rows)
    R.to_csv(out / "flow_all_runs.csv", index=False)
    summ = R.groupby("method").agg(**{f"{k}_mean": (k, "mean") for k in ["precision", "recall", "f1", "roc_auc", "pr_auc", "fpr"]},
                                   **{f"{k}_std": (k, "std") for k in ["f1", "roc_auc"]}).reset_index()
    ref = "FBTI-flow[B+I]"
    summ["f1_ci"] = summ.method.map(lambda n: bootstrap_ci(yB, scores[n], preds[n], B=300)["f1"] if n in preds else None)
    summ["mcnemar_vs_FBTI_p"] = summ.method.map(lambda n: mcnemar(yB, preds[ref], preds[n])["p"] if n in preds else None)
    summ.sort_values("f1_mean", ascending=False).to_csv(out / "flow_summary.csv", index=False)
    # per-attack recall for every method (seed 0)
    att = df.attack.values[fi_te][B]
    pa = pd.DataFrame({n: pd.Series(p).groupby(att).mean() for n, p in preds.items()})
    pa["n"] = pd.Series(att).value_counts()
    pa.to_csv(out / "flow_per_attack_recall.csv")
    np.savez_compressed(out / "flow_scores_seed0.npz", yB=yB, attackB=att.astype(str),
                        dayB=df["dayfile"].values[fi_te][B].astype(str),
                        **{"te__" + k: np.asarray(v, float) for k, v in scores.items()},
                        **{"va__" + k: np.asarray(v, float) for k, v in vals.items()})
    log(summ.sort_values("f1_mean", ascending=False)[["method", "precision_mean", "recall_mean", "f1_mean", "roc_auc_mean", "pr_auc_mean"]].to_string())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kind", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--minutes", type=int, nargs="+", default=[1, 15])
    ap.add_argument("--seeds", type=int, default=5)
    ap.add_argument("--q", type=float, default=0.95)
    ap.add_argument("--train-days", type=int, default=1)
    ap.add_argument("--no-flow", action="store_true")
    ap.add_argument("--only", nargs="*", default=None, help="restrict unsupervised flow detectors")
    ap.add_argument("--max-flows", type=int, default=0, help="subsample flows (keeps time order) for huge datasets")
    ap.add_argument("--sample-frac", type=float, default=1.0, help="uniform per-flow sampling rate (sensitivity analysis)")
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    log("loading", a.kind)
    df = load_dir(Path(a.data), a.kind)
    if a.max_flows and len(df) > a.max_flows:
        df = df.sample(a.max_flows, random_state=0).sort_values("ts", kind="stable").reset_index(drop=True)
    if a.sample_frac < 1.0:
        df = df.sample(frac=a.sample_frac, random_state=0).sort_values("ts", kind="stable").reset_index(drop=True)
    log(f"{len(df):,} flows; attack share {df.label.mean():.3f}; days {df.dayfile.nunique()}")
    df.groupby(["dayfile", "attack"]).size().to_csv(out / "dataset_composition.csv")
    for m in [m for m in a.minutes if m > 0]:
        window_experiment(df, a.kind, m, a.seeds, a.q, a.train_days, out)
    if not a.no_flow:
        flow_experiment(df, a.kind, a.seeds, a.q, a.train_days, out, only=a.only)
    log("DONE")


if __name__ == "__main__":
    main()
