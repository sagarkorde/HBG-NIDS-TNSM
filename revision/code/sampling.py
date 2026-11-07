"""Window-level effect of uniform flow sampling on the Benford context score (CIC-IDS-2017, full data available)."""
import sys
from pathlib import Path
import numpy as np, pandas as pd
from scipy.stats import spearmanr
sys.path.insert(0, str(Path(__file__).parent))
from detectors import FBTI, screen_benford
from loaders import load_dir
from run import split_masks
from windows import BENFORD_CANDIDATES, build_windows

out = Path(sys.argv[2]); out.mkdir(parents=True, exist_ok=True)
full = load_dir(Path(sys.argv[1]), "cic2017")
ref = None; rows = []
for frac in [1.0, 0.5, 0.2, 0.1, 0.05]:
    df = full if frac == 1.0 else full.sample(frac=frac, random_state=0).sort_values("ts", kind="stable").reset_index(drop=True)
    W, codes = build_windows(df, 1, with_graph=False)
    tr, va, te = split_masks(W, "cic2017", 1)
    scr = screen_benford(df[np.isin(codes, np.flatnonzero(tr))], BENFORD_CANDIDATES)
    feats = scr.loc[scr.passed, "feature"].tolist() or ["flow_bytes_s"]
    f = FBTI(feats, seed=0).fit(W, tr)
    L = f.layers(W)
    bc = [c for c in W.columns if c.startswith("bf_") and any(c.startswith(f"bf_{x}_d1") for x in feats)]
    scored = W[bc].notna().any(axis=1).values
    s = pd.Series(L.B.values, index=W.window_start.values)
    if frac == 1.0:
        ref = s; ref_te = W.window_start.values[te]
    common = np.intersect1d(ref_te, s.index.values)
    rho = spearmanr(ref.loc[common].values, s.loc[common].values).correlation
    from sklearn.metrics import roc_auc_score
    auc = roc_auc_score(W.y.values[te], L.B.values[te])
    rows.append(dict(sample_frac=frac, flows=len(df), median_flows_per_window=float(np.median(W.n_flows)),
                     p10_flows_per_window=float(np.percentile(W.n_flows, 10)), features=";".join(feats),
                     scored_window_share=float(scored[te].mean()), spearman_vs_full=float(rho), window_auc_benford=float(auc)))
    print(rows[-1], flush=True)
pd.DataFrame(rows).to_csv(out / "sampling_window.csv", index=False)
