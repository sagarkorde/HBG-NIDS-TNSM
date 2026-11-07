"""FBTI (Benford + Temporal + IF [+ Graph]) and all baselines (unsupervised + supervised)."""
from __future__ import annotations

import itertools

import numpy as np
import pandas as pd
from scipy.cluster.hierarchy import linkage, to_tree
from sklearn.ensemble import IsolationForest, RandomForestClassifier
from sklearn.neural_network import MLPClassifier, MLPRegressor
from sklearn.preprocessing import StandardScaler

from windows import leading_digits, P1

EPS = 1e-9


def posz(x, mu, sd):
    sd = np.where(np.asarray(sd) < EPS, 1.0, sd)
    return np.clip((x - mu) / sd, 0, None)


# ----------------------------------------------------------------------------- Benford screening
def screen_benford(train_flows: pd.DataFrame, candidates, mad_max=0.015, dr_min=2.0, sample=200_000, seed=0):
    """Dual criterion on benign training flows: Nigrini first-digit MAD conformity + dynamic range (orders of mag.).
    KS p-value is reported for reference (it rejects for any n in the 10^5 range, hence MAD is used)."""
    from scipy.stats import kstest  # noqa
    rows = []
    rs = np.random.default_rng(seed)
    for c in candidates:
        if c not in train_flows:
            continue
        x = train_flows[c].values.astype(float)
        x = x[np.isfinite(x) & (x > 0)]
        if len(x) < 100:
            continue
        if len(x) > sample:
            x = rs.choice(x, sample, replace=False)
        d = leading_digits(x, 1)
        obs = np.bincount(d[d > 0] - 1, minlength=9) / max((d > 0).sum(), 1)
        mad = np.abs(obs - P1).mean()
        dr = np.log10(np.percentile(x, 99) / max(np.percentile(x, 1), EPS))
        ks_stat = np.abs(np.cumsum(obs) - np.cumsum(P1)).max()
        n = len(x)
        # asymptotic KS p-value for a discrete distribution (conservative)
        ks_p = float(np.exp(-2 * n * ks_stat ** 2) * 2)
        rows.append(dict(feature=c, n=n, mad=mad, dyn_range_oom=dr, ks_stat=ks_stat, ks_p=min(1.0, ks_p),
                         passed=bool(mad <= mad_max and dr >= dr_min)))
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------- FBTI
class FBTI:
    LAYERS = ("B", "T", "I", "G")

    def __init__(self, benford_features, alpha=0.30, k=0.50, n_estimators=300, seed=42, lam=0.5):
        self.bf = list(benford_features)
        self.alpha, self.k, self.n_est, self.seed, self.lam = alpha, k, n_estimators, seed, lam

    def _bcols(self, W):
        return [c for c in W.columns if c.startswith("bf_") and any(c.startswith(f"bf_{f}_") for f in self.bf)]

    @staticmethod
    def acols(W):
        return [c for c in W.columns if c.startswith("a_")]

    @staticmethod
    def gcols(W):
        return [c for c in W.columns if c.startswith("g_") and c != "g_top_host"]

    def fit(self, W: pd.DataFrame, train_mask: np.ndarray):
        tr = W[train_mask]
        self.bcols = [c for c in self._bcols(W) if tr[c].notna().mean() > 0.5]
        self.b_mu, self.b_sd = tr[self.bcols].mean(), tr[self.bcols].std(ddof=0)
        s = self._stat_raw(tr)
        self.s_mu, self.s_sd = float(np.nanmean(s)), float(np.nanstd(s) or 1)
        self.ac = self.acols(W)
        self.sc = StandardScaler().fit(tr[self.ac].fillna(0))
        self.iforest = IsolationForest(n_estimators=self.n_est, random_state=self.seed, n_jobs=-1).fit(
            self.sc.transform(tr[self.ac].fillna(0)))
        self.gc = self.gcols(W)
        if self.gc:
            self.g_mu, self.g_sd = tr[self.gc].mean(), tr[self.gc].std(ddof=0)
        raw = self._raw_layers(W)
        trr = raw[train_mask]
        self.l_mu, self.l_sd = trr.mean(), trr.std(ddof=0).replace(0, 1)
        return self

    def _stat_raw(self, W):
        Z = posz(W[self.bcols].values, self.b_mu.values, self.b_sd.values)
        return np.nan_to_num(np.nanmean(np.where(np.isnan(W[self.bcols].values), np.nan, Z), 1), nan=0.0)

    def benford_contrib(self, W):
        """Exact per-feature attribution of the Benford layer (it is a mean of z-scores)."""
        Z = posz(W[self.bcols].values, self.b_mu.values, self.b_sd.values)
        Z = np.nan_to_num(Z) / max(len(self.bcols), 1)
        out = pd.DataFrame(Z, columns=self.bcols, index=W.index)
        feat = out.T.groupby(lambda c: c.split("_d")[0][3:]).sum().T
        return feat

    def _temporal(self, W, s):
        z = (s - self.s_mu) / (self.s_sd or 1)
        E = np.zeros_like(s); C = np.zeros_like(s)
        day = W["day"].values
        for i in range(len(s)):
            new = i == 0 or day[i] != day[i - 1]
            E[i] = s[i] if new else self.alpha * s[i] + (1 - self.alpha) * E[i - 1]
            C[i] = max(0.0, (0 if new else C[i - 1]) + z[i] - self.k)
        return E, C

    def _raw_layers(self, W):
        s = self._stat_raw(W)
        E, C = self._temporal(W, s)
        out = pd.DataFrame({"B": s, "E": E, "C": C}, index=W.index)
        out["I"] = -self.iforest.score_samples(self.sc.transform(W[self.ac].fillna(0)))
        if self.gc:
            G = posz(W[self.gc].fillna(0).values, self.g_mu.values, self.g_sd.values)
            out["G"] = G.mean(1)
        else:
            out["G"] = 0.0
        return out

    def layers(self, W):
        raw = self._raw_layers(W)
        z = posz(raw, self.l_mu, self.l_sd)
        z["T"] = 0.6 * z["E"] + 0.4 * z["C"]
        return z[["B", "T", "I", "G"]]

    # --- weight optimisation (paper Eq. 2) on benign validation windows
    def optimise_weights(self, L_val: pd.DataFrame, theta_q=0.95, subset=("B", "T", "I"), step=0.05, lo=0.10, hi=0.70):
        subset = list(subset)
        if len(subset) == 1:
            return {subset[0]: 1.0}
        grid = np.round(np.arange(lo, hi + 1e-9, step), 2)
        best, bw = np.inf, None
        Lv = L_val[subset].values
        for w in itertools.product(grid, repeat=len(subset) - 1):
            last = 1 - sum(w)
            if last < lo - 1e-9 or last > hi + 1e-9:
                continue
            wv = np.array(list(w) + [last])
            sc = Lv @ wv
            # leave-one-out style: threshold from first half, FPR on second half (both benign)
            h = len(sc) // 2
            th = np.quantile(sc[:h], theta_q)
            fpr = float((sc[h:] >= th).mean())
            obj = fpr + self.lam * float(np.std(sc))
            if obj < best - 1e-12:
                best, bw = obj, wv
        if bw is None:  # degenerate validation scores: fall back to equal weights
            return {k: round(1 / len(subset), 3) for k in subset}
        return dict(zip(subset, np.round(bw, 3)))

    @staticmethod
    def fuse(L, weights):
        return sum(L[k].values * v for k, v in weights.items())


# ----------------------------------------------------------------------------- KitNET (Kitsune's detector)
def _sig(x):
    return 1 / (1 + np.exp(-x))


class _DA:
    """Denoising-free tied-weight autoencoder as in Kitsune's dA.py (sigmoid, hidden ratio 0.75)."""

    def __init__(self, n, ratio=0.75, lr=0.1, rng=None):
        self.n, self.h = n, max(1, int(np.ceil(n * ratio)))
        a = 1.0 / n
        self.W = rng.uniform(-a, a, (n, self.h))
        self.hb, self.vb = np.zeros(self.h), np.zeros(n)
        self.lr = lr
        self.mn, self.mx = np.full(n, np.inf), np.full(n, -np.inf)

    def _norm(self, X):
        return (X - self.mn) / (self.mx - self.mn + 1e-16)

    def train(self, X, epochs=1, batch=32, rng=None):
        self.mn, self.mx = np.minimum(self.mn, X.min(0)), np.maximum(self.mx, X.max(0))
        Xn = self._norm(X)
        for _ in range(epochs):
            idx = rng.permutation(len(Xn))
            for s in range(0, len(Xn), batch):
                x = Xn[idx[s:s + batch]]
                y = _sig(x @ self.W + self.hb)
                z = _sig(y @ self.W.T + self.vb)
                L_h2 = x - z
                L_h1 = (L_h2 @ self.W) * y * (1 - y)
                gW = (x.T @ L_h1 + L_h2.T @ y) / len(x)
                self.W += self.lr * gW
                self.hb += self.lr * L_h1.mean(0)
                self.vb += self.lr * L_h2.mean(0)

    def rmse(self, X):
        Xn = self._norm(X)
        z = _sig(_sig(Xn @ self.W + self.hb) @ self.W.T + self.vb)
        return np.sqrt(((Xn - z) ** 2).mean(1))


class KitNET:
    """KitNET (Mirsky et al., NDSS 2018): correlation-clustered ensemble of small autoencoders + output AE.
    Applied to flow/window feature vectors (packet-level AfterImage extraction requires PCAPs)."""

    def __init__(self, max_ae=10, lr=0.1, epochs=3, seed=0):
        self.m, self.lr, self.epochs, self.rs = max_ae, lr, epochs, np.random.default_rng(seed)

    def _cluster(self, X):
        C = np.corrcoef(X, rowvar=False)
        C = np.nan_to_num(C)
        D = 1 - np.abs(C)
        iu = np.triu_indices(len(D), 1)
        Z = linkage(D[iu], "average")
        root = to_tree(Z)

        def rec(node):
            if node.get_count() <= self.m:
                return [node.pre_order()]
            return rec(node.get_left()) + rec(node.get_right())

        return rec(root)

    def fit(self, X):
        X = np.asarray(X, float)
        n_fm = max(1000, len(X) // 10)
        self.maps = self._cluster(X[:n_fm])
        self.ens = [_DA(len(m), lr=self.lr, rng=self.rs) for m in self.maps]
        for ae, m in zip(self.ens, self.maps):
            ae.train(X[:, m], self.epochs, rng=self.rs)
        R = self._ens_rmse(X)
        self.out = _DA(R.shape[1], lr=self.lr, rng=self.rs)
        self.out.train(R, self.epochs, rng=self.rs)
        return self

    def _ens_rmse(self, X):
        return np.column_stack([ae.rmse(X[:, m]) for ae, m in zip(self.ens, self.maps)])

    def score(self, X):
        return self.out.rmse(self._ens_rmse(np.asarray(X, float)))


# ----------------------------------------------------------------------------- baseline zoo
class AE:
    def __init__(self, seed=0, hidden=(64, 32, 64)):
        self.seed, self.hidden = seed, hidden

    def fit(self, X):
        self.sc = StandardScaler().fit(X)
        Xs = np.clip(self.sc.transform(X), -10, 10)
        hid = tuple(min(h, max(4, Xs.shape[1])) for h in self.hidden)
        self.m = MLPRegressor(hidden_layer_sizes=hid, activation="relu", max_iter=300, early_stopping=len(Xs) > 200, random_state=self.seed)
        self.m.fit(Xs, Xs)
        return self

    def score(self, X):
        Xs = np.clip(self.sc.transform(X), -10, 10)
        return ((self.m.predict(Xs) - Xs) ** 2).mean(1)


class Scaled:
    """Wrap a pyod / sklearn detector with standard scaling; higher score = more anomalous."""

    def __init__(self, make):
        self.make = make

    def fit(self, X):
        self.sc = StandardScaler().fit(X)
        self.m = self.make()
        self.m.fit(np.clip(self.sc.transform(X), -10, 10))
        return self

    def score(self, X):
        Xs = np.clip(self.sc.transform(X), -10, 10)
        if hasattr(self.m, "decision_function") and not isinstance(self.m, IsolationForest):
            s = self.m.decision_function(Xs)
        else:
            s = -self.m.score_samples(Xs)
        s = np.asarray(s, float)
        return np.nan_to_num(s, nan=np.nanmax(s) if np.isfinite(s).any() else 0.0, posinf=np.nanmax(s[np.isfinite(s)]) if np.isfinite(s).any() else 0.0)


def unsupervised_zoo(seed=0, n_train=None):
    from pyod.models.copod import COPOD
    from pyod.models.ecod import ECOD
    from pyod.models.hbos import HBOS
    from pyod.models.lof import LOF
    from pyod.models.ocsvm import OCSVM
    from pyod.models.pca import PCA as PPCA
    z = {
        "IF": lambda: Scaled(lambda: IsolationForest(n_estimators=300, random_state=seed, n_jobs=-1)),
        "AE": lambda: AE(seed=seed),
        "KitNET": lambda: _KitWrap(seed),
        "OCSVM": lambda: Scaled(lambda: OCSVM(kernel="rbf", nu=0.05)),
        "LOF": lambda: Scaled(lambda: LOF(n_neighbors=20)),
        "ECOD": lambda: Scaled(lambda: ECOD()),
        "COPOD": lambda: Scaled(lambda: COPOD()),
        "HBOS": lambda: Scaled(lambda: HBOS()),
        "PCA": lambda: Scaled(lambda: PPCA()),
    }
    return z


class _KitWrap:
    def __init__(self, seed):
        self.seed = seed

    def fit(self, X):
        self.sc = StandardScaler().fit(X)
        self.k = KitNET(seed=self.seed).fit(np.clip(self.sc.transform(X), -10, 10))
        return self

    def score(self, X):
        return self.k.score(np.clip(self.sc.transform(X), -10, 10))


def supervised_zoo(seed=0):
    from xgboost import XGBClassifier
    return {
        "RF": lambda: RandomForestClassifier(n_estimators=200, n_jobs=-1, random_state=seed, class_weight="balanced"),
        "XGBoost": lambda: XGBClassifier(n_estimators=300, max_depth=8, learning_rate=0.1, subsample=0.8,
                                         colsample_bytree=0.8, n_jobs=2, random_state=seed, eval_metric="logloss"),
        "MLP": lambda: _SupMLP(seed),
    }


class _SupMLP:
    def __init__(self, seed):
        self.seed = seed

    def fit(self, X, y):
        self.sc = StandardScaler().fit(X)
        self.m = MLPClassifier(hidden_layer_sizes=(128, 64), max_iter=300, early_stopping=len(X) > 200, random_state=self.seed)
        self.m.fit(np.clip(self.sc.transform(X), -10, 10), y)
        return self

    def predict_proba(self, X):
        return self.m.predict_proba(np.clip(self.sc.transform(X), -10, 10))
