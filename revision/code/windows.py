"""Window-level feature construction: Benford digit metrics, metadata aggregates, graph topology."""
from __future__ import annotations

import ipaddress

import networkx as nx
import numpy as np
import pandas as pd

BENFORD_CANDIDATES = ["duration", "fwd_bytes", "bwd_bytes", "total_bytes", "total_pkts", "fwd_pkts",
                      "flow_bytes_s", "flow_pkts_s", "iat_mean", "fwd_iat_mean", "bwd_iat_mean", "pkt_len_mean"]
P1 = np.log10(1 + 1 / np.arange(1, 10))
P2 = np.log10(1 + 1 / np.arange(10, 100))
METRICS = ("mad", "ks", "chi2", "euclid", "entgap")


def leading_digits(x: np.ndarray, nd: int) -> np.ndarray:
    """Return first nd significant digits as int (10**(nd-1)..10**nd-1), or 0 when undefined."""
    x = np.abs(np.asarray(x, dtype=float))
    ok = np.isfinite(x) & (x > 0)
    out = np.zeros(x.shape, dtype=np.int16)
    xv = x[ok]
    mag = np.floor(np.log10(xv))
    d = np.floor(xv / 10.0 ** mag * 10 ** (nd - 1) + 1e-9).astype(int)
    d = np.clip(d, 10 ** (nd - 1), 10 ** nd - 1)
    out[ok] = d
    return out


def digit_counts(win: np.ndarray, digits: np.ndarray, n_win: int, nd: int) -> np.ndarray:
    lo = 10 ** (nd - 1)
    k = 9 * lo
    m = digits > 0
    counts = np.zeros((n_win, k), dtype=np.float64)
    np.add.at(counts, (win[m], digits[m] - lo), 1.0)
    return counts


def benford_metrics(counts: np.ndarray, nd: int, min_n: int) -> dict[str, np.ndarray]:
    p = P1 if nd == 1 else P2
    n = counts.sum(1)
    with np.errstate(invalid="ignore", divide="ignore"):
        obs = counts / n[:, None]
        res = {
            "mad": np.abs(obs - p).mean(1),
            "ks": np.abs(np.cumsum(obs, 1) - np.cumsum(p)).max(1),
            "chi2": (((counts - n[:, None] * p) ** 2) / (n[:, None] * p)).sum(1) / np.maximum(n, 1),
            "euclid": np.linalg.norm(obs - p, axis=1),
        }
        ent = -np.nansum(np.where(obs > 0, obs * np.log2(obs), 0.0), 1)
        res["entgap"] = np.abs(ent - (-(p * np.log2(p)).sum()))
    for k in res:
        res[k][n < min_n] = np.nan
    res["n"] = n
    return res


def _entropy_by_window(win: np.ndarray, vals: pd.Series, n_win: int) -> np.ndarray:
    g = pd.DataFrame({"w": win, "v": vals.values}).groupby(["w", "v"], sort=False).size()
    cnt = g.values.astype(float)
    w = g.index.get_level_values(0).values
    tot = np.bincount(w, weights=cnt, minlength=n_win)
    p = cnt / tot[w]
    h = np.bincount(w, weights=-p * np.log2(p), minlength=n_win)
    uniq = np.bincount(w, minlength=n_win).astype(float)
    top = pd.Series(p).groupby(w).max().reindex(range(n_win)).fillna(0).values
    return h, uniq, top


INTERNAL_EXTRA = [ipaddress.ip_network(n) for n in ("149.171.126.0/24", "175.45.176.0/24")]


def _is_internal(ip: str) -> bool:
    try:
        a = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return a.is_private or any(a in n for n in INTERNAL_EXTRA)


def graph_features(df: pd.DataFrame, win: np.ndarray, n_win: int, seed: int = 0):
    """Per-window directed host graph (internal IPs; externals collapsed to one super-node)."""
    ips = pd.unique(pd.concat([df["src"], df["dst"]]))
    internal = {ip: _is_internal(ip) for ip in ips}
    s = df["src"].map(lambda x: x if internal[x] else "EXT")
    d = df["dst"].map(lambda x: x if internal[x] else "EXT")
    e = pd.DataFrame({"w": win, "s": s.values, "d": d.values, "b": df["total_bytes"].fillna(0).values})
    e = e.groupby(["w", "s", "d"], sort=True).agg(n=("b", "size"), b=("b", "sum")).reset_index()
    cols = ["g_nodes", "g_edges", "g_density", "g_max_in_share", "g_max_out_share", "g_max_pagerank",
            "g_max_betweenness", "g_top_host_score"]
    F = np.full((n_win, len(cols)), np.nan)
    top_host = np.array([""] * n_win, dtype=object)
    for w, grp in e.groupby("w", sort=False):
        grp = grp[grp["n"] >= 3]  # edge threshold: >=3 flows per pair per window
        if len(grp) == 0:
            F[w] = 0
            continue
        G = nx.DiGraph()
        G.add_weighted_edges_from(zip(grp["s"], grp["d"], grp["b"].astype(float) + 1.0))
        tw = grp["b"].sum() + len(grp)
        outw, inw = dict(G.out_degree(weight="weight")), dict(G.in_degree(weight="weight"))
        pr = nx.pagerank(G, weight="weight", max_iter=200, tol=1e-6)
        U = G.to_undirected()
        k = min(64, U.number_of_nodes())
        bt = nx.betweenness_centrality(U, k=k, seed=seed) if U.number_of_nodes() > 2 else {n: 0 for n in U}
        score = {n: 0.35 * outw.get(n, 0) / tw + 0.35 * inw.get(n, 0) / tw + 0.2 * pr[n] + 0.1 * bt.get(n, 0)
                 for n in G if n != "EXT"} or {"EXT": 0.0}
        h = max(score, key=score.get)
        nn = G.number_of_nodes()
        F[w] = [nn, G.number_of_edges(), nx.density(G), max(inw.values()) / tw, max(outw.values()) / tw,
                max(pr.values()), max(bt.values()), score[h]]
        top_host[w] = h
    return pd.DataFrame(F, columns=cols), top_host


def build_windows(df: pd.DataFrame, minutes: int, min_n: int = 50, with_graph: bool = True,
                  candidates=BENFORD_CANDIDATES) -> tuple[pd.DataFrame, np.ndarray]:
    """Returns (window table, flow->window index)."""
    ws = df["ts"].dt.floor(f"{minutes}min")
    codes, uniq = pd.factorize(ws, sort=True)
    n = len(uniq)
    W = pd.DataFrame({"window_start": uniq})
    if "dayfile" in df:
        first = pd.Series(df["dayfile"].values).groupby(codes).first()
        W["day"] = first.reindex(range(n)).values
    else:
        W["day"] = W["window_start"].dt.normalize()
    lab = df["label"].values
    W["n_flows"] = np.bincount(codes, minlength=n)
    W["attack_frac"] = np.bincount(codes, weights=lab, minlength=n) / W["n_flows"]
    W["y"] = (W["attack_frac"] > 0).astype(int)
    at = df.loc[lab == 1, "attack"]
    if len(at):
        dom = pd.DataFrame({"w": codes[lab == 1], "a": at.values}).groupby("w")["a"].agg(lambda s: s.value_counts().index[0])
        W["dominant_attack"] = dom.reindex(range(n)).fillna("BENIGN").values
    else:
        W["dominant_attack"] = "BENIGN"
    # --- Benford metrics
    cand = [c for c in candidates if c in df]
    BF = {}
    for c in cand:
        x = df[c].values
        for nd in (1, 2):
            m = benford_metrics(digit_counts(codes, leading_digits(x, nd), n, nd), nd, min_n if nd == 1 else 4 * min_n)
            for k in METRICS:
                BF[f"bf_{c}_d{nd}_{k}"] = m[k]
    W = pd.concat([W, pd.DataFrame(BF)], axis=1)
    # --- metadata aggregates (for IF / baselines)
    A = {}
    A["a_log_flows"] = np.log1p(W["n_flows"].values)
    for col, nm in [("dst", "dst"), ("src", "src"), ("dport", "dport")]:
        h, u, top = _entropy_by_window(codes, df[col], n)
        A[f"a_ent_{nm}"], A[f"a_log_uniq_{nm}"], A[f"a_top_share_{nm}"] = h, np.log1p(u), top
    for pv, nm in [(6, "tcp"), (17, "udp")]:
        A[f"a_share_{nm}"] = np.bincount(codes, weights=(df["proto"].values == pv), minlength=n) / W["n_flows"].values
    A["a_share_tiny"] = np.bincount(codes, weights=(df["total_pkts"].fillna(0).values <= 2), minlength=n) / W["n_flows"].values
    g = pd.DataFrame({"w": codes})
    for c in ["total_bytes", "total_pkts", "duration", "pkt_len_mean", "flow_pkts_s"]:
        g[c] = np.log1p(df[c].clip(lower=0).fillna(0).values)
    agg = g.groupby("w").agg(["mean", "std", "median", lambda s: s.quantile(0.95)])
    agg.columns = [f"a_{c}_{s if isinstance(s, str) else 'p95'}" for c, s in agg.columns]
    agg = agg.rename(columns=lambda x: x.replace("<lambda_0>", "p95")).reindex(range(n))
    for c in agg:
        A[c] = agg[c].fillna(0).values
    W = pd.concat([W, pd.DataFrame(A)], axis=1)
    if with_graph:
        G, top = graph_features(df, codes, n)
        G["g_top_host"] = top
        W = pd.concat([W, G], axis=1)
    return W, codes
