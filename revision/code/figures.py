"""Figures for the revised manuscript (read from result CSVs only)."""
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
sys.path.insert(0, str(Path(__file__).parent))
from attnames import ATT

R = Path(__file__).resolve().parent.parent / "results"
OUT = Path(__file__).resolve().parent.parent / "figures"; OUT.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({"font.family": "serif", "font.size": 8, "axes.spines.top": False, "axes.spines.right": False})
DETS = ["IF", "KitNET", "AE", "OCSVM", "LOF", "HBOS"]
C_BASE, C_B, C_BT = "#9aa5b1", "#1f5aa6", "#d98c2b"


def fig_plugin():
    fig, axes = plt.subplots(1, 2, figsize=(7.0, 2.3), sharey=True)
    for ax, ds, title in zip(axes, ["cic2017", "cic2018"], ["CIC-IDS-2017 (corrected)", "CSE-CIC-IDS-2018 (corrected, 5% sample)"]):
        r = pd.read_csv(R / ds / "flow_all_runs.csv")
        g = r.groupby("method").f1.agg(["mean", "std"]).fillna(0)
        x = np.arange(len(DETS)); w = 0.27
        for k, (suf, col, lab) in enumerate([("", C_BASE, "detector alone"), ("+B", C_B, "+ Benford context"), ("+B+T", C_BT, "+ Benford + temporal")]):
            m = [g.loc[d + suf, "mean"] if d + suf in g.index else np.nan for d in DETS]
            s = [g.loc[d + suf, "std"] if d + suf in g.index else 0 for d in DETS]
            ax.bar(x + (k - 1) * w, m, w, yerr=s, color=col, label=lab, error_kw=dict(lw=0.6, capsize=1.5))
        ax.set_xticks(x); ax.set_xticklabels(DETS); ax.set_title(title, fontsize=8); ax.set_ylim(0, 1)
        ax.grid(axis="y", lw=0.3, alpha=0.5)
    axes[0].set_ylabel("Flow-level F1")
    h, l = axes[0].get_legend_handles_labels()
    fig.legend(h, l, frameon=False, fontsize=7, loc="upper center", ncol=3, bbox_to_anchor=(0.5, 1.02))
    fig.tight_layout(rect=(0, 0, 1, 0.92)); fig.savefig(OUT / "fig_plugin.png", dpi=300); plt.close(fig)


def fig_attack():
    fig, axes = plt.subplots(1, 2, figsize=(7.0, 2.55), gridspec_kw=dict(width_ratios=[1.15, 1]))
    cols = ["IF", "IF+B", "AE", "AE+B", "KitNET", "KitNET+B", "RF (sup.)"]
    labels = ["IF", "IF+B", "AE", "AE+B", "KitNET", "KitNET+B", "RF (sup.)"]
    for ax, ds, title in zip(axes, ["cic2017", "cic2018"], ["CIC-IDS-2017", "CSE-CIC-IDS-2018"]):
        pa = pd.read_csv(R / ds / "flow_per_attack_recall.csv", index_col=0)
        pa = pa[(pa.index != "BENIGN") & (pa["n"] >= 5)]
        M = pa[cols].values
        im = ax.imshow(M, aspect="auto", cmap="Blues", vmin=0, vmax=1)
        ax.set_xticks(range(len(cols))); ax.set_xticklabels(cols, rotation=40, ha="right", fontsize=6.5)
        ax.set_yticks(range(len(pa))); ax.set_yticklabels([f"{ATT.get(i, i.title())} ({int(n)})" for i, n in zip(pa.index, pa["n"])], fontsize=5.8)
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                ax.text(j, i, f"{M[i, j]:.2f}", ha="center", va="center", fontsize=5.2, color="white" if M[i, j] > 0.6 else "black")
        ax.set_title(title, fontsize=7.5)
    fig.tight_layout(); fig.savefig(OUT / "fig_attack.png", dpi=300); plt.close(fig)


if __name__ == "__main__":
    fig_plugin(); fig_attack(); print("ok")
