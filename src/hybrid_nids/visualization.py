from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path

import gravis as gv
import matplotlib
import networkx as nx
import numpy as np
import pandas as pd
from matplotlib import cm, colors
from pyvis.network import Network


matplotlib.use("Agg")
from matplotlib import pyplot as plt


@dataclass(slots=True)
class FocusWindow:
    source_file: str
    window_start: pd.Timestamp
    final_score: float
    window_label: str
    alert_context: str


def _sanitize_filename(value: str) -> str:
    lowered = value.strip().lower()
    lowered = re.sub(r"[^a-z0-9]+", "_", lowered)
    return lowered.strip("_")


def _normalize_series(series: pd.Series, min_value: float, max_value: float) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce").fillna(0.0)
    value_min = float(numeric.min())
    value_max = float(numeric.max())
    if math.isclose(value_min, value_max):
        return pd.Series(np.full(len(numeric), (min_value + max_value) / 2.0), index=numeric.index)
    scaled = (numeric - value_min) / (value_max - value_min)
    return min_value + scaled * (max_value - min_value)


def load_pipeline_outputs(output_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    window_df = pd.read_csv(output_dir / "window_scores.csv", parse_dates=["window_start"])
    host_df = pd.read_csv(output_dir / "host_scores.csv", parse_dates=["window_start"])
    edge_df = pd.read_csv(output_dir / "edge_scores.csv", parse_dates=["window_start"])

    for frame in (window_df, host_df, edge_df):
        for bool_col in ("alert", "true_attack_window", "is_training_window"):
            if bool_col in frame.columns:
                frame[bool_col] = frame[bool_col].astype(str).str.lower().map({"true": True, "false": False})

    numeric_columns = {
        "window_df": ["final_score"],
        "host_df": ["node_score", "weighted_degree", "pagerank", "betweenness"],
        "edge_df": ["edge_score", "total_bytes", "flow_count"],
    }
    for column in numeric_columns["window_df"]:
        if column in window_df.columns:
            window_df[column] = pd.to_numeric(window_df[column], errors="coerce")
    for column in numeric_columns["host_df"]:
        if column in host_df.columns:
            host_df[column] = pd.to_numeric(host_df[column], errors="coerce")
    for column in numeric_columns["edge_df"]:
        if column in edge_df.columns:
            edge_df[column] = pd.to_numeric(edge_df[column], errors="coerce")

    return window_df, host_df, edge_df


def select_focus_window(
    window_df: pd.DataFrame,
    source_file: str | None = None,
    window_start: str | None = None,
) -> FocusWindow:
    filtered = window_df.copy()
    if source_file:
        filtered = filtered[filtered["source_file"] == source_file]
    if window_start:
        filtered = filtered[filtered["window_start"] == pd.Timestamp(window_start)]

    if filtered.empty:
        raise ValueError("No matching window found for the requested filters.")

    ranked = filtered.sort_values(["alert", "final_score"], ascending=[False, False]).iloc[0]
    return FocusWindow(
        source_file=str(ranked["source_file"]),
        window_start=pd.Timestamp(ranked["window_start"]),
        final_score=float(ranked["final_score"]),
        window_label=str(ranked["window_label"]),
        alert_context=str(ranked.get("alert_context", "")),
    )


def build_focus_graph(
    host_df: pd.DataFrame,
    edge_df: pd.DataFrame,
    focus: FocusWindow,
    max_nodes: int = 24,
    max_edges: int = 40,
) -> tuple[nx.DiGraph, pd.DataFrame, pd.DataFrame]:
    host_slice = host_df[
        (host_df["source_file"] == focus.source_file)
        & (host_df["window_start"] == focus.window_start)
    ].copy()
    edge_slice = edge_df[
        (edge_df["source_file"] == focus.source_file)
        & (edge_df["window_start"] == focus.window_start)
    ].copy()

    if host_slice.empty or edge_slice.empty:
        raise ValueError("No host or edge data found for the selected focus window.")

    host_slice = host_slice.sort_values("node_score", ascending=False).reset_index(drop=True)
    edge_slice = edge_slice.sort_values("edge_score", ascending=False).reset_index(drop=True)

    selected_nodes: set[str] = set(host_slice.head(min(max_nodes, len(host_slice)))["host"].astype(str))
    selected_edges: list[int] = []

    for idx, row in edge_slice.iterrows():
        src = str(row["source_ip"])
        dst = str(row["destination_ip"])
        proposed = selected_nodes | {src, dst}
        if len(proposed) <= max_nodes:
            selected_nodes = proposed
            selected_edges.append(idx)
        if len(selected_edges) >= max_edges:
            break

    if not selected_edges:
        selected_edges = list(edge_slice.head(min(max_edges, len(edge_slice))).index)
        selected_nodes = set(
            edge_slice.loc[selected_edges, ["source_ip", "destination_ip"]].astype(str).stack().tolist()
        )

    host_focus = host_slice[host_slice["host"].astype(str).isin(selected_nodes)].copy()
    edge_focus = edge_slice.loc[selected_edges].copy()
    edge_focus = edge_focus[
        edge_focus["source_ip"].astype(str).isin(selected_nodes)
        & edge_focus["destination_ip"].astype(str).isin(selected_nodes)
    ].copy()

    host_focus["viz_size"] = _normalize_series(host_focus["node_score"], 18.0, 52.0)
    edge_focus["viz_size"] = _normalize_series(edge_focus["edge_score"], 1.2, 6.5)

    node_norm = colors.Normalize(
        vmin=float(host_focus["node_score"].min()),
        vmax=float(host_focus["node_score"].max()) if len(host_focus) > 1 else float(host_focus["node_score"].max()) + 1e-9,
    )
    edge_norm = colors.Normalize(
        vmin=float(edge_focus["edge_score"].min()),
        vmax=float(edge_focus["edge_score"].max()) if len(edge_focus) > 1 else float(edge_focus["edge_score"].max()) + 1e-9,
    )
    node_cmap = cm.get_cmap("YlOrRd")
    edge_cmap = cm.get_cmap("PuBu")

    graph = nx.DiGraph()
    for row in host_focus.itertuples(index=False):
        node_color = colors.to_hex(node_cmap(node_norm(float(row.node_score))))
        graph.add_node(
            str(row.host),
            label=str(row.host),
            title=(
                f"Host: {row.host}<br>"
                f"Node score: {float(row.node_score):.4f}<br>"
                f"Weighted degree: {float(row.weighted_degree):.1f}<br>"
                f"Pagerank: {float(row.pagerank):.4f}<br>"
                f"Betweenness: {float(row.betweenness):.4f}"
            ),
            size=float(row.viz_size),
            color=node_color,
            score=float(row.node_score),
        )

    for row in edge_focus.itertuples(index=False):
        edge_color = colors.to_hex(edge_cmap(edge_norm(float(row.edge_score))))
        graph.add_edge(
            str(row.source_ip),
            str(row.destination_ip),
            title=(
                f"{row.source_ip} -> {row.destination_ip}<br>"
                f"Edge score: {float(row.edge_score):.4f}<br>"
                f"Flow count: {int(row.flow_count)}<br>"
                f"Total bytes: {float(row.total_bytes):.1f}"
            ),
            width=float(row.viz_size),
            size=float(row.viz_size),
            color=edge_color,
            score=float(row.edge_score),
            total_bytes=float(row.total_bytes),
            flow_count=int(row.flow_count),
        )

    return graph, host_focus, edge_focus


def export_networkx_static(
    graph: nx.DiGraph,
    host_focus: pd.DataFrame,
    edge_focus: pd.DataFrame,
    focus: FocusWindow,
    output_dir: Path,
) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{_sanitize_filename(focus.source_file)}_{focus.window_start.strftime('%Y%m%d_%H%M%S')}"
    png_path = output_dir / f"{base_name}_networkx.png"
    svg_path = output_dir / f"{base_name}_networkx.svg"

    pos = nx.spring_layout(graph, seed=42, k=1.5 / max(math.sqrt(max(graph.number_of_nodes(), 1)), 1.0), iterations=200)

    node_order = list(graph.nodes())
    node_scores = np.array([graph.nodes[node]["score"] for node in node_order], dtype=float)
    node_sizes = np.array([graph.nodes[node]["size"] for node in node_order], dtype=float) * 25.0
    node_colors = [graph.nodes[node]["color"] for node in node_order]

    edge_order = list(graph.edges())
    edge_widths = [graph.edges[edge]["width"] for edge in edge_order]
    edge_colors = [graph.edges[edge]["color"] for edge in edge_order]

    fig, ax = plt.subplots(figsize=(16, 11), constrained_layout=True)
    ax.set_facecolor("#fcfcfb")

    nx.draw_networkx_edges(
        graph,
        pos,
        edgelist=edge_order,
        width=edge_widths,
        edge_color=edge_colors,
        alpha=0.65,
        arrows=True,
        arrowsize=18,
        ax=ax,
        connectionstyle="arc3,rad=0.08",
    )
    nx.draw_networkx_nodes(
        graph,
        pos,
        nodelist=node_order,
        node_size=node_sizes,
        node_color=node_colors,
        linewidths=1.2,
        edgecolors="#243447",
        ax=ax,
    )

    label_count = min(12, len(host_focus))
    label_nodes = set(host_focus.sort_values("node_score", ascending=False).head(label_count)["host"].astype(str))
    labels = {node: node for node in node_order if node in label_nodes}
    nx.draw_networkx_labels(
        graph,
        pos,
        labels=labels,
        font_size=10,
        font_weight="bold",
        font_family="DejaVu Sans",
        bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.75, "pad": 0.25},
        ax=ax,
    )

    node_mapper = cm.ScalarMappable(norm=colors.Normalize(vmin=float(node_scores.min()), vmax=float(node_scores.max()) + 1e-9), cmap=cm.get_cmap("YlOrRd"))
    edge_mapper = cm.ScalarMappable(norm=colors.Normalize(vmin=float(edge_focus["edge_score"].min()), vmax=float(edge_focus["edge_score"].max()) + 1e-9), cmap=cm.get_cmap("PuBu"))
    cbar_nodes = fig.colorbar(node_mapper, ax=ax, fraction=0.03, pad=0.02)
    cbar_nodes.set_label("Host Correlation Score", fontsize=11)
    cbar_edges = fig.colorbar(edge_mapper, ax=ax, fraction=0.03, pad=0.08)
    cbar_edges.set_label("Edge Correlation Score", fontsize=11)

    ax.set_title(
        "Host-to-Host Correlation Graph\n"
        f"{focus.source_file} | {focus.window_start} | score={focus.final_score:.3f} | label={focus.window_label}",
        fontsize=16,
        fontweight="bold",
    )
    ax.text(
        0.01,
        0.02,
        f"Context: {focus.alert_context}",
        transform=ax.transAxes,
        fontsize=10,
        color="#334155",
        bbox={"facecolor": "white", "edgecolor": "#cbd5e1", "alpha": 0.85, "pad": 0.4},
    )
    ax.axis("off")

    fig.savefig(png_path, dpi=320, bbox_inches="tight", facecolor=fig.get_facecolor())
    fig.savefig(svg_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return png_path, svg_path


def export_pyvis_html(graph: nx.DiGraph, focus: FocusWindow, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{_sanitize_filename(focus.source_file)}_{focus.window_start.strftime('%Y%m%d_%H%M%S')}"
    html_path = output_dir / f"{base_name}_pyvis.html"

    network = Network(
        height="850px",
        width="100%",
        directed=True,
        bgcolor="#fbfbf8",
        font_color="#1f2937",
    )
    network.barnes_hut(gravity=-1800, central_gravity=0.12, spring_length=120, spring_strength=0.015)

    for node, attrs in graph.nodes(data=True):
        network.add_node(
            node,
            label=attrs.get("label", node),
            title=attrs.get("title", node),
            color=attrs.get("color", "#f59e0b"),
            size=float(attrs.get("size", 20.0)),
        )

    for source, target, attrs in graph.edges(data=True):
        network.add_edge(
            source,
            target,
            title=attrs.get("title", f"{source} -> {target}"),
            color=attrs.get("color", "#2563eb"),
            width=float(attrs.get("width", 1.5)),
            arrows="to",
        )

    network.set_options(
        """
        {
          "interaction": {
            "hover": true,
            "navigationButtons": true,
            "multiselect": true
          },
          "physics": {
            "stabilization": {
              "iterations": 250
            }
          },
          "edges": {
            "smooth": {
              "type": "dynamic"
            }
          }
        }
        """
    )
    network.write_html(str(html_path), open_browser=False, notebook=False)
    return html_path


def export_gravis_views(graph: nx.DiGraph, focus: FocusWindow, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{_sanitize_filename(focus.source_file)}_{focus.window_start.strftime('%Y%m%d_%H%M%S')}"
    html_path = output_dir / f"{base_name}_gravis.html"

    figure = gv.vis(
        graph,
        graph_height=900,
        show_menu=True,
        node_size_data_source="size",
        edge_size_data_source="size",
        use_node_size_normalization=False,
        use_edge_size_normalization=False,
        show_node_label=True,
        node_hover_tooltip=True,
        edge_hover_tooltip=True,
        node_label_data_source="label",
        show_details=False,
    )
    html_path.write_text(figure.to_html_standalone(), encoding="utf-8")

    # Browser-driven PNG/SVG export can hang on some Windows setups.
    # NetworkX already provides the publication-ready static outputs, so
    # gravis is used here for its richer interactive HTML rendering.
    return {"html": html_path}


def export_ranked_charts(
    host_df: pd.DataFrame,
    edge_df: pd.DataFrame,
    output_dir: Path,
    top_k: int = 12,
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs: dict[str, Path] = {}

    host_rank = (
        host_df[host_df["alert"] == True]  # noqa: E712
        .groupby("host", as_index=False)
        .agg(alert_windows=("window_start", "nunique"), max_node_score=("node_score", "max"))
        .sort_values(["alert_windows", "max_node_score"], ascending=[False, False])
        .head(top_k)
    )

    edge_rank = (
        edge_df[edge_df["alert"] == True]  # noqa: E712
        .assign(edge=lambda df: df["source_ip"].astype(str) + " -> " + df["destination_ip"].astype(str))
        .groupby("edge", as_index=False)
        .agg(alert_windows=("window_start", "nunique"), max_edge_score=("edge_score", "max"))
        .sort_values(["alert_windows", "max_edge_score"], ascending=[False, False])
        .head(top_k)
    )

    charts = [
        (
            host_rank,
            "host",
            "alert_windows",
            "Top Alerted Hosts",
            "Alert Windows",
            "top_alert_hosts",
            "#c2410c",
        ),
        (
            edge_rank,
            "edge",
            "alert_windows",
            "Top Alerted Host-to-Host Edges",
            "Alert Windows",
            "top_alert_edges",
            "#1d4ed8",
        ),
    ]

    for frame, label_col, value_col, title, xlabel, stem, color in charts:
        if frame.empty:
            continue
        frame = frame.sort_values(value_col, ascending=True)
        fig, ax = plt.subplots(figsize=(14, 8), constrained_layout=True)
        ax.barh(frame[label_col], frame[value_col], color=color, alpha=0.88)
        ax.set_title(title, fontsize=16, fontweight="bold")
        ax.set_xlabel(xlabel, fontsize=12)
        ax.set_ylabel("")
        ax.grid(axis="x", linestyle="--", alpha=0.3)
        png_path = output_dir / f"{stem}.png"
        svg_path = output_dir / f"{stem}.svg"
        fig.savefig(png_path, dpi=320, bbox_inches="tight")
        fig.savefig(svg_path, bbox_inches="tight")
        plt.close(fig)
        outputs[f"{stem}_png"] = png_path
        outputs[f"{stem}_svg"] = svg_path

    return outputs


def render_visualization_bundle(
    output_dir: Path,
    source_file: str | None = None,
    window_start: str | None = None,
    max_nodes: int = 24,
    max_edges: int = 40,
) -> dict[str, object]:
    output_dir = output_dir.resolve()
    viz_dir = output_dir / "visualizations"
    window_df, host_df, edge_df = load_pipeline_outputs(output_dir)
    focus = select_focus_window(window_df, source_file=source_file, window_start=window_start)
    graph, host_focus, edge_focus = build_focus_graph(
        host_df=host_df,
        edge_df=edge_df,
        focus=focus,
        max_nodes=max_nodes,
        max_edges=max_edges,
    )

    static_png, static_svg = export_networkx_static(graph, host_focus, edge_focus, focus, viz_dir)
    pyvis_html = export_pyvis_html(graph, focus, viz_dir)
    gravis_outputs = export_gravis_views(graph, focus, viz_dir)
    ranking_outputs = export_ranked_charts(host_df, edge_df, viz_dir)

    return {
        "focus_window": {
            "source_file": focus.source_file,
            "window_start": str(focus.window_start),
            "final_score": focus.final_score,
            "window_label": focus.window_label,
            "alert_context": focus.alert_context,
            "node_count": graph.number_of_nodes(),
            "edge_count": graph.number_of_edges(),
        },
        "outputs": {
            "networkx_png": str(static_png),
            "networkx_svg": str(static_svg),
            "pyvis_html": str(pyvis_html),
            **{f"gravis_{key}": str(path) for key, path in gravis_outputs.items()},
            **{key: str(path) for key, path in ranking_outputs.items()},
        },
    }
