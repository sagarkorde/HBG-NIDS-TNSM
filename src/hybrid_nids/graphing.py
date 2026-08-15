from __future__ import annotations

from dataclasses import dataclass

import networkx as nx
import numpy as np
import pandas as pd


GRAPH_FEATURE_COLUMNS = [
    "graph_node_count",
    "graph_edge_count",
    "graph_density",
    "graph_avg_flows_per_edge",
    "graph_avg_bytes_per_edge",
    "graph_max_in_weight_share",
    "graph_max_out_weight_share",
    "graph_max_pagerank",
    "graph_max_betweenness",
    "graph_mean_clustering",
    "graph_largest_component_ratio",
    "graph_top_host_score",
    "graph_top_edge_score",
]


@dataclass(slots=True)
class GraphArtifacts:
    window_metrics: dict[str, float | str]
    host_rows: list[dict[str, object]]
    edge_rows: list[dict[str, object]]


def extract_window_graph_artifacts(
    window_df: pd.DataFrame,
    source_file: str,
    window_start: pd.Timestamp,
) -> GraphArtifacts:
    if window_df.empty:
        empty_metrics = {column: np.nan for column in GRAPH_FEATURE_COLUMNS}
        empty_metrics["graph_top_source"] = ""
        empty_metrics["graph_top_destination"] = ""
        empty_metrics["graph_top_host"] = ""
        empty_metrics["graph_top_host_score"] = 0.0
        empty_metrics["graph_top_edge"] = ""
        empty_metrics["graph_top_edge_score"] = 0.0
        empty_metrics["graph_suspicious_hosts"] = ""
        empty_metrics["graph_suspicious_edges"] = ""
        return GraphArtifacts(window_metrics=empty_metrics, host_rows=[], edge_rows=[])

    edge_df = (
        window_df.groupby(["Source IP", "Destination IP"], dropna=False)
        .agg(
            flow_count=("Flow ID", "size"),
            total_bytes=("total_bytes", "sum"),
            benign_fraction=("is_attack", lambda s: 1.0 - float(s.mean())),
        )
        .reset_index()
    )

    graph = nx.DiGraph()
    for row in edge_df.itertuples(index=False):
        graph.add_edge(
            row[0],
            row[1],
            weight=float(row.total_bytes),
            flow_count=float(row.flow_count),
        )

    node_count = graph.number_of_nodes()
    edge_count = graph.number_of_edges()
    if node_count == 0 or edge_count == 0:
        empty_metrics = {column: 0.0 for column in GRAPH_FEATURE_COLUMNS}
        empty_metrics["graph_node_count"] = float(node_count)
        empty_metrics["graph_edge_count"] = float(edge_count)
        empty_metrics["graph_top_source"] = ""
        empty_metrics["graph_top_destination"] = ""
        empty_metrics["graph_top_host"] = ""
        empty_metrics["graph_top_host_score"] = 0.0
        empty_metrics["graph_top_edge"] = ""
        empty_metrics["graph_top_edge_score"] = 0.0
        empty_metrics["graph_suspicious_hosts"] = ""
        empty_metrics["graph_suspicious_edges"] = ""
        return GraphArtifacts(window_metrics=empty_metrics, host_rows=[], edge_rows=[])

    total_weight = sum(data["weight"] for _, _, data in graph.edges(data=True)) or 1.0
    total_flow_count = float(edge_df["flow_count"].sum()) or 1.0
    out_weight = dict(graph.out_degree(weight="weight"))
    in_weight = dict(graph.in_degree(weight="weight"))
    top_source = max(out_weight, key=out_weight.get)
    top_destination = max(in_weight, key=in_weight.get)

    pagerank = nx.pagerank(graph, weight="weight")
    undirected = graph.to_undirected()
    betweenness = nx.betweenness_centrality(undirected, normalized=True)
    clustering_by_node = nx.clustering(undirected)
    clustering = nx.average_clustering(undirected) if undirected.number_of_nodes() > 1 else 0.0
    components = list(nx.connected_components(undirected)) if undirected.number_of_nodes() else []
    largest_component = max((len(component) for component in components), default=0)

    host_rows: list[dict[str, object]] = []
    node_scores: dict[str, float] = {}
    for node in graph.nodes():
        node_score = (
            0.35 * (out_weight.get(node, 0.0) / total_weight)
            + 0.35 * (in_weight.get(node, 0.0) / total_weight)
            + 0.20 * pagerank.get(node, 0.0)
            + 0.10 * betweenness.get(node, 0.0)
        )
        node_scores[node] = float(node_score)
        host_rows.append(
            {
                "source_file": source_file,
                "window_start": window_start,
                "host": str(node),
                "out_weight": float(out_weight.get(node, 0.0)),
                "in_weight": float(in_weight.get(node, 0.0)),
                "weighted_degree": float(out_weight.get(node, 0.0) + in_weight.get(node, 0.0)),
                "out_weight_share": float(out_weight.get(node, 0.0) / total_weight),
                "in_weight_share": float(in_weight.get(node, 0.0) / total_weight),
                "pagerank": float(pagerank.get(node, 0.0)),
                "betweenness": float(betweenness.get(node, 0.0)),
                "clustering": float(clustering_by_node.get(node, 0.0)),
                "peer_count": int(len(set(graph.predecessors(node)).union(graph.successors(node)))),
                "node_score": float(node_score),
            }
        )

    edge_rows: list[dict[str, object]] = []
    for row in edge_df.itertuples(index=False):
        edge_weight_share = float(row.total_bytes) / total_weight
        edge_flow_share = float(row.flow_count) / total_flow_count
        source_node_score = node_scores.get(row[0], 0.0)
        destination_node_score = node_scores.get(row[1], 0.0)
        edge_score = (
            0.45 * edge_weight_share
            + 0.25 * edge_flow_share
            + 0.15 * source_node_score
            + 0.15 * destination_node_score
        )
        edge_rows.append(
            {
                "source_file": source_file,
                "window_start": window_start,
                "source_ip": str(row[0]),
                "destination_ip": str(row[1]),
                "flow_count": int(row.flow_count),
                "total_bytes": float(row.total_bytes),
                "edge_weight_share": float(edge_weight_share),
                "edge_flow_share": float(edge_flow_share),
                "source_node_score": float(source_node_score),
                "destination_node_score": float(destination_node_score),
                "edge_score": float(edge_score),
            }
        )

    top_host_row = max(host_rows, key=lambda row: row["node_score"])
    top_edge_row = max(edge_rows, key=lambda row: row["edge_score"])
    suspicious_hosts = "; ".join(
        row["host"] for row in sorted(host_rows, key=lambda row: row["node_score"], reverse=True)[:3]
    )
    suspicious_edges = "; ".join(
        f"{row['source_ip']}->{row['destination_ip']}"
        for row in sorted(edge_rows, key=lambda row: row["edge_score"], reverse=True)[:3]
    )

    return GraphArtifacts(
        window_metrics={
            "graph_node_count": float(node_count),
            "graph_edge_count": float(edge_count),
            "graph_density": float(nx.density(graph)),
            "graph_avg_flows_per_edge": float(edge_df["flow_count"].mean()),
            "graph_avg_bytes_per_edge": float(edge_df["total_bytes"].mean()),
            "graph_max_in_weight_share": float(max(in_weight.values()) / total_weight),
            "graph_max_out_weight_share": float(max(out_weight.values()) / total_weight),
            "graph_max_pagerank": float(max(pagerank.values())),
            "graph_max_betweenness": float(max(betweenness.values(), default=0.0)),
            "graph_mean_clustering": float(clustering),
            "graph_largest_component_ratio": float(largest_component / node_count),
            "graph_top_source": str(top_source),
            "graph_top_destination": str(top_destination),
            "graph_top_host": str(top_host_row["host"]),
            "graph_top_host_score": float(top_host_row["node_score"]),
            "graph_top_edge": f"{top_edge_row['source_ip']}->{top_edge_row['destination_ip']}",
            "graph_top_edge_score": float(top_edge_row["edge_score"]),
            "graph_suspicious_hosts": suspicious_hosts,
            "graph_suspicious_edges": suspicious_edges,
        },
        host_rows=host_rows,
        edge_rows=edge_rows,
    )
