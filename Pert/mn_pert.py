from pathlib import Path
import re

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D


# -----------------------------
# LOAD DATA
# -----------------------------
def load_tasks(filename):
    df = pd.read_csv(
        filename,
        dtype={
            "ID": str,
            "Dependencias": str,
        },
        sep="|"
    )

    print(df)

    df = df.dropna(subset=["Titulo"]).copy()
    df["Dependencias"] = df["Dependencias"].fillna("").astype(str)

    return df


# -----------------------------
# BUILD GRAPH
# -----------------------------
def build_graph(df):
    G = nx.DiGraph()

    # Add nodes
    for _, row in df.iterrows():
        G.add_node(
            row["ID"],
            title=row["Titulo"],
            duration=float(row.get("Expected", 0) or 0),
        )

    # Add edges
    for _, row in df.iterrows():
        task_id = row["ID"]
        deps_raw = str(row["Dependencias"]).strip()

        if not deps_raw:
            continue

        # robust split: comma, semicolon or pipe
        deps = [d.strip() for d in re.split(r"[,]", deps_raw) if d.strip()]

        print(f"{task_id=}")
        print(f"{deps=}")
        for dep in deps:
            if dep not in G.nodes:
                print(f"[WARN] '{task_id}' depends on missing task '{dep}'")
                continue

            G.add_edge(dep, task_id)
    print(G)

    for layer, nodes in enumerate(nx.topological_generations(G)):
        # `multipartite_layout` expects the layer as a node attribute, so add the
        # numeric layer value as a node attribute
        for node in nodes:
            G.nodes[node]["layer"] = layer
    return G

def draw_graph(G, df):
    fig, ax = plt.subplots(figsize=(16, 9))

    pos = nx.multipartite_layout(G,subset_key="layer")

    nx.draw_networkx(G, pos=pos, ax=ax)
    ax.set_title("PERT Diagram", fontsize=18, pad=20)
    ax.axis("off")

    # Legend
    handles = []
    legend_labels = []

    for _, row in df.sort_values("ID").iterrows():
        handles.append(
            Line2D(
                [0], [0],
                marker=f"${row['ID']}$",
                linestyle="none",
                color="black",
                markersize=12,
            )
        )
        legend_labels.append(f"{row['ID']} - {row['Titulo']}")

    ax.legend(
        handles,
        legend_labels,
        title="Activities",
        loc="upper center",
        bbox_to_anchor=(0.5, -0.15),
        ncol=3,
        frameon=False,
    )

    plt.subplots_adjust(bottom=0.30)

    Path("charts").mkdir(parents=True, exist_ok=True)

    plt.savefig(
        "charts/pert.png",
        dpi=300,
        bbox_inches="tight",
    )

    plt.show()


# -----------------------------
# MAIN
# -----------------------------
def main():
    df = load_tasks("noms.csv")

    G = build_graph(df)

    if not nx.is_directed_acyclic_graph(G):
        cycles = list(nx.simple_cycles(G))
        raise ValueError(f"Dependency graph contains cycles: {cycles}")

    draw_graph(G, df)


if __name__ == "__main__":
    main()
