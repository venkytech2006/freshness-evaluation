"""Generate figures for the FreshGuard project report and presentation."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import os

OUT = os.path.join(os.path.dirname(__file__), "figures")
os.makedirs(OUT, exist_ok=True)

NAVY = "#1E2761"
BLUE = "#2b6cb0"
TEAL = "#2c7a7b"
AMBER = "#b7791f"
GREEN = "#2f855a"
GREY = "#4a5568"
LIGHT = "#edf2f7"


def box(ax, x, y, w, h, text, fc, ec=None, tc="white", fs=9, rounded=0.02):
    ec = ec or fc
    p = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0.002,rounding_size={rounded}",
                       linewidth=1.2, edgecolor=ec, facecolor=fc, mutation_aspect=1)
    ax.add_patch(p)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            color=tc, fontsize=fs, wrap=True, zorder=5)


def arrow(ax, x1, y1, x2, y2, color=GREY, style="-|>"):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style,
                                 mutation_scale=12, linewidth=1.4, color=color, zorder=1))


# ---------------------------------------------------------------- architecture
def fig_architecture():
    fig, ax = plt.subplots(figsize=(9.2, 5.2))
    ax.set_xlim(0, 100); ax.set_ylim(0, 60); ax.axis("off")

    box(ax, 2, 46, 20, 10, "Web UI\n(Ask / History / Admin)", NAVY, fs=9)
    box(ax, 2, 30, 20, 10, "FastAPI backend\n/query  /feedback\n/history  /admin  /eval", GREY, fs=8)

    # pipeline stages
    stages = [
        (26, 50, "Preprocess\n(clean / classify)", BLUE),
        (26, 36, "Freshness gate\nLogisticRegression\nP(stale)", TEAL),
        (26, 22, "Decision policy\nretrieve if stale\nor low confidence", TEAL),
        (52, 50, "Retrieval\nFAISS + TF-IDF\nrelevance x recency", BLUE),
        (52, 36, "Live ingestion (KB miss)\nTavily web + Grok\nembed -> persist", AMBER),
        (52, 22, "Generation\ngrounded / Grok\ncitation-aware", BLUE),
        (78, 50, "Verify statements\nsupported / unsupported\n/ contradicted", GREEN),
        (78, 36, "Correct & flag\ndrop unsupported", GREEN),
        (78, 22, "Confidence + freshness\n+ cited sources", NAVY),
    ]
    for x, y, t, c in stages:
        box(ax, x, y, 20, 10, t, c, fs=7.5)

    # flow arrows
    arrow(ax, 22, 51, 26, 55)                 # backend -> preprocess
    arrow(ax, 36, 50, 36, 46)                 # preprocess -> freshness
    arrow(ax, 36, 36, 36, 32)                 # freshness -> decision
    arrow(ax, 46, 27, 52, 55)                 # decision -> retrieval
    arrow(ax, 62, 50, 62, 46)                 # retrieval -> live
    arrow(ax, 62, 36, 62, 32)                 # live -> generation
    arrow(ax, 72, 27, 78, 55)                 # generation -> verify
    arrow(ax, 88, 50, 88, 46)                 # verify -> correct
    arrow(ax, 88, 36, 88, 32)                 # correct -> confidence

    box(ax, 26, 6, 20, 9, "FAISS vector store\n(TF-IDF embeddings)", "#2d3748", fs=7.5)
    box(ax, 52, 6, 20, 9, "SQLite\nhistory / feedback\nconfig / KB", "#2d3748", fs=7.5)
    box(ax, 78, 6, 20, 9, "Learned estimator\n(joblib artifact)", "#2d3748", fs=7.5)
    arrow(ax, 36, 22, 36, 15, color="#a0aec0")
    arrow(ax, 62, 22, 62, 15, color="#a0aec0")
    arrow(ax, 36, 36, 30, 15, color="#a0aec0")

    arrow(ax, 12, 46, 12, 40)                 # UI -> backend
    arrow(ax, 22, 35, 26, 40, color="#a0aec0")
    plt.tight_layout()
    fig.savefig(os.path.join(OUT, "architecture.png"), dpi=170, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- gate flow
def fig_gate():
    fig, ax = plt.subplots(figsize=(8.8, 3.4))
    ax.set_xlim(0, 100); ax.set_ylim(0, 30); ax.axis("off")
    box(ax, 1, 10, 16, 10, "User query", NAVY, fs=9)
    box(ax, 20, 10, 18, 10, "Freshness score\nf = P(stale)\nmodel conf. c", TEAL, fs=8)
    box(ax, 41, 10, 20, 10, "Gate\nretrieve if\nf >= 0.55 or\nc < 0.60", AMBER, fs=8)
    box(ax, 64, 19, 17, 9, "Retrieve + verify\n(stale / time-sensitive)", BLUE, fs=7.5)
    box(ax, 64, 1, 17, 9, "Answer directly\n(fresh / static)", GREEN, fs=7.5)
    box(ax, 84, 10, 15, 10, "Confidence\n+ freshness\n+ sources", NAVY, fs=8)
    arrow(ax, 17, 15, 20, 15)
    arrow(ax, 38, 15, 41, 15)
    arrow(ax, 61, 14, 64, 23)
    arrow(ax, 61, 14, 64, 6)
    arrow(ax, 81, 23, 90, 20)
    arrow(ax, 81, 6, 90, 10)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT, "gate.png"), dpi=170, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- live ingest
def fig_live():
    fig, ax = plt.subplots(figsize=(8.8, 3.2))
    ax.set_xlim(0, 100); ax.set_ylim(0, 28); ax.axis("off")
    box(ax, 1, 9, 15, 10, "Retrieve\nfrom local KB", BLUE, fs=8)
    box(ax, 19, 9, 17, 10, "KB miss?\n(no evidence or\nyear mismatch)", AMBER, fs=8)
    box(ax, 40, 18, 17, 9, "Tavily web search", TEAL, fs=8)
    box(ax, 40, 1, 17, 9, "Grok LLM", TEAL, fs=8)
    box(ax, 60, 9, 16, 10, "Embed + persist\ninto FAISS / KB", GREEN, fs=8)
    box(ax, 79, 9, 19, 10, "Re-retrieve +\nverify + cite", NAVY, fs=8)
    arrow(ax, 16, 14, 19, 14)
    arrow(ax, 36, 14, 40, 22)
    arrow(ax, 36, 14, 40, 6)
    arrow(ax, 57, 22, 60, 15)
    arrow(ax, 57, 6, 60, 13)
    arrow(ax, 76, 14, 79, 14)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT, "live_ingest.png"), dpi=170, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- metrics
def fig_metrics():
    fig, ax = plt.subplots(figsize=(7.6, 3.8))
    labels = ["Freshness\ndecision acc.", "Factuality", "Halluc.\nprecision",
              "Halluc.\nrecall", "Halluc.\nF1", "Calibration\n(1 - ECE)"]
    vals = [1.00, 1.00, 1.00, 1.00, 1.00, 0.90]
    colors = [TEAL, BLUE, GREEN, GREEN, GREEN, AMBER]
    bars = ax.bar(labels, vals, color=colors, width=0.62)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.02, f"{v:.2f}",
                ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_ylim(0, 1.15)
    ax.set_ylabel("Score")
    ax.axhline(1.0, color="#cbd5e0", linewidth=0.8, linestyle="--")
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    ax.tick_params(labelsize=8)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT, "metrics.png"), dpi=170, bbox_inches="tight")
    plt.close(fig)


# -------------------------------------------------- future-event before/after
def fig_future():
    fig, ax = plt.subplots(figsize=(8.6, 3.0))
    ax.axis("off")
    ax.text(0.25, 0.92, "Query: \"Who won the 2026 FIFA World Cup?\"",
            ha="center", fontsize=10, fontweight="bold", transform=ax.transAxes)
    box(ax2 := ax, 0, 0, 0, 0, "", "white")  # noop to keep api
    # left = before, right = after (manual text panels)
    ax.add_patch(FancyBboxPatch((0.02, 0.08), 0.45, 0.68, boxstyle="round,pad=0.01,rounding_size=0.02",
                 transform=ax.transAxes, facecolor="#fff5f5", edgecolor="#c53030", linewidth=1.3))
    ax.add_patch(FancyBboxPatch((0.53, 0.08), 0.45, 0.68, boxstyle="round,pad=0.01,rounding_size=0.02",
                 transform=ax.transAxes, facecolor="#f0fff4", edgecolor="#2f855a", linewidth=1.3))
    ax.text(0.245, 0.70, "Before (stale KB only)", ha="center", color="#c53030",
            fontsize=9.5, fontweight="bold", transform=ax.transAxes)
    ax.text(0.245, 0.45,
            "Wrongly answers from the\n2022 World Cup document\n(Argentina) — a stale,\nyear-mismatched source.",
            ha="center", va="center", fontsize=8.5, transform=ax.transAxes)
    ax.text(0.755, 0.70, "After (live ingestion)", ha="center", color="#2f855a",
            fontsize=9.5, fontweight="bold", transform=ax.transAxes)
    ax.text(0.755, 0.45,
            "KB miss → Tavily + Grok →\nembed + verify → 2026 answer\nwith citations, or an honest\n\"not available\" if no key.",
            ha="center", va="center", fontsize=8.5, transform=ax.transAxes)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT, "future_event.png"), dpi=170, bbox_inches="tight")
    plt.close(fig)


fig_architecture()
fig_gate()
fig_live()
fig_metrics()
fig_future()
print("figures written to", OUT)
for f in sorted(os.listdir(OUT)):
    print(" -", f)
