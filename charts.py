# charts.py
import matplotlib
matplotlib.use("Agg")          # no display on a server
import matplotlib.pyplot as plt
from io import BytesIO


def score_chart(lab_codes, scores, ylabel, ylim=4):
    """Return a PNG (BytesIO) bar chart of proficiency-test scores."""
    n = len(lab_codes)
    x = list(range(n))
    clipped = [max(min(s, ylim), -ylim) for s in scores]

    fig, ax = plt.subplots(figsize=(7, 3.2))
    ax.bar(x, clipped, width=0.8, color="#0000FF", edgecolor="black", linewidth=1.2)
    ax.axhline(0, color="black", linestyle="--", linewidth=1)

    for xi, s, c in zip(x, scores, clipped):
        above = (c >= 0) != (abs(s) > ylim)   # label goes inside the bar if clipped
        ax.annotate(
            f"{s:.1f}", (xi, c),
            xytext=(0, 4 if above else -4), textcoords="offset points",
            ha="center", va="bottom" if above else "top", fontsize=10,
        )

    ax.set_xlim(-0.5, n - 0.5)
    ax.set_ylim(-ylim, ylim)
#    ax.set_yticks([-ylim, 0, ylim])
    ax.set_yticks(range(-ylim, ylim + 1))          # ticks at -4, -3, ... 3, 4
    ax.yaxis.grid(True, color="#BBBBBB", linewidth=0.6)
    ax.set_axisbelow(True)                          # draw grid behind the bars
    ax.set_xticks(x)
    ax.set_xticklabels([str(c) for c in lab_codes])
    ax.tick_params(axis="x", length=0)
    ax.set_xlabel("Laboratory Code")
    ax.set_ylabel(ylabel)
    ax.set_title(ylabel, fontweight="bold", fontsize=11)

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf