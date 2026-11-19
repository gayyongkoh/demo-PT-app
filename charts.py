# charts.py
import matplotlib
matplotlib.use("Agg")          # no display on a server
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import AutoMinorLocator, FormatStrFormatter, MaxNLocator
from io import BytesIO
import math


def score_chart(lab_codes, scores, ylabel, ylim=4, figsize=(4.5, 3.8)):
    """Return a PNG (BytesIO) bar chart of proficiency-test scores."""
    n = len(lab_codes)
    x = list(range(n))
    clipped = [max(min(s, ylim), -ylim) for s in scores]

    fig, ax = plt.subplots(figsize=figsize)
    ax.bar(x, clipped, width=0.8, color="#0000FF", edgecolor="black", linewidth=1.2)
    ax.axhline(0, color="black", linestyle="--", linewidth=1)

    for xi, s, c in zip(x, scores, clipped):
        above = (c >= 0) != (abs(s) > ylim)   # label goes inside the bar if clipped
        ax.annotate(
            f"{s:.1f}", (xi, c),
            xytext=(0, 4 if above else -4), textcoords="offset points",
            ha="center", va="bottom" if above else "top", fontsize=8,
        )

    ax.set_xlim(-0.5, n - 0.5)
    ax.set_ylim(-ylim, ylim)
    ax.set_yticks(range(-ylim, ylim + 1))          # ticks every 1
    ax.yaxis.grid(True, color="#BBBBBB", linewidth=0.6)
    ax.set_axisbelow(True)                          # grid behind the bars
    ax.set_xticks(x)
    ax.set_xticklabels([str(c) for c in lab_codes])
    ax.tick_params(axis="x", length=0)
    ax.tick_params(labelsize=8)
    ax.set_xlabel("Laboratory Code", fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(ylabel, fontweight="bold", fontsize=10)

    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=200)         # no bbox_inches="tight": fixed size
    plt.close(fig)
    buf.seek(0)
    return buf


def result_chart(lab_codes, results, lab_U, av, av_plus, av_minus, mean,
                 ylabel, figsize=(7, 3.4)):
    """Lab results with error bars, assigned value (thick line), AV ± U band and mean."""
    n = len(lab_codes)
    x = list(range(n))
    has_err  = [isinstance(u, (int, float)) for u in lab_U]
    errs     = [u if h else float("nan") for u, h in zip(lab_U, has_err)]  # NaN = no bar
    errs_num = [u if h else 0 for u, h in zip(lab_U, has_err)]             # for axis range

    fig, ax = plt.subplots(figsize=figsize)

    x0, x1 = -0.4, n - 0.6   # reference lines span slightly beyond first/last lab

    # Shaded band: expanded uncertainty of the assigned value
    ax.fill_between([x0, x1], av_minus, av_plus, color="#8C9ED9", alpha=0.8,
                    linewidth=0, zorder=1)
    for y in (av_plus, av_minus):
        ax.hlines(y, x0, x1, color="#2B3A8C", linewidth=0.6, zorder=2)

    # Assigned value (thick) and arithmetic mean (red)
    ax.hlines(av, x0, x1, color="black", linewidth=2, zorder=2)
    if mean is not None:
        ax.hlines(mean, x0, x1, color="#D93A1F", linewidth=1, zorder=2)

    # Lab results with error bars
    ax.errorbar(x, results, yerr=errs, fmt="s", color="black", markersize=5,
                ecolor="black", elinewidth=0.8, capsize=2.5, zorder=4)

    # Y range follows the data; extra room at the bottom for the legend
    lows  = [r - e for r, e in zip(results, errs_num)] + [av_minus] + ([mean] if mean is not None else [])
    highs = [r + e for r, e in zip(results, errs_num)] + [av_plus]  + ([mean] if mean is not None else [])
    span = (max(highs) - min(lows)) or 0.001
    ax.set_ylim(min(lows) - span * 0.55, max(highs) + span * 0.12)

    # Round-number ticks, with decimals derived from the tick spacing
    ax.yaxis.set_major_locator(MaxNLocator(nbins=9, steps=[1, 2, 5, 10]))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ticks = ax.get_yticks()
    step = abs(ticks[1] - ticks[0])
    decimals = max(1, -int(math.floor(math.log10(step))))
    ax.yaxis.set_major_formatter(FormatStrFormatter(f"%.{decimals}f"))

    ax.set_xlim(-0.5, n - 0.5)
    ax.set_xticks(x)
    ax.set_xticklabels([str(c) for c in lab_codes], fontsize=8)
    ax.tick_params(axis="x", length=0)
    ax.tick_params(axis="y", which="both", direction="in", labelsize=8)
    ax.set_xlabel("Laboratory Code", fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)

    legend_items = [
        Line2D([0], [0], color="black", linewidth=2),
        Patch(facecolor="#8C9ED9", edgecolor="#2B3A8C", linewidth=0.6),
        Line2D([0], [0], color="#D93A1F", linewidth=1),
    ]
    ax.legend(legend_items,
              ["Assigned Value",
               "Expanded Uncertainty of Assigned Value",
               "Arithmetic Mean of Reported Results"],
              title="Legend", title_fontsize=8, fontsize=7,
              loc="lower right", frameon=True, fancybox=False,
              edgecolor="black", framealpha=1, alignment="left")

    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf