"""matplotlib bar charts for the BHS Performance Review (spec section 9).

Single-series bars, so no legend - the title names the series. Faults are brand navy,
lost bags brand cyan; value labels and axis text stay in ink grey, not the bar colour.
"""
from __future__ import annotations

import textwrap
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import matplotlib.ticker  # noqa: E402,F401

from . import metrics  # noqa: E402
from .stats import BhsStats  # noqa: E402

NAVY = "#17406D"
CYAN = "#009DD9"
INK = "#333333"
MUTED = "#767676"
TOP_FAULTS = 10


def _bar(labels: list[str], values: list[int], colour: str, title: str, ylabel: str,
         out: Path, rotate: bool = True) -> Path:
    fig, ax = plt.subplots(figsize=(9, 4.6), dpi=200)
    xs = range(len(values))
    bars = ax.bar(xs, values, color=colour, width=0.72, edgecolor="white", linewidth=1.5,
                  zorder=2)
    top = max(values) if values else 0
    for bar, v in zip(bars, values):
        ax.annotate(f"{v:,}", (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    xytext=(0, 3), textcoords="offset points", ha="center", va="bottom",
                    fontsize=8.5, color=INK)
    wrapped = ["\n".join(textwrap.wrap(lbl, 16)) or lbl for lbl in labels]
    ax.set_xticks(list(xs))
    ax.set_xticklabels(wrapped, rotation=35 if rotate else 0,
                       ha="right" if rotate else "center", fontsize=8.5, color=INK)
    ax.set_ylabel(ylabel, fontsize=9, color=MUTED)
    ax.set_title(title, fontsize=11.5, color="#0B2545", fontweight="bold", loc="left", pad=10)
    ax.set_ylim(0, top * 1.15 if top else 1)
    ax.tick_params(axis="y", labelsize=8.5, colors=MUTED)
    ax.tick_params(axis="x", length=0)
    ax.yaxis.grid(True, color="#E6E6E6", linewidth=0.8, zorder=0)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color("#BBBBBB")
    fig.tight_layout()
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, facecolor="white")
    plt.close(fig)
    return out


def fault_chart(st: BhsStats, out_dir: Path) -> Path | None:
    top = st.faults[:TOP_FAULTS]
    if not top:
        return None
    return _bar([k for k, _ in top], [v for _, v in top], NAVY,
                f"{st.site} — Top Fault Categories ({metrics.short_range(st)})",
                "Event count", Path(out_dir) / f"{st.site}_faults.png")


def lost_chart(st: BhsStats, out_dir: Path) -> Path | None:
    data = sorted(st.lost_breakdown, key=lambda kv: -kv[1])
    if not data or not any(v for _, v in data):
        return None
    return _bar([k for k, _ in data], [v for _, v in data], CYAN,
                f"{st.site} — Lost-bag Events by {st.lost_granularity} "
                f"({metrics.short_range(st)})",
                "Lost-bag events", Path(out_dir) / f"{st.site}_lost.png")


def trend_chart(st: BhsStats, out_dir: Path) -> Path | None:
    if not st.trend:
        return None
    return _bar([f"{d:%a} {d.day}" for d, _ in st.trend], [v for _, v in st.trend], NAVY,
                f"{st.site} — Daily Main-line Throughput ({metrics.short_range(st)})",
                "Bags", Path(out_dir) / f"{st.site}_trend.png", rotate=False)
