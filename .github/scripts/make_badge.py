#!/usr/bin/env python
"""Generate a flat SVG coverage badge from a coverage.py JSON report.

Usage: python make_badge.py coverage.json docs/coverage.svg

Dependency-free: reads ``totals.percent_covered`` from the coverage JSON
and renders a two-cell flat badge (label + percentage) with a colour that
follows common thresholds (red < 50, orange < 70, yellowgreen < 90,
brightgreen >= 90).
"""

import json
import sys


def badge_svg(label: str, value: str, color: str) -> str:
    label_w = 70
    value_w = 8 * len(value) + 16
    total = label_w + value_w
    return f'''<svg xmlns="http://www.w3.org/2000/svg" width="{total}" height="20"
     role="img" aria-label="{label}: {value}">
  <linearGradient id="s" x2="0" y2="100%">
    <stop offset="0" stop-color="#bbb" stop-opacity=".1"/>
    <stop offset="1" stop-opacity=".1"/>
  </linearGradient>
  <clipPath id="r">
    <rect width="{total}" height="20" rx="3" fill="#fff"/>
  </clipPath>
  <g clip-path="url(#r)">
    <rect width="{label_w}" height="20" fill="#555"/>
    <rect x="{label_w}" width="{value_w}" height="20" fill="{color}"/>
    <rect width="{total}" height="20" fill="url(#s)"/>
  </g>
  <g fill="#fff" text-anchor="middle"
     font-family="Verdana,Geneva,DejaVu Sans,sans-serif"
     font-size="11">
    <text x="{label_w / 2}" y="14">{label}</text>
    <text x="{label_w + value_w / 2}" y="14">{value}</text>
  </g>
</svg>
'''


def color_for(pct: float) -> str:
    if pct < 50:
        return "#e05d44"
    if pct < 70:
        return "#dfb317"
    if pct < 90:
        return "#a4a61d"
    return "#4c1"


def main() -> None:
    if len(sys.argv) != 3:
        print("usage: make_badge.py coverage.json out.svg", file=sys.stderr)
        raise SystemExit(2)
    with open(sys.argv[1], encoding="utf-8") as fh:
        data = json.load(fh)
    pct = round(data["totals"]["percent_covered"], 1)
    svg = badge_svg("coverage", f"{pct}%", color_for(pct))
    with open(sys.argv[2], "w", encoding="utf-8") as fh:
        fh.write(svg)
    print(f"coverage badge written: {pct}%")


if __name__ == "__main__":
    main()
