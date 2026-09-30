# -*- coding: utf-8 -*-
"""
verify_sheet.py - render a Visio stencil set (.vssx) as a labeled contact
sheet PNG: every master is dropped onto a grid page and exported.

Usage:
    python verify_sheet.py icons.vssx [output.png] [--cols N] [--cell IN] [--labels]

Defaults: output = <vssx name>_sheet.png next to the vssx, 10 columns,
1 inch cells, name labels on.
"""
import argparse
import math
import os
import sys

import win32com.client


def main():
    ap = argparse.ArgumentParser(
        description="Render a Visio stencil set (.vssx) as a contact sheet PNG.")
    ap.add_argument("vssx", help="stencil set file (.vssx)")
    ap.add_argument("output", nargs="?", default=None,
                    help="output PNG path (default: <vssx>_sheet.png beside it)")
    ap.add_argument("--cols", type=int, default=10, help="grid columns (default 10)")
    ap.add_argument("--cell", type=float, default=1.0,
                    help="cell size in inches (default 1.0)")
    ap.add_argument("--labels", dest="labels", action="store_true", default=True,
                    help="draw master names under each icon (default: on)")
    ap.add_argument("--no-labels", dest="labels", action="store_false",
                    help="skip master names")
    args = ap.parse_args()

    vssx = os.path.abspath(args.vssx)
    if not os.path.isfile(vssx):
        sys.exit(f"not found: {vssx}")
    out = os.path.abspath(args.output) if args.output else \
        os.path.splitext(vssx)[0] + "_sheet.png"
    out_dir = os.path.dirname(out)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    app = win32com.client.DispatchEx("Visio.InvisibleApp")
    app.AlertResponse = 2  # auto-cancel any modal dialog
    try:
        doc = app.Documents.Add("")
        pg = doc.Pages(1)
        st = app.Documents.OpenEx(vssx, 4)  # 4 = visOpenDocked
        n = st.Masters.Count
        cols = max(1, args.cols)
        cell = max(0.5, args.cell)
        rows = math.ceil(n / cols)
        pg.PageSheet.Cells("PageWidth").FormulaU = f"{cols * cell + 1} in"
        pg.PageSheet.Cells("PageHeight").FormulaU = f"{rows * cell + 1} in"

        labels = []  # (name, x_in, y_in) for the PNG annotation pass
        for i in range(n):
            m = st.Masters.Item(i + 1)
            cx = 0.5 + (i % cols) * cell + cell / 2
            cy = 0.5 + (i // cols) * cell + cell / 2
            shp = pg.Drop(m, cx, cy)
            w = shp.Cells("Width").ResultIU
            h = shp.Cells("Height").ResultIU
            s = min((cell * 0.7) / w if w else 1,
                    (cell * 0.7) / h if h else 1)
            if s < 1:
                shp.Cells("Width").FormulaU = f"{w * s:.4f} in"
                shp.Cells("Height").FormulaU = f"{h * s:.4f} in"
            if args.labels:
                labels.append((m.Name, cx, 0.5 + (i // cols) * cell + 0.92))

        pg.Export(out)
        st.Close()
        doc.Close()
        print(f"exported {n} masters -> {out}")
    finally:
        app.Quit()

    # annotate the exported PNG with master names (Pillow, optional)
    if args.labels and labels:
        try:
            from PIL import Image, ImageDraw, ImageFont
        except ImportError:
            print("Pillow not installed - skipping name labels")
            return
        im = Image.open(out)
        d = ImageDraw.Draw(im)
        try:
            font = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 13)
        except Exception:
            font = ImageFont.load_default()
        page_w = cols * cell + 1.0
        page_h = rows * cell + 1.0
        sx = im.width / page_w
        sy = im.height / page_h
        for name, cx, cy in labels:
            d.text((cx * sx - 40, cy * sy - 7), name[:18], fill="black",
                   font=font)
        im.save(out)
        print("labels drawn")


if __name__ == "__main__":
    main()
