# -*- coding: utf-8 -*-
"""
drawio2stencils.py - extract every icon from a draw.io file into a Visio
stencil set (.vssx) with true-vector masters.

Usage:
    python drawio2stencils.py input.drawio output.vssx

The converter groups the stacked cells that form each icon (white backdrop +
dark compound + white cutout overlays), names each group from the label cell
below it, de-duplicates identical groups, then draws each one as native
Visio geometry (polylines with band-winding holes) into stencil masters.
"""
import base64
import hashlib
import math
import os
import re
import sys
import urllib.parse
import shutil
import zlib

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# reuse the converter's parsing + painting engine
import drawio2visio as d2v  # noqa: E402

PPI = 96.0
SRC = os.path.join(HERE, "src.drawio")
OUT = os.path.join(HERE, "sah_stencil_vector.vssx")
MASTER_H = 0.85  # inches, uniform master height


# --------------------------------------------------------------- clustering

def collect(cells, offsets):
    """Icon candidate cells (stencils of icon-ish size) + label cells."""
    icon_cells, labels = [], []
    for c in cells:
        if not c["vertex"]:
            continue
        r = d2v.rect_of(c, offsets)
        if not r:
            continue
        x, y, w, h = r
        mm = re.search(r"stencil\((.+?)\)", c["style"], re.S)
        if mm and 12 <= w <= 140 and 12 <= h <= 140:
            sig = hashlib.md5(mm.group(1).encode()).hexdigest()[:10]
            style = d2v.parse_style(c["style"])
            icon_cells.append({"id": c["id"], "rect": r, "sig": sig,
                               "fill": style.get("fillColor"),
                               "stroke": style.get("strokeColor"),
                               "payload": mm.group(1)})
        txt = re.sub(r"<[^>]+>", " ", c["label"] or "").strip()
        if c["label"] and 14 <= h <= 30 and w >= 30 and txt:
            labels.append({"id": c["id"], "rect": r, "text": txt})
    return icon_cells, labels


def overlap(a, b, pad=2):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return not (ax + aw + pad < bx or bx + bw + pad < ax or
                ay + ah + pad < by or by + bh + pad < ay)


def clusters_of(icon_cells):
    parent = list(range(len(icon_cells)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(len(icon_cells)):
        for j in range(i + 1, len(icon_cells)):
            if overlap(icon_cells[i]["rect"], icon_cells[j]["rect"]):
                parent[find(i)] = find(j)
    groups = {}
    for i, ic in enumerate(icon_cells):
        groups.setdefault(find(i), []).append(ic)
    return list(groups.values())


def dedupe(clusters):
    uniq = {}
    for cl in clusters:
        key = tuple(sorted(
            (c["sig"], str(c["fill"]), str(c["stroke"])) for c in cl))
        uniq.setdefault(key, cl)
    return list(uniq.values())


def name_cluster(cl, labels):
    """Nearest label below the icon; join multi-line labels that horizontally
    overlap the first one (stacked lines of the same label)."""
    x0 = min(c["rect"][0] for c in cl)
    y1 = max(c["rect"][1] + c["rect"][3] for c in cl)
    bcx = (x0 + max(c["rect"][0] + c["rect"][2] for c in cl)) / 2.0
    cands = []
    for lb in labels:
        lx, ly, lw, lh = lb["rect"]
        if abs(lx + lw / 2 - bcx) > max(70, lw / 2):
            continue
        gap = ly - y1
        if -5 <= gap <= 60:
            cands.append((gap, lx, lw, lb))
    if not cands:
        return None
    cands.sort()
    first = cands[0]
    fx0, fx1 = first[1], first[1] + first[2]
    parts = [first[3]["text"]]
    for gap, lx, lw, lb in cands[1:]:
        ov = min(fx1, lx + lw) - max(fx0, lx)
        if ov > 0.5 * min(lw, first[2]):
            parts.append(lb["text"])
    return " ".join(parts)


# --------------------------------------------------------------- master draw

def draw_cluster_into(editor, cl, cells, offsets):
    """Draw one icon cluster's cells (big-containers-first order) into a
    master editor page, normalised to the master box."""
    x0 = min(c["rect"][0] for c in cl)
    y0 = min(c["rect"][1] for c in cl)
    x1 = max(c["rect"][0] + c["rect"][2] for c in cl)
    y1 = max(c["rect"][1] + c["rect"][3] for c in cl)
    bw, bh = x1 - x0, y1 - y0
    if bw <= 0 or bh <= 0:
        return
    H_in = MASTER_H / PPI * PPI / (bh / PPI)  # = MASTER_H
    H_in = MASTER_H
    W_in = MASTER_H * (bw / bh)
    scale = MASTER_H / bh  # canvas px -> master inches

    def X(px):
        return (px - x0) * scale

    def Y(py):
        return H_in - (py - y0) * scale

    def rgb(v):
        return d2v.VisioBuilder.rgb(v)

    def paint(shp, fill, stroke, sw, dashed, alpha=1.0):
        if fill:
            shp.Cells("FillForegnd").FormulaU = rgb(fill)
            if alpha < 1:
                shp.Cells("FillTrans").FormulaU = \
                    f"{(1 - alpha) * 100:.0f} PERCENT"
        else:
            shp.Cells("FillPattern").FormulaU = "0"
        if stroke:
            shp.Cells("LineColor").FormulaU = rgb(stroke)
            shp.Cells("LineWeight").FormulaU = \
                f"{max(sw * 0.75, 0.5):.2f} pt"
        else:
            shp.Cells("LinePattern").FormulaU = "0"
        if dashed:
            shp.Cells("LinePattern").FormulaU = "2"

    def poly_shape(pts):
        arr = []
        for px, py in pts:
            arr.append(X(px))
            arr.append(Y(py))
        return editor.DrawPolyline(arr, 0)

    def paint_compound(subpaths, pfill, pstroke):
        n = len(subpaths)

        def signed_area(pts):
            s = 0.0
            for i in range(len(pts) - 1):
                s += pts[i][0] * pts[i + 1][1] - pts[i + 1][0] * pts[i][1]
            return s / 2.0

        def point_in_poly(pt, poly):
            x, y = pt
            inside = False
            j = len(poly) - 1
            for i in range(len(poly)):
                xi, yi = poly[i]
                xj, yj = poly[j]
                if ((yi > y) != (yj > y)) and \
                        (x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi):
                    inside = not inside
                j = i
            return inside

        def centroid(pts):
            return (sum(q[0] for q in pts) / len(pts),
                    sum(q[1] for q in pts) / len(pts))

        signs = [1 if signed_area(sp) > 0 else -1 for sp in subpaths]
        cents = [centroid(sp) for sp in subpaths]

        def band_state(i):
            pts = subpaths[i]
            bl, bm, bn = -1.0, None, None
            for k in range(len(pts) - 1):
                ax, ay = pts[k]
                bx, by = pts[k + 1]
                L = math.hypot(bx - ax, by - ay)
                if L > bl:
                    bl = L
                    mx, my = (ax + bx) / 2.0, (ay + by) / 2.0
                    nx, ny = -(by - ay) / L, (bx - ax) / L
                    if signs[i] < 0:
                        nx, ny = -nx, -ny
                    bm, bn = (mx, my), (nx, ny)
            eps = max(1e-3, bl * 1e-3)
            ip = (bm[0] + bn[0] * eps, bm[1] + bn[1] * eps)
            if not point_in_poly(ip, pts):
                ip = (bm[0] - bn[0] * eps, bm[1] - bn[1] * eps)
            w = signs[i]
            for j in range(n):
                if j != i and point_in_poly(ip, subpaths[j]):
                    w += signs[j]
            return w

        wind = [band_state(i) for i in range(n)]

        def depth_of(i):
            return sum(1 for j in range(n) if j != i and
                       point_in_poly(cents[i], subpaths[j]))

        plan = []
        for i in range(n):
            colour = pfill if wind[i] != 0 else "#FFFFFF"
            plan.append((depth_of(i), i, colour))
        plan.sort(key=lambda t: t[0])

        body_i = max(range(n), key=lambda k: abs(signed_area(subpaths[k])))
        body = subpaths[body_i]
        shp = poly_shape(body)
        paint(shp, pfill, pstroke, 1.0, False)

        for _, i, colour in plan:
            if i == body_i:
                continue
            hs = poly_shape(subpaths[i])
            hs.Cells("FillForegnd").FormulaU = rgb(colour)
            if pstroke:
                hs.Cells("LineColor").FormulaU = rgb(pstroke)
                hs.Cells("LineWeight").FormulaU = "0.5 pt"
            else:
                hs.Cells("LinePattern").FormulaU = "0"

    by_id = {c["id"]: c for c in cells}
    for member in cl:
        cell = by_id[member["id"]]
        style = d2v.parse_style(cell["style"])
        fill = style.get("fillColor")
        if fill == "none":
            fill = None
        stroke = style.get("strokeColor")
        if stroke == "none":
            stroke = None
        sw = float(style.get("strokeWidth", 1) or 1)
        dashed = bool(style.get("dashed"))
        cx, cy, cw, ch = member["rect"]
        mm = re.search(r"stencil\((.+?)\)", cell["style"], re.S)
        if not mm:
            continue
        sel = d2v.decode_stencil(mm.group(1))
        prims, (sw0, sh0) = d2v.stencil_primitives(sel, fill, stroke)
        sxx = cw / sw0
        syy = ch / sh0

        def map_pts(sp):
            return [(cx + px * sxx, cy + py * syy) for px, py in sp]

        for p in prims:
            alpha = p.get("alpha", 1.0)
            if p["kind"] == "poly":
                shp = poly_shape(map_pts(p["subpaths"][0]))
                paint(shp, p["fill"], p["stroke"], sw, dashed, alpha)
            elif p["kind"] == "compound":
                paint_compound([map_pts(sp) for sp in p["subpaths"]],
                               p["fill"], p["stroke"])
            elif p["kind"] == "rect":
                rx, ry = cx + p["x"] * sxx, cy + p["y"] * syy
                shp = editor.DrawRectangle(X(rx), Y(ry + p["h"] * syy),
                                           X(rx + p["w"] * sxx), Y(ry))
                paint(shp, p["fill"], p["stroke"], sw, dashed, alpha)
            elif p["kind"] == "oval":
                rx, ry = cx + p["x"] * sxx, cy + p["y"] * syy
                shp = editor.DrawOval(X(rx), Y(ry + p["h"] * syy),
                                      X(rx + p["w"] * sxx), Y(ry))
                paint(shp, p["fill"], p["stroke"], sw, dashed, alpha)


# ------------------------------------------------------------------- main

def main(src=SRC, out=OUT):
    cells = d2v.load_cells(src)
    offsets = d2v.build_offsets(cells)
    icon_cells, labels = collect(cells, offsets)
    clusters = clusters_of(icon_cells)
    uniq = dedupe(clusters)
    print(f"icon cells={len(icon_cells)} clusters={len(clusters)} "
          f"unique={len(uniq)}")

    app = win32com.client.DispatchEx("Visio.InvisibleApp")
    app.AlertResponse = 2
    try:
        doc = app.Documents.AddEx("", 4, 0, 0)  # stencil document
        doc.Title = os.path.splitext(os.path.basename(src))[0] + " icons"
        used_names = {}

        for idx, cl in enumerate(uniq):
            # per-cluster z-order: bigger containers first
            cl_sorted = sorted(
                cl, key=lambda mm: -(mm["rect"][2] * mm["rect"][3]))

            master = doc.Masters.Add()
            name = name_cluster(cl, labels) or f"Icon {idx + 1}"
            base = name
            k = 2
            while name in used_names:
                name = f"{base} {k}"
                k += 1
            used_names[name] = True
            master.Name = name

            editor = master.Open()
            try:
                draw_cluster_into(editor, cl_sorted, cells, offsets)
            finally:
                editor.Close()

        doc.SaveAs(os.path.abspath(out))
        print(f"masters={doc.Masters.Count}")
        doc.Close()
    finally:
        app.Quit()

    # Post-process: Visio saves stencil docs with a windows.xml that has no
    # Stencil window, so double-clicking the .vssx opens an EMPTY drawing
    # with "There are no stencils open". Inject a docked Stencil window
    # pointing at the file itself so it opens correctly for the user.
    try:
        _inject_stencil_window(os.path.abspath(out), os.path.basename(out))
    except PermissionError:
        # target locked (open in Visio/Explorer preview) - save beside it
        alt = os.path.splitext(out)[0] + ".fixed.vssx"
        shutil.move(out + ".tmp", alt)
        _inject_stencil_window(alt, os.path.basename(alt))
        print(f"WARNING: {out} was locked - wrote {alt} instead")
        out = alt
    print("saved:", out)


def _inject_stencil_window(path, own_name):
    """Make the .vssx open like Microsoft's own stencils: double-click shows
    a blank drawing with the stencil DOCKED LEFT in the Shapes pane.
    Discovered by A/B test against ANALYTICS_U.vssx: MS ships an EMPTY
    <Windows/> element. Any Drawing/Stencil window declaration in a stencil's
    windows.xml makes Visio open the file as a read-only drawing with
    'There are no stencils open' instead of docking the panel."""
    import zipfile
    tmp = path + ".tmp"
    zin = zipfile.ZipFile(path)
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "visio/windows.xml":
            s = ("<?xml version='1.0' encoding='utf-8' ?>\r\n"
                 "<Windows ClientWidth='0' ClientHeight='0' "
                 "xmlns='http://schemas.microsoft.com/office/visio/2012/main' "
                 "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006"
                 "/relationships' xml:space='preserve'/>")
            data = s.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    zin.close()
    os.replace(tmp, path)


if __name__ == "__main__":
    _src = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else SRC
    _out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else OUT
    main(_src, _out)
