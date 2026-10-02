# -*- coding: utf-8 -*-
"""
drawio2visio.py - convert any draw.io (.drawio) file to a native, fully
editable Visio .vsdx (true-vector shapes, no raster/images).

Usage:
    python drawio2visio.py input.drawio output.vsdx

Pipeline: parse mxGraph XML (mxCell + UserObject) -> decode embedded stencil
payloads (base64+raw-deflate of URL-encoded XML) -> primitives (poly /
compound with nonzero-winding holes / rect / oval) -> paint via Visio COM
with band-winding hole emulation and centre-containment z-order.
"""
import base64
import zlib
import re
import json
import math
import os
import urllib.parse
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

try:
    import win32com.client  # Windows + Visio only; parsing works without it
except ImportError:  # pragma: no cover - CI on non-Windows runners
    win32com = None

PPI = 96.0


# ---------------------------------------------------------------- helpers

def parse_style(style):
    d = {}
    if not style:
        return d
    for part in style.split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            k, v = part.split("=", 1)
            d[k.strip()] = v.strip()
        else:
            d[part] = True
    return d


def decode_stencil(data):
    """b64 + raw-deflate of URL-encoded XML; unquote until it parses."""
    data = data.strip()
    if "%" in data:
        xml = urllib.parse.unquote(data)
    else:
        xml = zlib.decompress(
            base64.b64decode(data + "=" * (-len(data) % 4)), -15).decode("utf-8")
    while not xml.lstrip().startswith("<"):
        new = urllib.parse.unquote(xml)
        if new == xml:
            break
        xml = new
    return ET.fromstring(xml)


class LabelHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.text = []
        self.size = None
        self.color = None
        self.bold = False
        self.align = None
        self.font = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        st = a.get("style", "")
        m = re.search(r"font-size:\s*([\d.]+)px", st)
        if m:
            v = float(m.group(1))
            if v > 2 and (self.size is None or v > self.size):
                self.size = v
        m = re.search(r"color:\s*(#[0-9a-fA-F]{3,6})", st)
        if m:
            self.color = m.group(1)
        m = re.search(r"text-align:\s*(\w+)", st)
        if m:
            self.align = m.group(1)
        m = re.search(r"font-family:\s*([^;\"']+)", st)
        if m:
            fam = m.group(1).strip()
            if fam and fam.lower() not in ("unknown",):
                self.font = fam
        if tag in ("b", "strong"):
            self.bold = True
        if tag == "br":
            self.text.append("\n")

    def handle_data(self, data):
        self.text.append(data)


def parse_label(html_label, style):
    info = {"text": "", "size": None, "color": None, "bold": False,
            "align": None, "font": None}
    if not html_label:
        return info
    p = LabelHTML()
    p.feed(html_label)
    info["text"] = "".join(p.text).strip()
    info["size"] = p.size
    info["color"] = p.color
    info["bold"] = p.bold
    info["align"] = p.align or style.get("align")
    info["font"] = p.font or style.get("fontFamily")
    return info


def hexrgb(v):
    v = v.lstrip("#")
    if len(v) == 3:
        v = "".join(c * 2 for c in v)
    return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)


# ---------------------------------------------------------------- geometry

def load_cells(path):
    tree = ET.parse(path)
    model = tree.getroot().find("diagram").find("mxGraphModel")
    nodes = model.find("root")
    cells = []
    for el in nodes:
        if el.tag == "UserObject":
            mx = el.find("mxCell")
        elif el.tag == "mxCell":
            mx = el
        else:
            continue
        if mx is None:
            continue
        cells.append({
            "id": el.get("id"),
            # drawio stores text in `value` (mxCell) or `label` (UserObject)
            "label": el.get("label", "") or mx.get("value", "") or "",
            "style": mx.get("style", "") or "", "vertex": mx.get("vertex"),
            "edge": mx.get("edge"), "source": mx.get("source"),
            "target": mx.get("target"), "parent": mx.get("parent", "1"),
            "geo": mx.find("mxGeometry"),
        })
    return cells


def build_offsets(cells):
    rects = {}
    for c in cells:
        if c["geo"] is not None:
            g = c["geo"]
            rects[c["id"]] = (float(g.get("x", 0) or 0), float(g.get("y", 0) or 0))
    offsets = {"1": (0.0, 0.0), "0": (0.0, 0.0)}
    changed = True
    while changed:
        changed = False
        for c in cells:
            pid = c["parent"]
            if pid in offsets and c["id"] not in offsets and c["geo"] is not None \
                    and c["geo"].get("relative") != "1":
                x, y = rects[c["id"]]
                ox, oy = offsets[pid]
                offsets[c["id"]] = (x + ox, y + oy)
                changed = True
    return offsets


def rect_of(cell, offsets):
    g = cell["geo"]
    if g is None:
        return None
    x = float(g.get("x", 0) or 0)
    y = float(g.get("y", 0) or 0)
    w = float(g.get("width", 0) or 0)
    h = float(g.get("height", 0) or 0)
    parent = cell["parent"]
    if parent in offsets and g.get("relative") != "1":
        ox, oy = offsets[parent]
        x += ox
        y += oy
    return x, y, w, h


# ------------------------------------------------- stencil rendering model

def arc_points(p0, p1, rx, ry, rot_deg, large_arc, sweep, n=24):
    x1, y1 = p0
    x2, y2 = p1
    if rx == 0 or ry == 0:
        return [p1]
    rx, ry = abs(rx), abs(ry)
    phi = math.radians(rot_deg)
    cosphi, sinphi = math.cos(phi), math.sin(phi)
    dx2, dy2 = (x1 - x2) / 2.0, (y1 - y2) / 2.0
    x1p = cosphi * dx2 + sinphi * dy2
    y1p = -sinphi * dx2 + cosphi * dy2
    lam = (x1p / rx) ** 2 + (y1p / ry) ** 2
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s
        ry *= s
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    co = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large_arc == sweep:
        co = -co
    cxp = co * rx * y1p / ry
    cyp = -co * ry * x1p / rx
    cx = cosphi * cxp - sinphi * cyp + (x1 + x2) / 2.0
    cy = sinphi * cxp + cosphi * cyp + (y1 + y2) / 2.0

    def ang(ux, uy, vx, vy):
        d = math.hypot(ux, uy) * math.hypot(vx, vy)
        c = max(-1.0, min(1.0, (ux * vx + uy * vy) / d)) if d else 1.0
        a = math.acos(c)
        if ux * vy - uy * vx < 0:
            a = -a
        return a

    ux = (x1p - cxp) / rx
    uy = (y1p - cyp) / ry
    vx = (-x1p - cxp) / rx
    vy = (-y1p - cyp) / ry
    da = ang(ux, uy, vx, vy)
    if not sweep and da > 0:
        da -= 2 * math.pi
    elif sweep and da < 0:
        da += 2 * math.pi
    pts = []
    th0 = math.atan2(uy, ux)
    for i in range(1, n + 1):
        th = th0 + da * (i / n)
        px = cx + rx * math.cos(th) * cosphi - ry * math.sin(th) * sinphi
        py = cy + rx * math.cos(th) * sinphi + ry * math.sin(th) * cosphi
        pts.append((px, py))
    return pts


def stencil_primitives(shape_el, style_fill, style_stroke):
    """Return primitives in stencil units. A fill/stroke paint directive closes
    every open subpath accumulated since the last directive; each <move> starts
    a new subpath. Filled multi-subpath shapes carry all loops in one compound
    prim so the builder can reproduce nonzero-winding holes."""
    W = float(shape_el.get("w", 100) or 100)
    H = float(shape_el.get("h", 100) or 100)
    prims = []

    def walk_section(sec):
        subpaths = []
        cur, start, closed = [], None, False
        segcur = []
        fill_on, stroke_on = False, False
        fc, sc = style_fill, style_stroke
        alpha = 1.0

        def flush_subpath():
            nonlocal cur, start, closed, segcur
            if len(cur) >= 2:
                subpaths.append((list(cur), closed, list(segcur)))
            cur, start, closed = [], None, False
            segcur = []

        def emit(fill_on, stroke_on, fc, sc, alpha):
            flush_subpath()
            if not subpaths:
                return
            filled = bool(fill_on)
            stroked = bool(stroke_on)
            if not filled and not stroked:
                subpaths.clear()
                return
            if filled and len(subpaths) == 1:
                prims.append({"kind": "poly", "subpaths": [list(subpaths[0][0])],
                              "closed": subpaths[0][1],
                              "segs": [subpaths[0][2]],
                              "fill": fc if filled else None,
                              "stroke": sc if stroked else None, "alpha": alpha})
            else:
                prims.append({"kind": "compound",
                              "subpaths": [list(p) for p, _, _ in subpaths],
                              "closeds": [c for _, c, _ in subpaths],
                              "segs": [s for _, _, s in subpaths],
                              "fill": fc if filled else None,
                              "stroke": sc if stroked else None, "alpha": alpha})
            subpaths.clear()

        def walk(node):
            nonlocal cur, start, closed, segcur, fill_on, stroke_on, fc, sc, alpha
            for ch in node:
                tag = ch.tag
                if tag == "path":
                    walk(ch)
                elif tag == "move":
                    if cur:
                        flush_subpath()
                    x = float(ch.get("x", 0) or 0)
                    y = float(ch.get("y", 0) or 0)
                    cur = [(x, y)]
                    segcur = [("M", x, y)]
                    start = (x, y)
                    closed = False
                elif tag == "line":
                    x = float(ch.get("x", 0) or 0)
                    y = float(ch.get("y", 0) or 0)
                    cur.append((x, y))
                    segcur.append(("L", x, y))
                elif tag == "curve":
                    x1 = float(ch.get("x1", 0) or 0)
                    y1 = float(ch.get("y1", 0) or 0)
                    x2 = float(ch.get("x2", 0) or 0)
                    y2 = float(ch.get("y2", 0) or 0)
                    x3 = float(ch.get("x3", 0) or 0)
                    y3 = float(ch.get("y3", 0) or 0)
                    if cur:
                        p0 = cur[-1]
                        segcur.append(("C", x1, y1, x2, y2, x3, y3))
                        for i in range(1, 49):
                            t = i / 48.0
                            mt = 1 - t
                            bx = mt**3 * p0[0] + 3 * mt * mt * t * x1 + \
                                3 * mt * t * t * x2 + t**3 * x3
                            by = mt**3 * p0[1] + 3 * mt * mt * t * y1 + \
                                3 * mt * t * t * y2 + t**3 * y3
                            cur.append((bx, by))
                elif tag == "quad":
                    x1 = float(ch.get("x1", 0) or 0)
                    y1 = float(ch.get("y1", 0) or 0)
                    x3 = float(ch.get("x3", 0) or 0)
                    y3 = float(ch.get("y3", 0) or 0)
                    if cur:
                        p0 = cur[-1]
                        # promote quadratic to cubic
                        cx = 2.0/3.0 * x1 + 1.0/3.0 * p0[0]
                        cy = 2.0/3.0 * y1 + 1.0/3.0 * p0[1]
                        cx2 = 2.0/3.0 * x1 + 1.0/3.0 * x3
                        cy2 = 2.0/3.0 * y1 + 1.0/3.0 * y3
                        segcur.append(("C", cx, cy, cx2, cy2, x3, y3))
                        for i in range(1, 49):
                            t = i / 48.0
                            mt = 1 - t
                            bx = mt * mt * p0[0] + 2 * mt * t * x1 + t * t * x3
                            by = mt * mt * p0[1] + 2 * mt * t * y1 + t * t * y3
                            cur.append((bx, by))
                elif tag == "arc":
                    if cur:
                        p0 = cur[-1]
                        cur.extend(arc_points(
                            p0,
                            (float(ch.get("x", 0) or 0),
                             float(ch.get("y", 0) or 0)),
                            float(ch.get("rx", 0) or 0),
                            float(ch.get("ry", 0) or 0),
                            float(ch.get("x-axis-rotation", 0) or 0),
                            ch.get("large-arc-flag", "0") == "1",
                            ch.get("sweep-flag", "0") == "1"))
                elif tag == "close":
                    if start:
                        cur.append(start)
                        segcur.append(("L", start[0], start[1]))
                    closed = True
                elif tag in ("rect", "roundrect"):
                    emit(fill_on, stroke_on, fc, sc, alpha)
                    prims.append({"kind": "rect",
                                  "x": float(ch.get("x", 0) or 0),
                                  "y": float(ch.get("y", 0) or 0),
                                  "w": float(ch.get("w", 0) or 0),
                                  "h": float(ch.get("h", 0) or 0),
                                  "fill": fc, "stroke": sc, "alpha": alpha})
                elif tag == "ellipse":
                    emit(fill_on, stroke_on, fc, sc, alpha)
                    prims.append({"kind": "oval",
                                  "x": float(ch.get("x", 0) or 0),
                                  "y": float(ch.get("y", 0) or 0),
                                  "w": float(ch.get("w", 0) or 0),
                                  "h": float(ch.get("h", 0) or 0),
                                  "fill": fc, "stroke": sc, "alpha": alpha})
                elif tag == "fillstroke":
                    fill_on, stroke_on = True, True
                    emit(fill_on, stroke_on, fc, sc, alpha)
                elif tag == "fill":
                    fill_on = True
                    emit(fill_on, stroke_on, fc, sc, alpha)
                elif tag == "stroke":
                    stroke_on = True
                    emit(fill_on, stroke_on, fc, sc, alpha)
                elif tag == "fillcolor":
                    fc = ch.get("color")
                elif tag == "strokecolor":
                    sc = ch.get("color")
                elif tag == "alpha":
                    alpha = float(ch.get("alpha", 100) or 100) / 100.0
                elif tag == "stencil":
                    emit(fill_on, stroke_on, fc, sc, alpha)
                    sx = float(ch.get("x", 0) or 0)
                    sy = float(ch.get("y", 0) or 0)
                    subsc = float(ch.get("scale", 1) or 1)
                    data = (ch.text or "").strip()
                    if data:
                        try:
                            sub = decode_stencil(data)
                            subprims, _, _ = stencil_primitives(
                                sub, style_fill, style_stroke)
                            for sp in subprims:
                                prims.append(scale_shift_prim(sp, sx, sy, subsc))
                        except Exception:
                            pass
                else:
                    walk(ch)

        walk(sec)
        emit(fill_on, stroke_on, fc, sc, alpha)

    bg = shape_el.find("background")
    if bg is not None:
        walk_section(bg)
    fg = shape_el.find("foreground")
    if fg is not None:
        walk_section(fg)
    return prims, (W, H)


def transform_segs(segs, tx):
    out = []
    for s in segs:
        if s[0] in ("M", "L"):
            x, y = tx(s[1], s[2])
            out.append((s[0], x, y))
        elif s[0] == "C":
            a, b = tx(s[1], s[2]); c, d = tx(s[3], s[4]); e, f = tx(s[5], s[6])
            out.append(("C", a, b, c, d, e, f))
        else:
            out.append(s)
    return out


def scale_shift_prim(p, sx, sy, subsc):
    def tx(px, py):
        return (sx + px * subsc, sy + py * subsc)
    if p["kind"] == "poly":
        sp0, cl0, sg0 = p["subpaths"][0], p["closed"], p.get("segs")
        nsegs = {"kind": "poly",
                 "subpaths": [[tx(a, b) for a, b in sp0]],
                 "closed": cl0, "fill": p["fill"], "stroke": p["stroke"],
                 "alpha": p.get("alpha", 1)}
        if sg0:
            nsegs["segs"] = [transform_segs(sg0, tx)]
        return nsegs
    if p["kind"] == "compound":
        q = {"kind": "compound",
             "subpaths": [[tx(a, b) for a, b in sp] for sp in p["subpaths"]],
             "closeds": p["closeds"], "fill": p["fill"], "stroke": p["stroke"],
             "alpha": p.get("alpha", 1)}
        if p.get("segs"):
            q["segs"] = [transform_segs(sg, tx) for sg in p["segs"]]
        return q
    q = dict(p)
    qx, qy = tx(p["x"], p["y"])
    q["x"], q["y"] = qx, qy
    q["w"] = p["w"] * subsc
    q["h"] = p["h"] * subsc
    return q


# ---------------------------------------------------------------- visio

class VisioBuilder:
    def __init__(self, width_px, height_px):
        self.w_in = width_px / PPI
        self.h_in = height_px / PPI
        self.app = win32com.client.DispatchEx("Visio.InvisibleApp")
        self.app.AlertResponse = 2
        self.doc = self.app.Documents.Add("")
        self.win = self.app.ActiveWindow
        self.page = self.doc.Pages(1)
        self.page.PageSheet.Cells("PageWidth").FormulaU = f"{self.w_in:.4f} in"
        self.page.PageSheet.Cells("PageHeight").FormulaU = f"{self.h_in:.4f} in"
        self.text_shapes = []
        self.rects = []          # (x, y, w, h, shape) for connector gluing   # ((x, y, w, h) canvas px, shape)

    def X(self, x):
        return x / PPI

    def Y(self, y):
        return self.h_in - y / PPI

    @staticmethod
    def rgb(v):
        r, g, b = hexrgb(v)
        return f"RGB({r},{g},{b})"

    def _paint(self, shp, fill, stroke, sw, dashed, alpha=1.0):
        if fill:
            shp.Cells("FillForegnd").FormulaU = self.rgb(fill)
            if alpha < 1:
                shp.Cells("FillTrans").FormulaU = f"{(1 - alpha) * 100:.0f} PERCENT"
        else:
            shp.Cells("FillPattern").FormulaU = "0"
        if stroke:
            shp.Cells("LineColor").FormulaU = self.rgb(stroke)
            shp.Cells("LineWeight").FormulaU = f"{max(sw * 0.75, 0.5):.2f} pt"
        else:
            shp.Cells("LinePattern").FormulaU = "0"
        if dashed:
            shp.Cells("LinePattern").FormulaU = "2"

    def _poly_shape(self, pts, closed):
        arr = []
        last = None
        for px, py in pts:
            cur = (self.X(px), self.Y(py))
            if last is not None and abs(cur[0]-last[0]) < 1e-6 and abs(cur[1]-last[1]) < 1e-6:
                continue
            arr.append(cur[0]); arr.append(cur[1]); last = cur
        try:
            return self.page.DrawPolyline(arr, 0)
        except Exception:
            if len(arr) <= 8:
                raise
            pts2 = list(zip(arr[0::2], arr[1::2]))[::2]
            arr = [c for pt in pts2 for c in pt]
            try:
                return self.page.DrawPolyline(arr, 0)
            except Exception:
                pts2 = pts2[::2]
                arr = [c for pt in pts2 for c in pt]
                return self.page.DrawPolyline(arr, 0)

    def poly(self, subpaths, closed, fill, stroke, sw, dashed, alpha=1.0):
        pts = subpaths[0]
        if closed and pts and pts[0] != pts[-1]:
            pts = pts + [pts[0]]
        shp = self._poly_shape(pts, closed)
        self._paint(shp, fill, stroke, sw, dashed, alpha)
        return shp

    def compound(self, subpaths, closeds, fill, stroke, sw, dashed, alpha=1.0,
                 hole_fill="#FFFFFF"):
        """Faithful nonzero-winding reproduction. For each loop sample its
        BAND winding (longest-edge midpoint nudged inward); winding != 0 ->
        fill with the icon colour, == 0 -> hole -> hole_fill. Paint by
        containment depth ascending. Never use Selection.Subtract: it unions."""
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
            return (sum(p[0] for p in pts) / len(pts),
                    sum(p[1] for p in pts) / len(pts))

        signs = [1 if signed_area(sp) > 0 else -1 for sp in subpaths]
        cents = [centroid(sp) for sp in subpaths]

        def band_state(i):
            pts = subpaths[i]
            bl, bm, bn = -1.0, None, None
            for k in range(len(pts) - 1):
                x0, y0 = pts[k]
                x1, y1 = pts[k + 1]
                L = math.hypot(x1 - x0, y1 - y0)
                if L > bl:
                    bl = L
                    mx, my = (x0 + x1) / 2.0, (y0 + y1) / 2.0
                    nx, ny = -(y1 - y0) / L, (x1 - x0) / L
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
            colour = fill if wind[i] != 0 else hole_fill
            plan.append((depth_of(i), i, colour))
        plan.sort(key=lambda t: t[0])

        if not fill:
            for pts, cl in zip(subpaths, closeds):
                p2 = pts
                if cl and pts and pts[0] != pts[-1]:
                    p2 = pts + [pts[0]]
                shp = self._poly_shape(p2, cl)
                self._paint(shp, None, stroke, sw, dashed, alpha)
            return

        body_i = max(range(n), key=lambda k: abs(signed_area(subpaths[k])))
        body = subpaths[body_i]
        if closeds[body_i] and body and body[0] != body[-1]:
            body = body + [body[0]]
        shp = self._poly_shape(body, True)
        self._paint(shp, fill, stroke, sw, dashed, alpha)

        for _, i, colour in plan:
            if i == body_i:
                continue
            h = subpaths[i]
            if closeds[i] and h and h[0] != h[-1]:
                h = h + [h[0]]
            hs = self._poly_shape(h, True)
            hs.Cells("FillForegnd").FormulaU = self.rgb(colour)
            if alpha < 1:
                hs.Cells("FillTrans").FormulaU = \
                    f"{(1 - alpha) * 100:.0f} PERCENT"
            if stroke:
                hs.Cells("LineColor").FormulaU = self.rgb(stroke)
                hs.Cells("LineWeight").FormulaU = \
                    f"{max(sw * 0.75, 0.5):.2f} pt"
            else:
                hs.Cells("LinePattern").FormulaU = "0"
            if dashed:
                hs.Cells("LinePattern").FormulaU = "2"
        return shp

    def rect(self, x, y, w, h, fill, stroke, sw, dashed, alpha=1.0):
        shp = self.page.DrawRectangle(self.X(x), self.Y(y + h),
                                      self.X(x + w), self.Y(y))
        self._paint(shp, fill, stroke, sw, dashed, alpha)
        self.rects.append((x, y, w, h, shp))
        return shp

    def oval(self, x, y, w, h, fill, stroke, sw, dashed, alpha=1.0):
        shp = self.page.DrawOval(self.X(x), self.Y(y + h),
                                 self.X(x + w), self.Y(y))
        self._paint(shp, fill, stroke, sw, dashed, alpha)
        return shp

    def text(self, x, y, w, h, info, v_align):
        shp = self.page.DrawRectangle(self.X(x), self.Y(y + h),
                                      self.X(x + w), self.Y(y))
        shp.Cells("FillPattern").FormulaU = "0"
        shp.Cells("LinePattern").FormulaU = "0"
        shp.Cells("LeftMargin").FormulaU = "0 pt"
        shp.Cells("RightMargin").FormulaU = "0 pt"
        shp.Cells("TopMargin").FormulaU = "0 pt"
        shp.Cells("BottomMargin").FormulaU = "0 pt"
        shp.Text = info["text"]
        self.text_shapes.append(((x, y, w, h), shp))
        if info["size"]:
            shp.Cells("Char.Size").FormulaU = f"{info['size'] * 0.75:.2f} pt"
        shp.Cells("Char.Color").FormulaU = self.rgb(info["color"] or "#000000")
        if info.get("font"):
            # map draw.io font names to installed Visio fonts; Char.Font
            # takes the font's numeric ID from the document font table
            fam = info["font"]
            if fam.lower() in ("helvetica", "arial", "oracle sans"):
                fam = "Arial"
            try:
                shp.Cells("Char.Font").FormulaU = str(
                    self.doc.Fonts(fam).ID)
            except Exception:
                pass
        if info["bold"]:
            shp.Cells("Char.Style").FormulaU = "1"
        align = info["align"]
        shp.Cells("Para.HorzAlign").FormulaU = {
            "center": "1", "right": "2"}.get(align, "0")
        shp.Cells("VerticalAlign").FormulaU = {
            "top": "0", "middle": "1", "bottom": "2"}.get(v_align or "middle", "1")
        return shp

    def save(self, path):
        self.doc.SaveAs(path)

    def quit(self):
        try:
            # mark clean so no AutoRecovery data is written and no
            # "Recovered Documents" pane appears on the user's next
            # real Visio launch
            self.doc.Saved = True
            self.doc.Close()
        except Exception:
            pass
        self.app.Quit()


# ---------------------------------------------------------------- holes
def compound_has_holes(prims):
    """True if any filled compound primitive has a loop whose band winding
    is 0 (i.e. the compound paints an opaque white hole)."""
    for p in prims:
        if p["kind"] != "compound" or not p.get("fill"):
            continue
        sps = p["subpaths"]
        n = len(sps)
        if n < 2:
            continue

        def _sa(pts):
            s = 0.0
            for i in range(len(pts) - 1):
                s += pts[i][0] * pts[i + 1][1] - pts[i + 1][0] * pts[i][1]
            return s / 2.0

        def _pip(pt, poly):
            x, y = pt
            ins = False
            j = len(poly) - 1
            for i in range(len(poly)):
                xi, yi = poly[i]
                xj, yj = poly[j]
                if ((yi > y) != (yj > y)) and \
                        (x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi):
                    ins = not ins
                j = i
            return ins

        signs = [1 if _sa(sp) > 0 else -1 for sp in sps]
        for i in range(n):
            pts = sps[i]
            bl, bm, bn = -1.0, None, None
            for k in range(len(pts) - 1):
                x0, y0 = pts[k]
                x1, y1 = pts[k + 1]
                L = math.hypot(x1 - x0, y1 - y0)
                if L > bl:
                    bl = L
                    mx, my = (x0 + x1) / 2.0, (y0 + y1) / 2.0
                    nx, ny = -(y1 - y0) / L, (x1 - x0) / L
                    if signs[i] < 0:
                        nx, ny = -nx, -ny
                    bm, bn = (mx, my), (nx, ny)
            eps = max(1e-3, bl * 1e-3)
            ip = (bm[0] + bn[0] * eps, bm[1] + bn[1] * eps)
            if not _pip(ip, pts):
                ip = (bm[0] - bn[0] * eps, bm[1] - bn[1] * eps)
            w = signs[i]
            for j in range(n):
                if j != i and _pip(ip, sps[j]):
                    w += signs[j]
            if w == 0:
                return True
    return False


def cell_hole_ids(cells):
    """IDs of vertex cells whose stencil paints at least one hole loop."""
    ids = set()
    for c in cells:
        if not c["vertex"]:
            continue
        mm = re.search(r"stencil\((.+?)\)", c["style"], re.S)
        if not mm:
            continue
        try:
            sel = decode_stencil(mm.group(1))
        except Exception:
            continue
        style = parse_style(c["style"])
        fill = style.get("fillColor")
        if fill == "none":
            fill = None
        try:
            prims, _ = stencil_primitives(sel, fill, style.get("strokeColor"))
        except Exception:
            continue
        if compound_has_holes(prims):
            ids.add(c["id"])
    return ids


# ---------------------------------------------------------------- main

def main(src_path, out_path, stencil_vssx=None,
         real_connectors=True):
    """Convert a draw.io file to Visio. When `stencil_vssx` is given, the
    pipeline is:
      1. extract every icon into a named, de-duplicated .vssx stencil set
         (same engine as drawio2stencils.py)
      2. dock that stencil set and build the final diagram by DROPping the
         masters - each icon in the output is a single editable master
         instance instead of many loose polylines.
    Without it, the classic inline-geometry conversion runs (default)."""
    if stencil_vssx:
        from .stencils import (collect, clusters_of, dedupe,
                               name_cluster, draw_cluster_into)

    cells = load_cells(src_path)
    offsets = build_offsets(cells)

    max_x = max_y = 0.0
    for c in cells:
        if c["vertex"]:
            r = rect_of(c, offsets)
            if r:
                max_x = max(max_x, r[0] + r[2])
                max_y = max(max_y, r[1] + r[3])
    print(f"canvas: {max_x:.0f} x {max_y:.0f} px")

    b = VisioBuilder(max_x + 10, max_y + 10)

    # ---- STEP 1 (optional): build the stencil set and cluster the icons --
    member_of = {}       # cell id -> (master name, cluster bbox)
    if stencil_vssx:
        from .stencils import (collect, clusters_of, dedupe,
                               name_cluster, draw_cluster_into)
        icon_cells, labels2, conn_parts = collect(
            cells, offsets,
            real_connectors=real_connectors)
        conn_by_id = {cp["id"]: cp for cp in conn_parts}
        clusters = clusters_of(icon_cells)
        uniq = dedupe(clusters)
        print(f"step 1: unique stencil masters = {len(uniq)}")

        used_names = {}
        master_of_cluster = []  # (cluster, master name)
        doc_order = {cell["id"]: i for i, cell in enumerate(cells)}
        for idx, cl in enumerate(uniq):
            # draw in the ORIGINAL drawio z-order (document order):
            # knockout icons (white cutouts under dark glyphs) break
            # when sorted by area
            cl_sorted = sorted(
                cl, key=lambda mm: -(mm["rect"][2] * mm["rect"][3]))
            name = name_cluster(cl, labels2) or f"Icon {idx + 1}"
            base, k = name, 2
            while name in used_names:
                name = f"{base} {k}"
                k += 1
            used_names[name] = True
            master_of_cluster.append((cl_sorted, name))

        stencil_doc = b.app.Documents.AddEx("", 4, 0, 0)
        stencil_doc.Title = os.path.splitext(
            os.path.basename(stencil_vssx))[0]
        for cl_sorted, name in master_of_cluster:
            master = stencil_doc.Masters.Add()
            master.Name = name
            editor = master.Open()
            try:
                ci = id(cl_sorted)
                hf = "#FFFFFF"
                for (clx, nm), hfv in ((x, y) for x, y in []):
                    pass
                draw_cluster_into(editor, cl_sorted, cells, offsets, b.app)
            finally:
                editor.Close()
        stencil_doc.SaveAs(os.path.abspath(stencil_vssx))
        print(f"step 1 done: {stencil_doc.Masters.Count} masters "
              f"-> {stencil_vssx}")

        # map every member cell id -> its cluster (by cluster index)
        cluster_of_cell = {}
        for ci, (cl_sorted, name) in enumerate(master_of_cluster):
            for mm in cl_sorted:
                cluster_of_cell[mm["id"]] = ci
        # cluster bbox in canvas coords
        cluster_bbox = {}
        for ci, (cl_sorted, name) in enumerate(master_of_cluster):
            bx0 = min(mm["rect"][0] for mm in cl_sorted)
            by0 = min(mm["rect"][1] for mm in cl_sorted)
            bx1 = max(mm["rect"][0] + mm["rect"][2] for mm in cl_sorted)
            by1 = max(mm["rect"][1] + mm["rect"][3] for mm in cl_sorted)
            cluster_bbox[ci] = (bx0, by0, bx1 - bx0, by1 - by0)
        member_of = cluster_of_cell

    hole_ids = cell_hole_ids(cells)

    def z_key(c):
        # Containment depth, centre-based: a cell whose bbox contains this
        # cell's CENTRE (and is bigger) is a container and must draw first.
        # Full-bbox containment fails for label cells that overhang their
        # container's edge (e.g. the OCI container titles, width 338 > 244).
        r = rect_of(c, offsets)
        if r is None:
            return (0, 0)
        x, y, w, h = r
        cx0, cy0 = x + w / 2.0, y + h / 2.0
        depth = 0
        for o in cells:
            if o is c or not o["vertex"]:
                continue
            ro = rect_of(o, offsets)
            if ro is None:
                continue
            ox, oy, ow, oh = ro
            if ow <= 0 or oh <= 0 or ow * oh <= w * h:
                continue
            if ox <= cx0 <= ox + ow and oy <= cy0 <= oy + oh:
                depth += 1
        return (depth, 0)

    ordered = sorted(cells, key=z_key)

    n_v = n_t = n_e = n_fail = 0
    try:
        if stencil_vssx:
            st = b.app.Documents.OpenEx(os.path.abspath(stencil_vssx), 4)
        dropped_clusters = set()
        icon_page_shapes = []   # ((x, y, w, h) canvas px, master instance)
        label_page_shapes = []  # ((x, y, w, h) canvas px, text shape)

        for c in ordered:
            # ---- STEP 2: when the first member of an icon cluster comes up
            # in z-order, drop its master here (preserves stacking) -------
            ci = member_of.get(c["id"]) if stencil_vssx else None
            if ci is not None and ci not in dropped_clusters:
                dropped_clusters.add(ci)
                bx, by, bw2, bh2 = cluster_bbox[ci]
                m = st.Masters.ItemU(master_of_cluster[ci][1])
                shp = b.page.Drop(m, b.X(bx + bw2 / 2), b.Y(by + bh2 / 2))
                shp.Cells("PinX").FormulaU = f"{b.X(bx + bw2 / 2):.4f} in"
                shp.Cells("PinY").FormulaU = f"{b.Y(by + bh2 / 2):.4f} in"
                shp.Cells("Width").FormulaU = f"{bw2 / PPI:.4f} in"
                shp.Cells("Height").FormulaU = f"{bh2 / PPI:.4f} in"
                icon_page_shapes.append(
                    ((bx, by, bw2, bh2), shp))
                n_v += 1
            if stencil_vssx and c["id"] in conn_by_id:
                # connector fragment: replaced by real Visio connectors
                # created after the main loop
                continue
            if ci is not None:
                # member cell graphics are replaced by its master instance,
                # but its TEXT label must still be rendered (icon labels like
                # e.g. "App / on-premises / agent" labels live on cluster cells)
                lbl = c.get("label")
                if lbl:
                    mstyle = parse_style(c["style"])
                    minfo = parse_label(lbl, mstyle)
                    if minfo["text"]:
                        mr = rect_of(c, offsets)
                        if mr and mr[2] > 0 and mr[3] > 0:
                            tshp = b.text(mr[0], mr[1], mr[2], mr[3], minfo,
                                          mstyle.get("verticalAlign"))
                            if tshp is not None:
                                label_page_shapes.append((mr, tshp))
                            n_t += 1
                continue  # member cell - replaced by its master instance
            style = parse_style(c["style"])
            if c["vertex"]:
                r = rect_of(c, offsets)
                if not r or r[2] <= 0 or r[3] <= 0:
                    continue
                x, y, w, h = r
                fill = style.get("fillColor")
                if fill == "none":
                    fill = None
                stroke = style.get("strokeColor")
                if stroke == "none":
                    stroke = None
                sw = float(style.get("strokeWidth", 1) or 1)
                dashed = bool(style.get("dashed"))
                label = c["label"]

                stencil_data = None
                m = re.search(r"stencil\((.+?)\)", style.get("shape", ""), re.S)
                if m:
                    stencil_data = m.group(1)

                if stencil_data:
                    try:
                        sel = decode_stencil(stencil_data)
                    except Exception as ex:
                        sel = None
                        n_fail += 1
                        if n_fail <= 3:
                            print("stencil decode fail id", c["id"], ex)
                    if sel is not None:
                        prims, (sw0, sh0) = stencil_primitives(sel, fill, stroke)
                        sxx = w / sw0
                        syy = h / sh0
                        for p in prims:
                            alpha = p.get("alpha", 1.0)
                            if p["kind"] == "poly":
                                sps = [[(x + px * sxx, y + py * syy)
                                        for px, py in sp]
                                       for sp in p["subpaths"]]
                                b.poly(sps, p["closed"], p["fill"], p["stroke"],
                                       sw, dashed, alpha)
                            elif p["kind"] == "compound":
                                sps = [[(x + px * sxx, y + py * syy)
                                        for px, py in sp]
                                       for sp in p["subpaths"]]
                                b.compound(sps, p["closeds"], p["fill"],
                                           p["stroke"], sw, dashed, alpha)
                            elif p["kind"] == "rect":
                                b.rect(x + p["x"] * sxx, y + p["y"] * syy,
                                       p["w"] * sxx, p["h"] * syy,
                                       p["fill"], p["stroke"], sw, dashed, alpha)
                            elif p["kind"] == "oval":
                                b.oval(x + p["x"] * sxx, y + p["y"] * syy,
                                       p["w"] * sxx, p["h"] * syy,
                                       p["fill"], p["stroke"], sw, dashed, alpha)
                        n_v += 1
                        if label:
                            info = parse_label(label, style)
                            if info["text"]:
                                b.text(x, y, w, h, info,
                                       style.get("verticalAlign"))
                                n_t += 1
                        continue

                is_text = style.get("shape") == "text" or "text" in style
                if is_text:
                    info = parse_label(label, style)
                    if info["text"]:
                        b.text(x, y, w, h, info, style.get("verticalAlign"))
                        n_t += 1
                    continue

                if style.get("ellipse"):
                    b.oval(x, y, w, h, fill, stroke, sw, dashed)
                    n_v += 1
                else:
                    shp = b.rect(x, y, w, h, fill, stroke, sw, dashed)
                    if style.get("rounded") == "1" and shp is not None:
                        try:
                            shp.Cells("Rounding").FormulaU = \
                                f"{min(w, h) * 0.06 / PPI:.4f} in"
                        except Exception:
                            pass
                    n_v += 1
                if label:
                    info = parse_label(label, style)
                    if info["text"]:
                        b.text(x, y, w, h, info, style.get("verticalAlign"))
                        n_t += 1

            elif c["edge"]:
                geo = c["geo"]
                pts = []
                by_id = cells_by_id
                spc = by_id.get(c["source"])
                tpc = by_id.get(c["target"])
                sp = rect_of(spc, offsets) if spc else None
                tp = rect_of(tpc, offsets) if tpc else None
                if sp:
                    ex = float(style.get("exitX", 0.5) or 0.5)
                    ey = float(style.get("exitY", 0.5) or 0.5)
                    pts.append((sp[0] + ex * sp[2], sp[1] + ey * sp[3]))
                elif geo is not None:
                    gsp = geo.find("sourcePoint")
                    if gsp is not None:
                        pts.append((float(gsp.get("x", 0) or 0),
                                    float(gsp.get("y", 0) or 0)))
                if geo is not None:
                    arr = geo.find("Array")
                    if arr is not None:
                        for mp in arr.findall("mxPoint"):
                            pts.append((float(mp.get("x", 0) or 0),
                                        float(mp.get("y", 0) or 0)))
                if tp:
                    enx = float(style.get("entryX", 0.5) or 0.5)
                    eny = float(style.get("entryY", 0.5) or 0.5)
                    pts.append((tp[0] + enx * tp[2], tp[1] + eny * tp[3]))
                elif geo is not None:
                    gtp = geo.find("targetPoint")
                    if gtp is not None:
                        pts.append((float(gtp.get("x", 0) or 0),
                                    float(gtp.get("y", 0) or 0)))
                if len(pts) >= 2:
                    color = style.get("strokeColor", "#000000")
                    sw = float(style.get("strokeWidth", 1) or 1)
                    dashed = bool(style.get("dashed"))
                    ea = style.get("endArrow", "none") != "none"
                    ba = style.get("startArrow", "none") != "none"
                    varr = []
                    for px, py in pts:
                        varr.append(b.X(px))
                        varr.append(b.Y(py))
                    shp = b.page.DrawPolyline(varr, 0)
                    shp.Cells("FillPattern").FormulaU = "0"
                    if color:
                        shp.Cells("LineColor").FormulaU = b.rgb(color)
                    shp.Cells("LineWeight").FormulaU = \
                        f"{max(sw * 0.75, 0.75):.2f} pt"
                    if dashed:
                        shp.Cells("LinePattern").FormulaU = "2"
                    if ea:
                        shp.Cells("EndArrow").FormulaU = "5"
                    if ba:
                        shp.Cells("BeginArrow").FormulaU = "5"
                    n_e += 1


        # ---- STEP 3 (--real-connectors): replace slim directional-arrow
        # stencil glyphs with real Visio dynamic connectors (editable,
        # routable, arrowhead included)
        if stencil_vssx:
            # close the docked stencil document so the file is unlocked
            try:
                for d in list(b.app.Documents):
                    if d.Type == 2:  # visDocTypeStencil
                        d.Close()
            except Exception:
                pass
            from .stencils import _inject_stencil_window
            try:
                _inject_stencil_window(os.path.abspath(stencil_vssx),
                                       os.path.basename(stencil_vssx))
            except (PermissionError, OSError) as e:
                print(f"note: stencil window injection skipped ({e})")
        if stencil_vssx and real_connectors and conn_parts:
            # path-for-path: draw each detected arrow as a polyline that
            # reproduces the drawio route exactly (control point per
            # corner). Polylines never cross panels unexpectedly, so
            # icons and labels stay clean.
            n_conn = 0
            for cp in conn_parts:
                x, y, w, h = cp["rect"]
                chain = cp.get("chain")
                try:
                    if chain and len(chain) >= 2:
                        pts = []
                        for cx, cy in chain:
                            pts.extend([b.X(x + cx / 100.0 * w),
                                        b.Y(y + cy / 100.0 * h)])
                        conn = b.page.DrawPolyline(pts, 0)
                    else:
                        if h >= w:
                            x0, y0 = b.X(x + w / 2), b.Y(y)
                            x1, y1 = b.X(x + w / 2), b.Y(y + h)
                        else:
                            x0, y0 = b.X(x), b.Y(y + h / 2)
                            x1, y1 = b.X(x + w), b.Y(y + h / 2)
                        conn = b.page.DrawLine(x0, y0, x1, y1)
                    conn.Cells("EndArrow").FormulaU = "5"
                    # use the arrow cell's own fill colour (drawio draws
                    # these arrows as filled stencil glyphs)
                    col = parse_style(
                        cells_by_id.get(cp["id"], {}).get("style", "")
                    ).get("fillColor") or "#312d2a"
                    conn.Cells("LineColor").FormulaU = b.rgb(col)
                    conn.Cells("LineWeight").FormulaU = "1.25 pt"
                    n_conn += 1
                except Exception as ex:
                    print(f"note: connector creation failed ({ex})")
            print(f"connectors={n_conn}")

        b.save(out_path)
    finally:
        b.quit()

    print(f"vertexes={n_v} texts={n_t} edges={n_e} stencil_fail={n_fail}")
    print("saved:", out_path)


cells_by_id = {}

if __name__ == "__main__":
    import sys as _sys
    if len(_sys.argv) < 3:
        raise SystemExit("usage: python -m drawio_to_visio.core "
                         "input.drawio output.vsdx [icons.vssx] "
                         "[--no-real-connectors]")
    _args = [a for a in _sys.argv[1:] if a != "--no-real-connectors"]
    _real = "--no-real-connectors" not in _sys.argv
    _src = os.path.abspath(_args[0])
    _out = os.path.abspath(_args[1])
    _vssx = os.path.abspath(_args[2]) if len(_args) > 2 else None
    _cells = load_cells(_src)
    cells_by_id = {c["id"]: c for c in _cells}
    main.__globals__["load_cells"] = lambda path: _cells
    main(_src, _out, _vssx, real_connectors=_real)
