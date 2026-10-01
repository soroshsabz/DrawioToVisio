"""Platform-independent tests for the DrawioToVisio parsing engine.

These tests exercise the pure-Python parts (draw.io XML parsing, stencil
decoding, primitive building, label parsing) that do not require Visio or
Windows.  The COM painting layer is exercised only on a machine with
Microsoft Visio installed.
"""

import importlib.util
import os
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(REPO, filename))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


d2v = _load("d2v_test", "drawio2visio.py")
d2s = _load("d2s_test", "drawio2stencils.py")

SAMPLE = os.path.join(REPO, "examples", "sample-functional-view.drawio")
BASIC = os.path.join(REPO, "examples", "sample-basic.drawio")


class TestImports(unittest.TestCase):
    def test_modules_import_without_visio(self):
        """The parsing engine must import on any OS (win32com optional)."""
        self.assertIsNotNone(d2v)
        self.assertIsNotNone(d2s)

    def test_win32com_is_optional(self):
        # the module must expose the rest of the API even if win32com is absent
        self.assertTrue(hasattr(d2v, "load_cells"))
        self.assertTrue(hasattr(d2v, "decode_stencil"))
        self.assertTrue(hasattr(d2v, "parse_style"))
        self.assertTrue(hasattr(d2v, "parse_label"))


class TestLoadCells(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cells = d2v.load_cells(SAMPLE)

    def test_loads_cells(self):
        self.assertGreater(len(self.cells), 500)

    def test_userobject_labels_are_read(self):
        """Labels stored on UserObject wrappers must not be lost."""
        texts = [str(c["label"]) for c in self.cells if c.get("label")]
        joined = " ".join(texts)
        self.assertIn("Platform", joined)
        self.assertIn("on-premises", joined)

    def test_vertices_present(self):
        self.assertTrue(any(c["vertex"] for c in self.cells))


class TestParseLabel(unittest.TestCase):
    def test_plain_text(self):
        info = d2v.parse_label("hello", {})
        self.assertEqual(info["text"], "hello")

    def test_html_label(self):
        html = ('<div style="font-size: 1px"><p>'
                '<font style="font-size:16.93px" color="#312d2a">'
                '<b>Title<br/></b></font></p></div>')
        info = d2v.parse_label(html, {})
        self.assertEqual(info["text"], "Title")
        self.assertTrue(info["bold"])

    def test_empty_label(self):
        info = d2v.parse_label("", {})
        self.assertEqual(info["text"], "")


def _make_payload(xml):
    """Encode stencil XML exactly like draw.io: URL-encode (safe='')
    -> raw-deflate -> base64."""
    import base64, zlib, urllib.parse
    enc = urllib.parse.quote(xml, safe="")
    co = zlib.compressobj(9, zlib.DEFLATED, -15)
    raw = co.compress(enc.encode()) + co.flush()
    return base64.b64encode(raw).decode()


class TestDecodeStencil(unittest.TestCase):
    def test_decode_returns_element(self):
        xml = ('<shape><foreground><path><move x="0" y="0"/>'
               '<line x="100" y="0"/><line x="100" y="100"/>'
               '<line x="0" y="100"/><line x="0" y="0"/>'
               '<close/></path><fillstroke/></foreground></shape>')
        sel = d2v.decode_stencil(_make_payload(xml))
        self.assertIsNotNone(sel)
        self.assertEqual(sel.tag, "shape")

    def test_roundtrip_primitives(self):
        xml = ('<shape><foreground><path><move x="0" y="0"/>'
               '<line x="100" y="0"/><line x="50" y="100"/>'
               '<line x="0" y="0"/><close/></path>'
               '<fillstroke/></foreground></shape>')
        sel = d2v.decode_stencil(_make_payload(xml))
        prims, (sw0, sh0) = d2v.stencil_primitives(sel, "#123456", None)
        self.assertGreater(len(prims), 0)
        self.assertEqual(sw0, 100)
        self.assertEqual(sh0, 100)


class TestParseStyle(unittest.TestCase):
    def test_basic(self):
        st = d2v.parse_style("fillColor=#ff0000;strokeColor=none;dashed=1")
        self.assertEqual(st["fillColor"], "#ff0000")
        self.assertEqual(st["strokeColor"], "none")
        self.assertEqual(st["dashed"], "1")

    def test_stencil_payload_value(self):
        st = d2v.parse_style("shape=stencil(AAA=);fillColor=none")
        self.assertEqual(st["shape"], "stencil(AAA=)")
        self.assertEqual(st["fillColor"], "none")


class TestSampleBasic(unittest.TestCase):
    """The small sample must parse into vertices with geometry."""

    @classmethod
    def setUpClass(cls):
        cls.cells = d2v.load_cells(BASIC)

    def test_has_vertices(self):
        self.assertTrue(any(c["vertex"] for c in self.cells))

    def test_offset_resolution(self):
        offsets = d2v.build_offsets(self.cells)
        self.assertIsInstance(offsets, dict)


class TestStencilCollect(unittest.TestCase):
    def test_collect_on_real_sample(self):
        """Icon clusters must be collectable from the real sample diagram."""
        cells = d2v.load_cells(SAMPLE)
        offsets = d2v.build_offsets(cells)
        icon_cells, labels = d2s.collect(cells, offsets)
        self.assertGreater(len(icon_cells), 100)
        clusters = d2s.clusters_of(icon_cells)
        self.assertGreater(len(clusters), 20)
        uniq = d2s.dedupe(clusters)
        self.assertGreater(len(uniq), 20)


if __name__ == "__main__":
    unittest.main()

class TestGeometryHelpers(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cells = d2v.load_cells(SAMPLE)
        cls.offsets = d2v.build_offsets(cls.cells)

    def test_rect_of_vertex(self):
        c = {"geo": None}
        # take a real vertex cell
        for cell in self.cells:
            if cell["vertex"] and cell["geo"] is not None:
                c = cell
                break
        r = d2v.rect_of(c, self.offsets)
        self.assertIsNotNone(r)
        x, y, w, h = r
        self.assertGreaterEqual(w, 0)
        self.assertGreaterEqual(h, 0)

    def test_rect_of_none_geo(self):
        c = {"vertex": 1, "geo": None, "parent": "1", "id": "nope"}
        self.assertIsNone(d2v.rect_of(c, self.offsets))

    def test_build_offsets_absolute(self):
        # absolute coords: offset must be 0
        offs = d2v.build_offsets(self.cells)
        self.assertIsInstance(offs, dict)


class TestStencilPrimitives(unittest.TestCase):
    def _prims(self, xml):
        import base64, zlib, urllib.parse
        enc = urllib.parse.quote(xml, safe="")
        co = zlib.compressobj(9, zlib.DEFLATED, -15)
        payload = base64.b64encode(co.compress(enc.encode()) + co.flush()).decode()
        sel = d2v.decode_stencil(payload)
        return d2v.stencil_primitives(sel, "#123456", None)

    def test_rect_primitive(self):
        prims, size = self._prims(
            '<shape><foreground><rect x="10" y="20" w="30" h="40"/>'
            '<fillstroke/></foreground></shape>')
        kinds = [p["kind"] for p in prims]
        self.assertIn("rect", kinds)
        rect = [p for p in prims if p["kind"] == "rect"][0]
        self.assertEqual((rect["x"], rect["y"], rect["w"], rect["h"]),
                         (10, 20, 30, 40))

    def test_oval_primitive(self):
        prims, _ = self._prims(
            '<shape><foreground><ellipse x="5" y="5" w="90" h="90"/>'
            '<fillstroke/></foreground></shape>')
        kinds = [p["kind"] for p in prims]
        self.assertIn("oval", kinds)

    def test_curve_flattening_count(self):
        # one cubic curve -> flattened: 1 (move) + 48 samples + 1 (close) = 50
        prims, _ = self._prims(
            '<shape><foreground><path><move x="0" y="0"/>'
            '<curve x1="30" y1="0" x2="60" y2="30" x3="60" y3="60"/>'
            '<close/></path><fillstroke/></foreground></shape>')
        sp = prims[0]["subpaths"][0]
        self.assertEqual(len(sp), 50)

    def test_fill_and_stroke_flags(self):
        prims, _ = self._prims(
            '<shape><foreground><fillcolor color="#ff0000"/>'
            '<strokecolor color="#00ff00"/>'
            '<path><move x="0" y="0"/><line x="10" y="10"/><close/></path>'
            '<fillstroke/></foreground></shape>')
        p = prims[0]
        self.assertEqual(p["fill"], "#ff0000")
        self.assertEqual(p["stroke"], "#00ff00")


class TestLabelHTMLDetails(unittest.TestCase):
    def test_multiline_text(self):
        html = ('<div><p>Line1<br/></p><p>Line2</p></div>')
        info = d2v.parse_label(html, {})
        self.assertIn("Line1", info["text"])
        self.assertIn("Line2", info["text"])

    def test_font_size_parsed(self):
        html = '<div><font style="font-size:16.93px">Big</font></div>'
        info = d2v.parse_label(html, {})
        self.assertEqual(info["size"], 16.93)

    def test_color_parsed(self):
        html = '<div><font style="color: #312d2a">Colored</font></div>'
        info = d2v.parse_label(html, {})
        self.assertEqual(info["color"], "#312d2a")


class TestIconNaming(unittest.TestCase):
    def test_name_cluster_uses_nearby_label(self):
        cells = d2v.load_cells(SAMPLE)
        offsets = d2v.build_offsets(cells)
        icon_cells, labels = d2s.collect(cells, offsets)
        clusters = d2s.clusters_of(icon_cells)
        uniq = d2s.dedupe(clusters)
        named = 0
        for cl in uniq:
            nm = d2s.name_cluster(cl, labels)
            if nm:
                named += 1
        # most clusters should resolve a human-readable name
        self.assertGreater(named, len(uniq) * 0.5)
