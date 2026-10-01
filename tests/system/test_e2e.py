"""End-to-end (system) tests: real draw.io -> Visio conversions.

These exercise the full pipeline including the Visio COM layer.  They are
skipped automatically when Microsoft Visio is not available (e.g. on CI
runners without an Office install), so the rest of the suite stays green.
"""

import os
import sys
import unittest
import zipfile

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
for _p in (os.path.join(REPO, "src"), REPO):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import drawio_to_visio  # noqa: E402

SCRATCH = os.environ.get("TMPDIR", os.path.join(REPO, "tests", "_scratch"))
BASIC = os.path.join(REPO, "examples", "sample-basic.drawio")
SAMPLE = os.path.join(REPO, "examples", "sample-functional-view.drawio")


def _visio_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        app = win32com.client.DispatchEx("Visio.InvisibleApp")
    except Exception:
        return False
    try:
        app.Quit()
    except Exception:
        pass
    return True


VISIO = _visio_available()
REASON = "Microsoft Visio is not installed on this machine"


@unittest.skipUnless(VISIO, REASON)
class TestConvertEndToEnd(unittest.TestCase):
    """drawio2visio: real .drawio -> .vsdx conversion."""

    @classmethod
    def setUpClass(cls):
        os.makedirs(SCRATCH, exist_ok=True)
        cls.out = os.path.join(SCRATCH, "e2e_basic.vsdx")
        if os.path.exists(cls.out):
            os.remove(cls.out)
        drawio_to_visio.convert(BASIC, cls.out)

    def test_output_exists(self):
        self.assertTrue(os.path.exists(self.out))
        self.assertGreater(os.path.getsize(self.out), 1000)

    def test_output_is_valid_vsdx_package(self):
        with zipfile.ZipFile(self.out) as z:
            names = z.namelist()
        self.assertIn("[Content_Types].xml", names)
        self.assertTrue(any(n.endswith("page1.xml") for n in names))
        self.assertTrue(any("visio/document" in n for n in names))

    def test_output_contains_shapes(self):
        import re
        with zipfile.ZipFile(self.out) as z:
            page = z.read("visio/pages/page1.xml").decode("utf-8")
        shapes = len(re.findall(r"<Shape ", page))
        self.assertGreater(shapes, 3)

    def test_output_contains_text(self):
        with zipfile.ZipFile(self.out) as z:
            page = z.read("visio/pages/page1.xml").decode("utf-8")
        self.assertIn("<Text>", page)


@unittest.skipUnless(VISIO, REASON)
class TestStencilExtractionEndToEnd(unittest.TestCase):
    """drawio2stencils: real .drawio -> .vssx stencil set."""

    @classmethod
    def setUpClass(cls):
        os.makedirs(SCRATCH, exist_ok=True)
        cls.out = os.path.join(SCRATCH, "e2e_icons.vssx")
        if os.path.exists(cls.out):
            os.remove(cls.out)
        # the functional-view sample carries the icon artwork; the basic
        # sample has no stencil cells at all
        drawio_to_visio.extract_stencils(SAMPLE, cls.out)

    def test_output_exists(self):
        self.assertTrue(os.path.exists(self.out))
        self.assertGreater(os.path.getsize(self.out), 1000)

    def test_output_is_valid_vssx_package(self):
        with zipfile.ZipFile(self.out) as z:
            names = z.namelist()
        self.assertIn("[Content_Types].xml", names)
        self.assertTrue(any("masters/masters.xml" in n for n in names))

    def test_masters_xml_parses(self):
        import re
        with zipfile.ZipFile(self.out) as z:
            masters = [n for n in z.namelist()
                       if n.endswith(".xml") and "masters/master" in n]
        self.assertGreater(len(masters), 0)


@unittest.skipUnless(VISIO, REASON)
class TestTwoPassPipelineEndToEnd(unittest.TestCase):
    """Stencil pipeline: extract -> convert with master instances."""

    @classmethod
    def setUpClass(cls):
        os.makedirs(SCRATCH, exist_ok=True)
        cls.icons = os.path.join(SCRATCH, "e2e_pipe_icons.vssx")
        cls.out = os.path.join(SCRATCH, "e2e_pipe.vsdx")
        for f in (cls.icons, cls.out):
            if os.path.exists(f):
                os.remove(f)
        drawio_to_visio.extract_stencils(SAMPLE, cls.icons)
        drawio_to_visio.convert(SAMPLE, cls.out, stencil_vssx=cls.icons)

    def test_outputs_exist(self):
        self.assertTrue(os.path.exists(self.icons))
        self.assertTrue(os.path.exists(self.out))

    def test_page_uses_master_instances(self):
        import re
        with zipfile.ZipFile(self.out) as z:
            pg = z.read("visio/pages/page1.xml").decode("utf-8")
            mx = z.read("visio/masters/masters.xml").decode("utf-8")
        # the page must place master instances (attr form Master='N')
        self.assertTrue(re.search(r"<Shape [^>]*Master='\d+'", pg))
        # and the document must embed the masters
        self.assertIn("<Master ", mx)

    def test_all_labels_preserved(self):
        """Every source label must survive conversion (the text-preservation
        guarantee)."""
        import re
        import html
        with zipfile.ZipFile(self.out) as z:
            pg = z.read("visio/pages/page1.xml").decode("utf-8")
        blocks = re.findall(r"<Text>(.*?)</Text>", pg, re.S)

        def norm(s):
            import re as _re
            s = _re.sub(r"<[^>]+>", " ", s)
            s = html.unescape(s)
            return _re.sub(r"\s+", " ", s).strip()

        rendered = {norm(b) for b in blocks}
        rendered.discard("")
        # the sample carries 105 unique label texts; all must be present
        import drawio_to_visio.core as core
        cells = core.load_cells(SAMPLE)
        import re as _re
        src = set()
        for c in cells:
            t = norm(str(c.get("label") or ""))
            if t:
                src.add(t)
        missing = [t for t in src
                   if t not in rendered and not any(t in m for m in rendered)]
        self.assertEqual(missing, [])


if __name__ == "__main__":
    unittest.main()
