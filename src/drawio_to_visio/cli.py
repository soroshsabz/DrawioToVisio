"""Command-line entry points for drawio_to_visio.

Installed console scripts (see pyproject.toml):
    drawio2visio    input.drawio output.vsdx [icons.vssx]
    drawio2stencils input.drawio icons.vssx
"""

import argparse
import os
import sys


def main_visio(argv=None):
    """CLI: convert a .drawio file to an editable .vsdx."""
    parser = argparse.ArgumentParser(
        prog="drawio2visio",
        description="Convert a draw.io diagram to a native, fully "
                    "editable Microsoft Visio .vsdx file.")
    parser.add_argument("input", help="input .drawio file")
    parser.add_argument("output", help="output .vsdx file")
    parser.add_argument(
        "stencils", nargs="?",
        help="optional .vssx stencil set: enables the two-pass pipeline "
             "where every icon becomes a reusable master instance")
    args = parser.parse_args(argv)

    from .core import main as convert
    convert(os.path.abspath(args.input),
            os.path.abspath(args.output),
            os.path.abspath(args.stencils) if args.stencils else None)
    return 0


def main_stencils(argv=None):
    """CLI: extract every icon of a .drawio file into a .vssx stencil set."""
    parser = argparse.ArgumentParser(
        prog="drawio2stencils",
        description="Extract every icon from a draw.io diagram into a "
                    "reusable Microsoft Visio stencil set (.vssx).")
    parser.add_argument("input", help="input .drawio file")
    parser.add_argument("output", help="output .vssx stencil set")
    args = parser.parse_args(argv)

    from .stencils import main as extract
    extract(os.path.abspath(args.input), os.path.abspath(args.output))
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main_visio())
