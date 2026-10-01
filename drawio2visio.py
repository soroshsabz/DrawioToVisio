#!/usr/bin/env python
"""Convert a draw.io diagram to a native, editable Visio .vsdx.

Thin CLI wrapper - the implementation lives in the installable
``drawio_to_visio`` package (see pyproject.toml).
"""

import sys

from drawio_to_visio.cli import main_visio

if __name__ == "__main__":
    sys.exit(main_visio())
