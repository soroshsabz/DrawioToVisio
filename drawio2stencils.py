#!/usr/bin/env python
"""Extract every icon from a draw.io diagram into a Visio stencil set.

Thin CLI wrapper - the implementation lives in the installable
``drawio_to_visio`` package (see pyproject.toml).
"""

import sys

from drawio_to_visio.cli import main_stencils

if __name__ == "__main__":
    sys.exit(main_stencils())
