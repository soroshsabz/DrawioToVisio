"""drawio_to_visio - convert draw.io diagrams to native Visio files.

Library usage:

    from drawio_to_visio import convert, extract_stencils

    # full conversion (Visio required, Windows only)
    convert("input.drawio", "output.vsdx")

    # stencil pipeline: extract icons first, then convert using them
    extract_stencils("input.drawio", "icons.vssx")
    convert("input.drawio", "output.vsdx", stencil_vssx="icons.vssx")

Pure parsing helpers (cross-platform, no Visio required):

    from drawio_to_visio.core import load_cells, decode_stencil
"""

from .core import (
    LabelHTML,
    arc_points,
    build_offsets,
    decode_stencil,
    hexrgb,
    load_cells,
    main as convert,
    parse_label,
    parse_style,
    rect_of,
    scale_shift_prim,
    stencil_primitives,
    transform_segs,
)
from .stencils import (
    clusters_of,
    collect,
    dedupe,
    draw_cluster_into,
    main as _extract_stencils_impl,
    name_cluster,
    overlap,
)


def extract_stencils(src, out, real_connectors=True):
    """Extract every icon of *src* into the .vssx stencil set *out*.

    With ``real_connectors=True`` slim directional-arrow glyphs are
    classified as connectors instead of icons (pairs with
    ``convert(..., real_connectors=True)``)."""
    return _extract_stencils_impl(src, out,
                                  real_connectors=real_connectors)

__version__ = "1.0.0"

__all__ = [
    "convert",
    "extract_stencils",
    "load_cells",
    "decode_stencil",
    "stencil_primitives",
    "parse_label",
    "parse_style",
    "parse_label",
    "build_offsets",
    "rect_of",
    "arc_points",
    "hexrgb",
    "scale_shift_prim",
    "transform_segs",
    "LabelHTML",
    "collect",
    "clusters_of",
    "dedupe",
    "name_cluster",
    "overlap",
    "draw_cluster_into",
    "__version__",
]
