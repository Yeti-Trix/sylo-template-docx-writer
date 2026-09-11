"""PDF page render into manual project inputs/.

DOCX image extraction moved to the sylo-docx package
(``extract_docx_images`` tool) — use that for pulling pictures from
source .docx files, with ``output_dir`` set to the project inputs/.
"""

from __future__ import annotations

import re
from pathlib import Path

DEFAULT_PDF_DPI = 150


def resolve_project_inputs_dir(
    *,
    draft_path: Path | None = None,
    project_root: Path | None = None,
    project_id: str | None = None,
    inputs_dir: Path | None = None,
) -> Path:
    """Return ``projects/<id>/inputs`` (created if missing)."""
    if inputs_dir is not None:
        out = Path(inputs_dir).resolve()
    elif draft_path is not None:
        out = Path(draft_path).resolve().parent / "inputs"
    elif project_root is not None and project_id:
        out = (Path(project_root).resolve() / "projects" / project_id / "inputs")
    else:
        raise ValueError(
            "Provide draft_path, inputs_dir, or project_root + project_id for the inputs folder."
        )
    out.mkdir(parents=True, exist_ok=True)
    return out


def _safe_stem(name: str) -> str:
    stem = re.sub(r"[^\w.\-]+", "-", name.strip()).strip("-")
    return stem or "image"


def _parse_four_floats(raw: str, label: str) -> tuple[float, float, float, float]:
    parts = [p.strip() for p in raw.split(",")]
    if len(parts) != 4:
        raise ValueError(f"{label} must be four comma-separated numbers (got {raw!r})")
    try:
        vals = tuple(float(p) for p in parts)
    except ValueError as exc:
        raise ValueError(f"{label} must be numeric: {raw!r}") from exc
    return vals  # type: ignore[return-value]


def _clip_rect_from_norm(page, norm: tuple[float, float, float, float]):
    import fitz

    x0, y0, x1, y1 = norm
    for v in norm:
        if v < 0 or v > 1:
            raise ValueError("clip_norm values must be between 0 and 1")
    if x0 >= x1 or y0 >= y1:
        raise ValueError("clip_norm must have x0 < x1 and y0 < y1")
    r = page.rect
    return fitz.Rect(
        r.x0 + x0 * r.width,
        r.y0 + y0 * r.height,
        r.x0 + x1 * r.width,
        r.y0 + y1 * r.height,
    )


def _clip_rect_from_pt(page, pt: tuple[float, float, float, float]):
    import fitz

    x0, y0, x1, y1 = pt
    if x0 >= x1 or y0 >= y1:
        raise ValueError("clip_pt must have x0 < x1 and y0 < y1")
    rect = fitz.Rect(x0, y0, x1, y1)
    if not rect.intersects(page.rect):
        raise ValueError("clip_pt does not intersect the page")
    return rect & page.rect


def render_pdf_page(
    pdf_path: Path,
    page: int,
    inputs_dir: Path,
    *,
    dpi: int = DEFAULT_PDF_DPI,
    name: str | None = None,
    clip_norm: str | None = None,
    clip_pt: str | None = None,
) -> dict:
    """Render one PDF page (optional crop) to PNG under ``inputs_dir``."""
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise RuntimeError(
            "PyMuPDF is required for PDF render. "
            "pip install -r packages/sylo-template-docx-writer/scripts/requirements.txt"
        ) from exc

    pdf_path = pdf_path.resolve()
    if not pdf_path.is_file():
        raise ValueError(f"PDF not found: {pdf_path}")

    inputs_dir = Path(inputs_dir).resolve()
    inputs_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    try:
        if page < 1 or page > doc.page_count:
            raise ValueError(f"Page {page} out of range (1–{doc.page_count})")

        pg = doc.load_page(page - 1)
        clip = None
        if clip_norm and clip_pt:
            raise ValueError("Use only one of clip_norm or clip_pt")
        if clip_norm:
            clip = _clip_rect_from_norm(pg, _parse_four_floats(clip_norm, "clip_norm"))
        elif clip_pt:
            clip = _clip_rect_from_pt(pg, _parse_four_floats(clip_pt, "clip_pt"))

        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        pix = pg.get_pixmap(matrix=matrix, clip=clip, alpha=False)

        label = _safe_stem(name) if name else pdf_path.stem
        suffix = "-crop" if clip is not None else ""
        out_name = f"{_safe_stem(label)}-p{page:03d}{suffix}-{dpi}dpi.png"
        out_path = inputs_dir / out_name
        if out_path.exists():
            n = 2
            while out_path.exists():
                out_path = inputs_dir / f"{_safe_stem(label)}-p{page:03d}{suffix}-{dpi}dpi-{n}.png"
                n += 1

        pix.save(out_path)
        rel = f"inputs/{out_path.name}"

        return {
            "pdf_path": str(pdf_path),
            "page": page,
            "page_count": doc.page_count,
            "dpi": dpi,
            "clipped": clip is not None,
            "width_px": pix.width,
            "height_px": pix.height,
            "png_path": str(out_path.resolve()),
            "relative_markdown": rel,
            "markdown_snippet": f"![{label}]({rel})",
        }
    finally:
        doc.close()
