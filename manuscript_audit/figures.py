"""
Visual Figure Audit & Dark-Mode RGBA Transparency Repair
========================================================
Extracts and inspects embedded figures in `word/media/` inside `.docx` packages.
Detects transparent RGBA PNG backgrounds (`alpha < 255`) where black axes, ticks,
and text labels become illegible in dark-mode readers, and composites them onto
a solid white RGB `(255, 255, 255)` canvas.
"""

import os
import shutil
import tempfile
import zipfile
from typing import Any, Dict, List, Optional

try:
    from PIL import Image
    import numpy as np
    HAVE_IMAGING = True
except ImportError:
    HAVE_IMAGING = False


def composite_rgba_pngs_in_dir(media_dir: str) -> List[str]:
    """
    Composites any PNG in `media_dir` containing transparent pixels (`alpha < 255`)
    onto a solid white RGB `(255, 255, 255)` background in place.
    Returns the list of repaired filenames.
    """
    repaired: List[str] = []
    if not HAVE_IMAGING or not os.path.isdir(media_dir):
        return repaired

    for fname in sorted(os.listdir(media_dir)):
        if not fname.lower().endswith(".png"):
            continue
        fpath = os.path.join(media_dir, fname)
        with Image.open(fpath) as img:
            if img.mode == "RGBA":
                arr = np.array(img)
                if arr[:, :, 3].min() < 255:
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[3])
                    bg.save(fpath, "PNG", compress_level=1)
                    repaired.append(fname)
    return repaired


def audit_and_repair_docx_figures(
    input_docx: str,
    output_docx: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Audits all images inside `word/media/` of `input_docx` for RGBA transparency
    and low pixel dimensions, and optionally writes a repaired `.docx` with solid
    white RGB backgrounds.
    """
    if not HAVE_IMAGING:
        raise RuntimeError("Pillow and numpy are required for figure auditing (`pip install Pillow numpy`).")

    tmp_dir = tempfile.mkdtemp(prefix="manuscript_fig_audit_")
    try:
        with zipfile.ZipFile(input_docx, "r") as z:
            z.extractall(tmp_dir)

        media_dir = os.path.join(tmp_dir, "word", "media")
        figures_report: List[Dict[str, Any]] = []
        repaired_files: List[str] = []

        if os.path.isdir(media_dir):
            for fname in sorted(os.listdir(media_dir)):
                fpath = os.path.join(media_dir, fname)
                if not os.path.isfile(fpath):
                    continue
                try:
                    with Image.open(fpath) as img:
                        mode = img.mode
                        width, height = img.size
                        has_transparency = False
                        min_alpha = 255
                        if mode == "RGBA":
                            alpha = np.array(img)[:, :, 3]
                            min_alpha = int(alpha.min())
                            has_transparency = min_alpha < 255
                        figures_report.append({
                            "filename": fname,
                            "format": img.format,
                            "mode": mode,
                            "width_px": width,
                            "height_px": height,
                            "has_transparent_pixels": has_transparency,
                            "min_alpha": min_alpha,
                            "low_resolution_warning": min(width, height) < 800,
                        })
                except Exception as exc:
                    figures_report.append({"filename": fname, "error": str(exc)})

            if output_docx:
                repaired_files = composite_rgba_pngs_in_dir(media_dir)
                if os.path.exists(output_docx):
                    os.remove(output_docx)
                with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
                    for foldername, _, filenames in os.walk(tmp_dir):
                        for filename in filenames:
                            filepath = os.path.join(foldername, filename)
                            arcname = os.path.relpath(filepath, tmp_dir)
                            zout.write(filepath, arcname)

        return {
            "input_docx": input_docx,
            "output_docx": output_docx,
            "figure_count": len(figures_report),
            "repaired_rgba_files": repaired_files,
            "figures": figures_report,
        }
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
