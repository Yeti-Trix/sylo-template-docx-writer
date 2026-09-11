#!/usr/bin/env python3
"""Convert Strict OOXML .docx to Transitional so python-docx can open it."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

STRICT_TO_TRANSITIONAL = {
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    ),
    "http://purl.oclc.org/ooxml/drawingml/main": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": (
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    ),
    "http://purl.oclc.org/ooxml/wordprocessingml/main": (
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ),
    "http://purl.oclc.org/ooxml/markup-compatibility/2006": (
        "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ),
}


def patch_xml(text: str) -> str:
    out = text
    for strict, trans in STRICT_TO_TRANSITIONAL.items():
        out = out.replace(strict, trans)
    # Strict paths omit /2006/ segment in some tags
    out = out.replace("/wordprocessingml/main", "/wordprocessingml/2006/main")
    out = out.replace("/drawingml/main", "/drawingml/2006/main")
    return out


def convert_docx(src: Path, dest: Path | None = None) -> Path:
    dest = dest or src
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        extracted = tmp_path / "doc"
        with zipfile.ZipFile(src, "r") as zin:
            zin.extractall(extracted)
        for xml_file in extracted.rglob("*.xml"):
            text = xml_file.read_text(encoding="utf-8")
            patched = patch_xml(text)
            if patched != text:
                xml_file.write_text(patched, encoding="utf-8")
        for rels in extracted.rglob("*.rels"):
            text = rels.read_text(encoding="utf-8")
            patched = patch_xml(text)
            if patched != text:
                rels.write_text(patched, encoding="utf-8")
        out = dest if dest != src else tmp_path / "fixed.docx"
        if dest == src:
            backup = src.with_suffix(".docx.bak")
            shutil.copy2(src, backup)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for file in extracted.rglob("*"):
                if file.is_file():
                    zout.write(file, file.relative_to(extracted).as_posix())
        if dest == src:
            shutil.move(out, src)
            return src
        shutil.copy2(out, dest)
        return dest


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch Strict OOXML docx for python-docx.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--in-place", action="store_true")
    args = parser.parse_args()
    if not args.docx.is_file():
        print(json.dumps({"error": f"Not found: {args.docx}"}))
        return 1
    try:
        path = convert_docx(args.docx, args.docx if args.in_place else None)
        print(json.dumps({"ok": True, "path": str(path.resolve())}))
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": str(exc)}))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
