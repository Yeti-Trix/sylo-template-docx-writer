"""
Build controls-template.docx from a filled controls manual.

Keeps cover, warnings; removes body text between section anchors; inserts or
preserves {SNN_SECTION} placeholders (H1 and tables from staging).
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from manual_builder_lib import open_document, replace_substring_preserving_runs

SECTIONS: list[tuple[str, str]] = [
    (r"^Terms$", "S01_SECTION"),
    (r"^2\.0\s+Introduction", "S02_SECTION"),
    (r"^3\.0\s+Machine Capacities", "S03_SECTION"),
    (r"^4\.0\s+Startup", "S04_SECTION"),
    (r"^5\.0\s+Shutdown", "S05_SECTION"),
    (r"^6\.0\s+(HMI|Operation)", "S06_SECTION"),
    (r"^7\.0\s+Operator Maintenance", "S07_SECTION"),
    (r"^8\.0\s+Safety", "S08_SECTION"),
    (r"^9\.0\s+Physical Controls", "S09_SECTION"),
    (r"^16\.0\s+Operator Maintenance", "S07_SECTION"),
    (r"^17\.0\s+Safety", "S08_SECTION"),
    (r"^18\.0\s+Physical Controls", "S09_SECTION"),
    (r"^19\.0\s+Installation", "S19_SECTION"),
    (r"^20\.0\s+Control Hardware", "S20_SECTION"),
    (r"^11\.0\s+Control Hardware", "S11_SECTION"),
    (r"^12\.0\s+Passwords", "S12_SECTION"),
    (r"^21\.0\s+Passwords", "S12_SECTION"),
]

HMI_RENUMBER: list[tuple[str, str]] = [
    (r"^6\.0\s+Main Dashboard", "6.0\tHMI Operator Screens"),
    (r"^7\.0\s+I/O Status", "6.2\tI/O Status Screen"),
    (r"^8\.0\s+Alarm", "6.3\tAlarm Screen"),
    (r"^9\.0\s+Utility", "6.4\tUtility Screen"),
    (r"^10\.0\s+Servo Utility", "6.5\tServo Utility Screen"),
    (r"^11\.0\s+Trend", "6.6\tTrend Screens"),
    (r"^12\.0\s+Continuous Mode", "6.7\tContinuous Mode Screen"),
    (r"^13\.0\s+Recipe Edit", "6.8\tRecipe Edit Screen"),
    (r"^14\.0\s+Weld I/O", "6.9\tWeld I/O Screen"),
]

IMG_BY_SLUG: dict[str, str] = {
    "S06_01_MAIN_DASHBOARD": "IMG_S06_01_MAIN_DASHBOARD",
    "S06_02_IO_STATUS": "IMG_S06_02_IO_STATUS",
    "S06_03_ALARMS": "IMG_S06_03_ALARMS",
    "S06_04_UTILITY": "IMG_S06_04_UTILITY",
    "S06_05_SERVO_UTILITY": "IMG_S06_05_SERVO_UTILITY",
    "S06_06_TRENDS": "IMG_S06_06_TRENDS",
    "S06_07_CONTINUOUS_MODE": "IMG_S06_07_CONTINUOUS_MODE",
    "S06_08_RECIPE_EDIT": "IMG_S06_08_RECIPE_EDIT",
    "S06_09_WELD_IO": "IMG_S06_09_WELD_IO",
}

COVER_TAGS = {
    "#12345-001": "{S00_COVER_JOB_NUMBER}",
    "12345-001": "{S00_COVER_JOB_NUMBER}",
    "HV Boom and Column Welder": "{S00_COVER_PROJECT_TITLE}",
    "Controls Manual": "{S00_COVER_MANUAL_TYPE}",
}


def _delete_paragraph(para) -> None:
    el = para._element
    el.getparent().remove(el)


def _is_toc_line(text: str) -> bool:
    return bool(re.match(r"^\d+\.\d+\t.+?\t\d+$", text)) or bool(
        re.match(r"^\d+\.0\t.+?\t\d+$", text)
    )


def restructure_hmi_headings(doc) -> int:
    changed = 0
    insert_61_after: int | None = None
    for i, para in enumerate(doc.paragraphs):
        t = para.text.strip()
        for pattern, replacement in HMI_RENUMBER:
            if re.match(pattern, t):
                para.text = replacement
                changed += 1
                if replacement.startswith("6.0\tHMI"):
                    insert_61_after = i
                break

    if insert_61_after is not None:
        host = doc.paragraphs[insert_61_after]
        parent = host._element.getparent()  # noqa: SLF001
        pos = parent.index(host._element) + 1  # noqa: SLF001
        new_para = doc.add_paragraph("6.1\tMain Dashboard")
        parent.insert(pos, new_para._element)  # noqa: SLF001
        changed += 1

    return changed


def prune_orphan_paragraphs(doc) -> int:
    removed = 0
    for para in list(doc.paragraphs):
        if not para.text.strip():
            _delete_paragraph(para)
            removed += 1
    return removed


SLUG_IN_PARA_RE = re.compile(r"\{(S\d{2}_SECTION)\}", re.I)


def seed_placeholders(doc, *, clear_body: bool = True) -> list[str]:
    header_indices: list[tuple[int, str]] = []
    seen_slugs: set[str] = set()

    for i, para in enumerate(doc.paragraphs):
        t = para.text.strip()
        if _is_toc_line(t):
            continue
        slug_match = SLUG_IN_PARA_RE.search(para.text)
        if slug_match:
            slug = slug_match.group(1).upper()
            if slug not in seen_slugs:
                header_indices.append((i, slug))
                seen_slugs.add(slug)
            continue
        for pattern, slug in SECTIONS:
            if slug in seen_slugs:
                continue
            if re.match(pattern, t):
                header_indices.append((i, slug))
                seen_slugs.add(slug)
                break

    inserted: list[str] = []
    for idx in range(len(header_indices) - 1, -1, -1):
        header_i, slug = header_indices[idx]
        next_header = header_indices[idx + 1][0] if idx + 1 < len(header_indices) else len(doc.paragraphs)

        if clear_body:
            for j in range(next_header - 1, header_i, -1):
                _delete_paragraph(doc.paragraphs[j])

        host = doc.paragraphs[header_i]
        tag = f"{{{slug}}}"
        if tag not in host.text:
            parent = host._element.getparent()  # noqa: SLF001
            pos = parent.index(host._element) + 1  # noqa: SLF001

            lines = [tag]
            img_slug = IMG_BY_SLUG.get(slug)
            if img_slug:
                lines.append(f"{{{img_slug}}}")

            for line in reversed(lines):
                new_para = doc.add_paragraph(line)
                parent.insert(pos, new_para._element)  # noqa: SLF001

            inserted.insert(0, slug)
            if img_slug:
                inserted.append(img_slug)
        else:
            inserted.insert(0, slug)
            img_slug = IMG_BY_SLUG.get(slug)
            if img_slug and f"{{{img_slug}}}" not in host.text:
                parent = host._element.getparent()  # noqa: SLF001
                pos = parent.index(host._element) + 1  # noqa: SLF001
                new_para = doc.add_paragraph(f"{{{img_slug}}}")
                parent.insert(pos, new_para._element)  # noqa: SLF001
                inserted.append(img_slug)

    return inserted


def apply_cover_placeholders(doc) -> None:
    for para in doc.paragraphs:
        for old, new in COVER_TAGS.items():
            if old in para.text:
                replace_substring_preserving_runs(para, old, new)


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepare controls manual template")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--keep-body", action="store_true")
    parser.add_argument("--no-hmi-renumber", action="store_true")
    args = parser.parse_args()

    if not args.source.is_file():
        print(f"Source not found: {args.source}", file=sys.stderr)
        return 1

    args.output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.source, args.output)

    doc = open_document(args.output)
    if not args.no_hmi_renumber:
        print(f"Renumbered {restructure_hmi_headings(doc)} HMI section headings")
    apply_cover_placeholders(doc)
    slugs = seed_placeholders(doc, clear_body=not args.keep_body)
    removed = prune_orphan_paragraphs(doc)
    print(f"Removed {removed} empty/orphan paragraphs")
    doc.save(str(args.output))

    print(f"Wrote {args.output} with {len(slugs)} prose/image placeholders")
    for s in slugs:
        print(f"  {{{s}}}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
