"""Quick template hygiene report."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from manual_builder_lib import open_document  # noqa: E402


def main() -> int:
    path = Path(sys.argv[1])
    doc = open_document(path)
    empty = num_orphan = pb = ph = 0
    orphans: list[tuple] = []
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        has_num = bool(p._element.xpath(".//w:numPr"))
        has_pb = bool(p._element.xpath('.//w:br[@w:type="page"]'))
        if not t:
            empty += 1
        if has_num and not t:
            num_orphan += 1
        if has_pb:
            pb += 1
        if t.startswith("{S") or t.startswith("{IMG"):
            ph += 1
        if (not t and has_num) or (has_pb and not t):
            orphans.append((i, t[:30], has_num, has_pb))
    print(f"file: {path.name}")
    print(f"paragraphs={len(doc.paragraphs)} tables={len(doc.tables)}")
    print(f"empty={empty} num_orphan={num_orphan} page_breaks={pb} placeholders={ph}")
    print("orphan samples:", len(orphans))
    for row in orphans[:30]:
        print(" ", row)
    for ti, tbl in enumerate(doc.tables):
        rows = len(tbl.rows)
        hdr = [c.text.strip()[:18] for c in tbl.rows[0].cells] if rows else []
        print(f"table[{ti}] {rows} rows cols={len(tbl.columns)} hdr={hdr}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
