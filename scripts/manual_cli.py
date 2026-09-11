#!/usr/bin/env python3
"""CLI entry for sylo-template-docx-writer Python helpers."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Allow imports when run as script from scripts/
sys.path.insert(0, str(Path(__file__).resolve().parent))

import manual_builder_lib as mb  # noqa: E402
import manual_media as mm  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Sylo manual creator docx tools")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_disc = sub.add_parser("discover")
    p_disc.add_argument("docx", type=Path)

    p_start = sub.add_parser("start")
    p_start.add_argument("project_root", type=Path)
    p_start.add_argument("project_id")
    p_start.add_argument("template", type=Path)
    p_start.add_argument("--no-fix-strict", action="store_true")

    p_ensure = sub.add_parser("ensure-workspace")
    p_ensure.add_argument("project_root", type=Path)

    p_cover = sub.add_parser("fill-cover")
    p_cover.add_argument("docx", type=Path)
    p_cover.add_argument("--project-title", default=None)
    p_cover.add_argument("--job-number", default=None)
    p_cover.add_argument("--manual-type", default=None)

    p_inj = sub.add_parser("inject")
    p_inj.add_argument("docx", type=Path)
    p_inj.add_argument("placeholder")
    p_inj.add_argument("content", nargs="?", default="")
    p_inj.add_argument("--content-file", type=Path, default=None)

    p_img = sub.add_parser("image")
    p_img.add_argument("docx", type=Path)
    p_img.add_argument("placeholder")
    p_img.add_argument("image_path", type=Path)
    p_img.add_argument("--width", type=float, default=6.5)
    p_img.add_argument("--caption", default=None, help="Figure caption text (alt text); filename used if omitted")

    p_rev = sub.add_parser("revision")
    p_rev.add_argument("docx", type=Path)
    p_rev.add_argument("--rev", required=True)
    p_rev.add_argument("--date", required=True)
    p_rev.add_argument("--description", required=True)
    p_rev.add_argument("--by", required=True)

    p_state = sub.add_parser("state")
    p_state.add_argument("state_json", type=Path)
    p_state.add_argument("--mark-done", default=None)

    p_access = sub.add_parser("access-table")
    p_access.add_argument("docx", type=Path)
    p_access.add_argument("rows_json", type=Path, help="JSON array of access row objects")

    p_access_row = sub.add_parser("access-row")
    p_access_row.add_argument("docx", type=Path)
    p_access_row.add_argument("--system", required=True)
    p_access_row.add_argument("--purpose", required=True)
    p_access_row.add_argument("--username", default="")
    p_access_row.add_argument("--password", default="")
    p_access_row.add_argument("--notes", default="")

    p_hmi_stage = sub.add_parser("hmi-stage")
    p_hmi_stage.add_argument("state_json", type=Path)
    p_hmi_stage.add_argument("--title", required=True)
    p_hmi_stage.add_argument("--body-file", type=Path, required=True)
    p_hmi_stage.add_argument("--image", default=None)
    p_hmi_stage.add_argument("--placeholder", default=mb.HMI_DEFAULT_PLACEHOLDER)

    p_hmi_finalize = sub.add_parser("hmi-finalize")
    p_hmi_finalize.add_argument("state_json", type=Path)
    p_hmi_finalize.add_argument("--draft", type=Path, default=None)
    p_hmi_finalize.add_argument("--placeholder", default=mb.HMI_DEFAULT_PLACEHOLDER)
    p_hmi_finalize.add_argument("--width", type=float, default=6.5)

    p_hmi_staging = sub.add_parser("hmi-staging")
    p_hmi_staging.add_argument("state_json", type=Path)
    p_hmi_staging.add_argument("--placeholder", default=mb.HMI_DEFAULT_PLACEHOLDER)

    p_stage = sub.add_parser("stage-section")
    p_stage.add_argument("state_json", type=Path)
    p_stage.add_argument("placeholder")
    p_stage.add_argument("--body-file", type=Path, required=True)

    p_revise = sub.add_parser("revise")
    p_revise.add_argument("state_json", type=Path)
    p_revise.add_argument("placeholder")
    p_revise.add_argument("--body-file", type=Path, required=True)
    p_revise.add_argument("--draft", type=Path, default=None)

    p_list_staging = sub.add_parser("list-staging")
    p_list_staging.add_argument("state_json", type=Path)

    p_build = sub.add_parser("build-draft")
    p_build.add_argument("state_json", type=Path)
    p_build.add_argument("--draft", type=Path, default=None)
    p_build.add_argument("--skip-hmi", action="store_true")

    p_repair = sub.add_parser("repair-lists")
    p_repair.add_argument("docx", type=Path)

    p_add = sub.add_parser("add-section")
    p_add.add_argument("state_json", type=Path)
    p_add.add_argument("--number", required=True, help='e.g. "12.0"')
    p_add.add_argument("--title", required=True, help='e.g. "Cybersecurity"')
    p_add.add_argument("--placeholder", required=True, help='e.g. S12_SECTION')
    p_add.add_argument("--after", required=True, help='Insert after this slug, e.g. S11_SECTION')
    p_add.add_argument("--draft", type=Path, default=None)

    p_pdf = sub.add_parser("render-pdf")
    p_pdf.add_argument("pdf", type=Path)
    p_pdf.add_argument("--page", type=int, required=True, help="1-based page number")
    p_pdf.add_argument("--draft", type=Path, default=None)
    p_pdf.add_argument("--project-root", type=Path, default=None)
    p_pdf.add_argument("--project-id", default=None)
    p_pdf.add_argument("--inputs-dir", type=Path, default=None)
    p_pdf.add_argument("--dpi", type=int, default=mm.DEFAULT_PDF_DPI)
    p_pdf.add_argument("--name", default=None, help="Output filename stem")
    p_pdf.add_argument(
        "--clip-norm",
        default=None,
        help="Crop: left,top,right,bottom as fractions 0–1 (top-left origin)",
    )
    p_pdf.add_argument(
        "--clip-pt",
        default=None,
        help="Crop: x0,y0,x1,y1 in PDF points (top-left origin)",
    )

    args = parser.parse_args()

    try:
        if args.cmd == "discover":
            result = {"placeholders": mb.discover_placeholders(args.docx)}
        elif args.cmd == "ensure-workspace":
            result = mb.ensure_manual_workspace(args.project_root)
        elif args.cmd == "start":
            result = mb.start_project(
                args.project_root,
                args.project_id,
                args.template,
                fix_strict=not args.no_fix_strict,
            )
        elif args.cmd == "fill-cover":
            result = mb.fill_cover_placeholders(
                args.docx,
                project_title=args.project_title,
                job_number=args.job_number,
                manual_type=args.manual_type,
            )
        elif args.cmd == "inject":
            body = args.content
            if args.content_file:
                body = args.content_file.read_text(encoding="utf-8")
            if not body.strip():
                raise ValueError("inject requires content or --content-file")
            result = mb.inject_section(args.docx, args.placeholder, body)
            state_guess = args.docx.parent / "state.json"
            if state_guess.is_file():
                mb.mark_placeholder_done(state_guess, args.placeholder)
                result["state_updated"] = True
        elif args.cmd == "image":
            result = mb.insert_image(
                args.docx,
                args.placeholder,
                args.image_path,
                width_inches=args.width,
                caption=args.caption,
            )
            state_guess = args.docx.parent / "state.json"
            if state_guess.is_file():
                mb.mark_placeholder_done(state_guess, args.placeholder)
                result["state_updated"] = True
        elif args.cmd == "revision":
            result = mb.append_revision(
                args.docx, args.rev, args.date, args.description, args.by
            )
        elif args.cmd == "state":
            result = mb.load_state(args.state_json)
            if args.mark_done:
                mb.mark_placeholder_done(args.state_json, args.mark_done)
                result = mb.load_state(args.state_json)
        elif args.cmd == "access-table":
            rows = json.loads(args.rows_json.read_text(encoding="utf-8"))
            result = mb.replace_access_table(args.docx, rows)
            state_guess = args.docx.parent / "state.json"
            if state_guess.is_file():
                state = mb.load_state(state_guess)
                state["access_table"] = "filled"
                state_guess.write_text(json.dumps(state, indent=2), encoding="utf-8")
                result["state_updated"] = True
        elif args.cmd == "access-row":
            result = mb.append_access_row(
                args.docx,
                system=args.system,
                purpose=args.purpose,
                username=args.username,
                password=args.password,
                notes=args.notes,
            )
        elif args.cmd == "hmi-stage":
            body = args.body_file.read_text(encoding="utf-8")
            result = mb.stage_hmi_screen(
                args.state_json,
                title=args.title,
                body=body,
                image_path=args.image,
                placeholder=args.placeholder,
            )
        elif args.cmd == "hmi-finalize":
            result = mb.finalize_hmi_from_state(
                args.state_json,
                draft_path=args.draft,
                placeholder=args.placeholder,
                width_inches=args.width,
            )
        elif args.cmd == "hmi-staging":
            result = mb.get_hmi_staging(args.state_json, placeholder=args.placeholder)
        elif args.cmd == "stage-section":
            body = args.body_file.read_text(encoding="utf-8")
            result = mb.stage_section(args.state_json, args.placeholder, body)
        elif args.cmd == "revise":
            body = args.body_file.read_text(encoding="utf-8")
            result = mb.revise_section(
                args.state_json,
                args.placeholder,
                body,
                draft_path=args.draft,
            )
        elif args.cmd == "list-staging":
            result = mb.list_staging(args.state_json)
        elif args.cmd == "build-draft":
            result = mb.build_draft_from_staging(
                args.state_json,
                draft_path=args.draft,
                include_hmi=not args.skip_hmi,
            )
        elif args.cmd == "repair-lists":
            result = mb.repair_embedded_lists(args.docx)
        elif args.cmd == "add-section":
            state = mb.load_state(args.state_json)
            draft = Path(args.draft or state["draft_path"])
            result = mb.add_section_to_draft(
                draft,
                args.state_json,
                section_number=args.number,
                title=args.title,
                placeholder=args.placeholder,
                after_placeholder=args.after,
            )
        elif args.cmd == "render-pdf":
            inputs = mm.resolve_project_inputs_dir(
                draft_path=args.draft,
                project_root=args.project_root,
                project_id=args.project_id,
                inputs_dir=args.inputs_dir,
            )
            result = mm.render_pdf_page(
                args.pdf,
                args.page,
                inputs,
                dpi=args.dpi,
                name=args.name,
                clip_norm=args.clip_norm,
                clip_pt=args.clip_pt,
            )
        else:
            raise ValueError(f"Unknown command: {args.cmd}")
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": str(exc)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
