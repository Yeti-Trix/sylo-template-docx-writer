# sylo-template-docx-writer

Fill any standard Word `.docx` template from a **template bundle** — inject AI-authored sections, images, and tables. Generalizes the former `sylo-manual-creator`; controls manuals are one template bundle. Other bundles can describe HMI manuals, robot manuals, or any docx that follows a fixed placeholder standard.

## Enable

**Capability manager → Sylo optional packages → Template docx writer → On** (installs python-docx) → **Restart broker** → `npm run bootstrap-pi`

## The three layers

| Layer | What | Where |
|-------|------|-------|
| **Tools** (`manual_*`) | Generic inject/stage/build/render/image/table tools — operate on placeholders + a draft.docx | This package (`extensions/index.ts` + `scripts/`) |
| **Process** | The shared workflow — pacing, placeholder→tool mapping, build order | Bundled workflow `build-docx-from-template.md` (in `sylo-workflows/shared/workflows/`) |
| **Bundle** | The template-specific data: the `.docx` skeleton, styles, section catalog, manifest | Operator-owned, git-backed: `sylo-user/docx-templates/<id>/` |

A newcomer only needs to read **`template-bundle.spec.md`** (in this skill folder) to understand what files a bundle requires, then provide a bundle (or ask the agent to help build one). The agent loads the bundled workflow on demand.

## Template bundles (single source of truth)

A bundle is a folder under the operator's git-backed `sylo-user/docx-templates/<id>/`:

```
<id>/
  controls-template.docx     # Word skeleton with {SNN_SECTION} / {S00_COVER_*} / {IMG_*} placeholders
  manifest.json              # template_id, required_sections, optional_sections, write_last, placeholder_format
  manual-styles.json         # this template's look (heading color, caption font, table borders)
  section-catalog.md         # what to write per slug + optional Document styles overrides
  table-formats.json         # optional — access/revision table markers
  access-table-defaults.json # optional
```

See **`skills/template-docx-writer/template-bundle.spec.md`** for the full file contract and manifest schema.

**Single source of truth:** `manual_start_project` points at a bundle folder (or the `.docx` inside one). The bundle's `manual-styles.json` + `section-catalog.md` are read at inject time from the path recorded in `state.json` (`template_bundle_path`). Edit the bundle once → every future project built from it picks up the change. Per-project one-off tweaks are made directly in that project's `draft.docx`, not the bundle.

**Legacy bare `.docx`:** passing a `.docx` with no sibling `manifest.json` preserves the old behavior (workspace `templates/` seeded from package defaults). Existing projects are unaffected.

## Tools (names kept stable from manual-creator)

- `manual_start_project` — point at a bundle folder or `.docx`; records `template_bundle_path`; copies the template → `projects/<id>/draft.docx` + writes `state.json`
- `manual_discover_placeholders` — list `{TAGS}` in a docx
- `manual_render_pdf_page` — PDF page → `projects/<id>/inputs/` (optional crop)
- `manual_stage_section` / `manual_revise_section` — stage or re-stage one slug (one per call, serial)
- `manual_build_draft` — inject all staged sections
- `manual_stage_hmi_screen` / `manual_finalize_hmi` — §6 HMI path
- `manual_insert_image` / `manual_append_revision` / `manual_replace_access_table` / `manual_append_access_row`
- `manual_project_state` / `manual_list_staging` — resume / checklist

Use with **pdf-reader** for PDF sources. Pulling pictures from a source `.docx` is in **sylo-docx** (`extract_docx_images` with `output_dir` = project `inputs/`).

**Agent guidance:** Pi skill **`template-docx-writer`** (`skills/template-docx-writer/SKILL.md`) maps operator requests ("change heading color", "fix §4", "add figure") to the bundle's `manual-styles.json`, `section-catalog.md`, `state.json`, and `manual_*` tools.


## Install

`pi install npm:sylo-template-docx-writer` — or from the **Capability manager → Pi.dev package catalog** in Sylo (it appears in the Sylo packages strip).

Releases publish automatically from GitHub Actions (npm trusted publishing, with provenance): bump `version` in `package.json`, commit, tag `vX.Y.Z`, push the tag.
