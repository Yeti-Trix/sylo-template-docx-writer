# Template bundle specification

A **template bundle** is a folder that describes one kind of docx document this engine can build — e.g. a controls manual, an HMI manual, a robot manual, or any docx that follows a fixed placeholder standard. Bundles live in the operator's git-backed library, by default:

```
C:\Users\<user>\Documents\GitHub\sylo-user\docx-templates\<id>\
```

`sylo-user` is synced across machines and shareable. One engine (`sylo-template-docx-writer`), many bundles.

## What a newcomer needs to know

If you want to build a docx from a template and no bundle exists yet, you provide **two required files** (the rest are optional). The agent can help you author them:

1. **`controls-template.docx`** — a Word document whose body is mostly literal `{PLACEHOLDER}` paragraphs (one per chapter, in document order) plus any static boilerplate you want in every document (CAUTION blocks, warnings, cover layout). The chapter text itself does **not** live here — it comes from staging at build time.
2. **`manifest.json`** — the placeholder schema + build order (see schema below).

Then point `manual_start_project` at the bundle folder (or the `.docx` inside it).

## Required files

| File | Role |
|------|------|
| `controls-template.docx` | Word skeleton with `{SNN_SECTION}` / `{S00_COVER_*}` / `{IMG_*}` placeholders and static boilerplate |
| `manifest.json` | Placeholder schema + build order |

## Optional files

| File | Role | Default if missing |
|------|------|--------------------|
| `manual-styles.json` | Inject styling (heading color, caption font, table borders, blank-line layout) | Package defaults (see `skills/template-docx-writer/manual-styles.json`) |
| `section-catalog.md` | What to write per slug + optional **Document styles** overrides | No content guidance (agent writes from operator input) |
| `table-formats.json` | How to find/fill structured tables (access, revision) | None |
| `access-table-defaults.json` | Seed rows for the access table | None |
| `access-table-example-rows.json` | Example rows shown to the operator | None |

## `manifest.json` schema

```json
{
  "template_id": "controls-manuals",
  "description": "Controls manual — generic {SNN_SECTION} placeholders",
  "placeholder_format": "{SNN_SECTION}",
  "required_sections": [
    "S00_COVER_PROJECT_TITLE",
    "S00_COVER_JOB_NUMBER",
    "S01_SECTION",
    "S02_SECTION",
    "S06_SECTION"
  ],
  "optional_sections": ["S12_SECTION"],
  "write_last": ["S01_SECTION"],
  "notes": "Optional free text for the agent/operator."
}
```

| Key | Meaning |
|-----|---------|
| `template_id` | Short slug; matches the bundle folder name |
| `description` | Human description shown when listing templates |
| `placeholder_format` | The placeholder syntax used (default `{SNN_SECTION}`) |
| `required_sections` | Placeholders the project must fill, in document order |
| `optional_sections` | Placeholders that may be present but are not required |
| `write_last` | Slugs staged only after all others (e.g. Terms `S01_SECTION` last) |
| `notes` | Free text |

## Placeholder conventions

| Pattern | Filled by |
|---------|-----------|
| `{S00_COVER_PROJECT_TITLE}`, `{S00_COVER_JOB_NUMBER}`, `{S00_COVER_MANUAL_TYPE}` | `manual_fill_cover` (preserves template font/size on those lines) |
| `{SNN_SECTION}` (one line per chapter, no topic in the slug) | `manual_stage_section` → `manual_build_draft` |
| `{IMG_*}` | `manual_insert_image` |
| Marked table (`table-formats.json` id: `access`) | `manual_replace_access_table` / `manual_append_access_row` |

All caps, no topic in the slug: `{S06_SECTION}`, not `{S06_OPERATION}`. Remove static Heading 1 lines (`2.0 Introduction`, etc.) from `controls-template.docx` — the H1 comes from staged prose, not the template.

## The three layers (why files split this way)

| Layer | What | Constant across bundles? |
|-------|------|--------------------------|
| **Tools** (`manual_*`) | Generic inject/stage/build/render/image/table mechanics | Yes — package code |
| **Process** (workflow `build-docx-from-template.md`) | Pacing, placeholder→tool mapping, build order, approval gates | Yes — one bundled workflow |
| **Bundle** (this folder) | The `.docx` skeleton, styles, section catalog, manifest | No — one per document type |

Keyword/placeholder→tool mapping lives in the **workflow** (process), not the catalog — because it is constant. The catalog holds **content guidance** (what to write per slug), which differs per bundle. The manifest holds **which placeholders + order**, which differs per bundle.

## Creating a new bundle

1. Author `controls-template.docx` — `{SNN_SECTION}` paragraphs in order, cover placeholders, static boilerplate only.
2. Write `manifest.json` (schema above). Run `manual_discover_placeholders` on the `.docx` to confirm the placeholder inventory matches `required_sections`.
3. Optionally add `manual-styles.json` (copy the package default and edit), `section-catalog.md`, `table-formats.json`.
4. Place the folder under `sylo-user/docx-templates/<id>/`.
5. Point `manual_start_project` at the bundle folder.

The agent can walk a newcomer through any of these steps — describe the document type you want and let it draft the template + manifest.