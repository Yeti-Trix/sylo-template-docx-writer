---
name: template-docx-writer
description: Fill any standard Word .docx template from a template bundle — inject AI-authored sections, images, and tables. Routes operator requests to the bundle's manual-styles.json, section-catalog.md, state.json staging, and manual_* tools (inject, revise, tables, images). Controls manuals are one template bundle.
metadata:
  sylo:
    category: documents
    icon: file-text
---

# Template docx writer workflow

> **Bundle model:** templates live as bundles under `sylo-user/docx-templates/<id>/` (single source of truth). See **`template-bundle.spec.md`** (this folder) for the file contract and the bundled workflow **`build-docx-from-template`** for the process. The routing tables below still apply to a controls-manuals bundle. `manual_start_project` accepts a bundle folder or a `.docx`; passing a bundle records `template_bundle_path` in `state.json` and reads styles/catalog from there.

## Operator says … → you use …

When the operator asks to change the manual in plain language, **do not guess**. Match intent to the right **workspace file** and **`manual_*` tool**, then **re-inject** if inject styling or staged prose changed.

| Operator intent (examples) | Edit / read | Then |
|----------------------------|-------------|------|
| Write or rewrite §4 startup, fix wording in a chapter | **`projects/<id>/state.json`** staging + **`templates/section-catalog.md`** (what belongs in that slug) | **One slug at a time** → `manual_stage_section` → next slug (see **Pacing**); `manual_revise_section` or `manual_build_draft` when ready |
| Change **heading color**, **caption font**, blank line before H1/H2, figure prefix | **`templates/manual-styles.json`** (preferred) or catalog **Document styles** table for one-off overrides | Edit JSON → **`manual_revise_section`** on affected slugs (or full `manual_build_draft`) — inject applies styles; editing `draft.docx` by hand does not stick on re-inject |
| Change **what** each chapter should cover, add checklist for a slug | **`templates/section-catalog.md`** only | No inject until prose is re-staged |
| Cover title / job number / manual type | **`draft.docx`** placeholders via tool | `manual_fill_cover` |
| Add layout photo, diagram, HMI screenshot | **`projects/<id>/inputs/`** + staging `![caption](inputs/...)` — or **`manual_render_pdf_page`** / `manual_insert_image` / `manual_stage_hmi_screen`; source `.docx` pictures via **docx-reader** `extract_docx_images` (`output_dir` = `inputs/`) | PDF: full page or **crop** (`clip_norm` 0–1); caption = alt text |
| Password / access table rows | Staged pipe table in §12 **or** `access-table-example-rows.json` + tools | `manual_replace_access_table` / `manual_append_access_row` — table **layout** rules in **`templates/table-formats.json`** |
| New project, first manual on this workspace | Template `.docx` path | `manual_start_project` (seeds `templates/` including default **`manual-styles.json`**) |
| List what’s left, resume job | **`state.json`** | `manual_project_state` / `manual_list_staging` |
| Template structure, new `{SNN_SECTION}` in Word | **`controls-template.docx`** | `manual_discover_placeholders`; rare new chapter → operator OK → `manual_add_section` |

**Not in prose:** Heading color, caption Times 10 italic, table borders, list behavior — those come from **`manual-styles.json`** at inject time.

**Read this skill** for workflow; read **`section-catalog.md`** per slug before drafting; read **`manual-styles.json`** before changing global look-and-feel.

## Workspace files (under Project Manuals or Pi cwd)

| File | Purpose |
|------|---------|
| **`templates/manual-styles.json`** | Inject styling (headings, captions, layout) — **main knob** for “make the document look like …” |
| **`templates/section-catalog.md`** | What to write per `{SNN_SECTION}`; optional **Document styles** overrides |
| **`templates/table-formats.json`** | How to find/fill access & revision tables |
| **`templates/access-table-example-rows.json`** | Example rows for access table tool (optional) |
| **`projects/<id>/state.json`** | Staged markdown per section, pending/completed checklist |
| **`projects/<id>/draft.docx`** | Built manual — change via **`manual_*` tools**, not direct edit for staged content |
| **`projects/<id>/inputs/`** | Images referenced from staging |

## Skill vs section catalog

| | **Skill (this file)** | **`templates/section-catalog.md`** |
|---|----------------------|--------------------------------------|
| Purpose | **How** to run the job and **which file/tool** for each request | **What** each slug means (title, sources, §6 HMI vs prose) |
| Slugs in Word | Generic `{SNN_SECTION}` | Maps `S06_SECTION` → operation chapter, etc. |
| Document styles | Points to **`manual-styles.json`** | Optional **Document styles** table overrides |

**New workspace:** First **`manual_start_project`** copies package **`manual-styles.json`** into `templates/` (controls-manual defaults) if missing. Inject always uses package + workspace file + catalog overrides.

**Subsections:** Catalog checklists (3.1, 4.2, etc.) are **patterns**, not mandatory outlines.

**Catalog is generic:** Do not paste catalog patterns or examples into prose — only project sources. Omit any device/subsystem not confirmed for **this** job.

**Operator manual:** Staged prose is for **people running the machine**, not controls engineers. Use **`section-catalog.md` → Audience and voice** — never mention **L5X** or PLC program files in the manual; L5X is for **your** research only.

**Extra top-level chapters:** The template lists the **usual** chapters only. **Do not add** new §12-style chapters unless the operator agrees — most manuals need **no** extra H1s. When truly needed: operator OK → **`manual_add_section`** (placeholder only) → stage with `# 12.0 Title` as the first line → build. Prefer editing within existing `{SNN_SECTION}` slugs.

Pair with **pdf-reader** for electrical PDFs.

## Placeholder naming (Word template)

| Pattern | Use |
|---------|-----|
| `{S00_COVER_PROJECT_TITLE}`, `{S00_COVER_JOB_NUMBER}`, `{S00_COVER_MANUAL_TYPE}` | Cover — **`manual_fill_cover`** (keeps template font/size) |
| `{S01_SECTION}` … `{S20_SECTION}` | **One line per chapter** in the template — **no** static Heading 1 text; body (including H1) comes from staging |
| *(none)* | Static boilerplate (warnings, CAUTION blocks) |

**All caps, no topic in the slug:** `{S06_SECTION}` not `{S06_OPERATION}`.

## Template & catalog authoring

1. **Remove** static Heading 1 lines (`2.0 Introduction`, etc.) from `controls-template.docx`.
2. Leave **only** `{SNN_SECTION}` (one paragraph per chapter), in document order.
3. `manual_discover_placeholders` on the template.
4. Update **section-catalog.md** per slug.
5. `prepare_controls_template.py` (optional) or hand-edit, then `manual_start_project`.

## Pacing (always) vs approval (when the operator wants it)

**Pacing is mandatory.** Work **one section at a time** in catalog order (Terms last). For each slug:

1. Read catalog + gather only the sources needed for **that** slug  
2. Draft prose in chat (optional)  
3. Call **`manual_stage_section`** once for that slug — **wait for success** before the next slug  
4. Move to the next slug  

**Never** stage the whole manual in one burst: no parallel `manual_stage_section` calls, no multi-slug tool batch in a single step. `state.json` is not safe under concurrent writes.

**Approval is optional.** Default: show the draft and **wait for operator OK** before staging. If the operator says **auto-approve**, **build everything**, **test the workflow**, or similar, **keep the same one-at-a-time pacing** but **do not stop for OK after each section** — stage each slug sequentially, then `manual_build_draft` when all slugs are done (or when they ask to build).

| Operator says | You do |
|---------------|--------|
| Normal / production | Draft section → **wait for OK** → `manual_stage_section` → next slug |
| Auto-approve / test workflow / no stops | Same **one slug per `manual_stage_section`**, no OK between slugs → build when complete |

Light shared research up front (e.g. `manual_start_project`, copy inputs, read L5X/PDF for **agent** context) is fine — **do not put L5X references in staged text**; **staging** must still be **serial, one placeholder per tool call**.

## Core protocol

```
For each row in section-catalog (Terms last):
  1. Read catalog — what to write, sources, static skip (this slug only)
  2. Research if needed → draft in chat
  3. manual_stage_section for THIS slug only — then next slug
     (operator OK before step 3 unless they asked for auto-approve)
  4. Repeat until all slugs staged

When all staged (or operator asks) → manual_build_draft → Word F9 → PDF
```

**§6 (`S06_SECTION`):** Ask if operator HMI exists.

- **HMI:** `manual_stage_hmi_screen` per screen → `manual_build_draft` / `manual_finalize_hmi`
- **No HMI:** one `manual_stage_section` with slug `S06_SECTION`

Do not mix both paths on one project.

## Images in any section

The operator may add photos or diagrams to **any** section — not only §6 HMI.

**Without chat paste:** save assets under `inputs/` via tools, then reference in staging:

| Tool | Use |
|------|-----|
| `manual_render_pdf_page` | PDF page → PNG in `inputs/` (`draft_path` required). Optional **`clip_norm`**: `left,top,right,bottom` as **0–1** fractions of the page (top-left origin) to grab one diagram, not the whole sheet. Optional **`clip_pt`**: x0,y0,x1,y1 in PDF points after you measure from a full-page render. |
| `extract_docx_images` (**docx-reader** package) | Embedded pictures from a source `.docx` (old manual, vendor doc) — set `output_dir` to the project `inputs/`. Returns per-image `caption_guess` + `context_text` for alt text. Enable **sylo-docx-reader** if missing. |
| `manual_insert_image` | Rare: `{IMG_...}` placeholder in template instead of markdown image line |

After render/extract, use returned **`markdown_snippet`** in `manual_stage_section` (edit caption in `![...]` alt text).

**Figure list / image index:** Word needs a **caption**, not a heading. On every inject, the script adds a **Caption** line under the image (`Figure` + auto number + your description). Staging uses markdown **alt text** as the description:

```markdown
![Main enclosure layout — door E-Stop shown](inputs/s09-layout.png)
```

**§9 Physical Controls (`S09_SECTION`):** Operator supplies a **layout image** under `inputs/`. Image line **first**, then device text.

**Rule: image (with caption alt text) first, then the text that describes it.**

Caption font (Times New Roman 10 pt italic by default) is set in **`templates/manual-styles.json`** / catalog **Document styles** — not in prose.

## Tools

| Tool | When |
|------|------|
| `manual_start_project` | New job |
| `manual_discover_placeholders` | After template edits |
| `manual_fill_cover` | Cover placeholders |
| `manual_project_state` / `manual_list_staging` | Resume |
| `manual_stage_section` | **One** prose slug per call — never parallel with other stage calls |
| `manual_revise_section` | Edit a section already built |
| `manual_add_section` | **Rare** — new chapter placeholder (operator OK only) |
| `manual_stage_hmi_screen` | **One** §6 HMI screen per call (HMI path) |
| `manual_build_draft` | After all slugs staged (or operator requests build) |
| `manual_finalize_hmi` | §6 HMI only |
| `manual_render_pdf_page` | PDF page (optional crop) → `inputs/` |
| `extract_docx_images` (**docx-reader**) | Pictures from source `.docx` → `inputs/` (set `output_dir`) |
| `manual_replace_access_table` | Fill **`access`** marked table from JSON (if catalog §12 uses it) |
| `manual_append_access_row` | One row on **`access`** table |

## Formatted prose (markdown at staging)

**Every `{SNN_SECTION}` body must start with Heading 1:**

```markdown
# 2.0 Introduction

Opening paragraph…

## 2.1 Scope

### 8.3.1 E-Stop Guidelines

Emergency-stop devices are for emergency use only…

**Short H3 title (≤ ~7 words):** one Heading 3 line — number + title (shows in TOC).

**Long step:** short title on H3, full step in Normal below — preferred:

```markdown
### 10.4.2 Confirm Guards

Confirm all guards are secure and the boom/column motion envelope is clear…
```

Same-line optional: `### 10.4.2 Confirm Guards — Confirm all guards are secure…`
```

- **`# 2.0 …` / `2.0 …` on line 1** → **Heading 1** (chapter title — only once per slug; do **not** repeat as `## 2.0`).
- **`## 3.1 …`** → **Heading 2**
- **`### 4.3.1 …`** → **Heading 3** (number only) + **Normal** body
- **§4 / §5 procedures:** `## 4.1 Topic` then **`1.` / `2.` / `3.`** action lines — **not** `### 4.1.1` for one-line steps
- **§6 operation** (multi-step skills): use `### 6.2.1` when each step is a real sub-topic with prose; do not use `1.` lists for whole subsections there
- `**bold**` → Word bold
- **Footnote line:** start with `\* ` (backslash-asterisk) so inject prints a literal `*` in Normal text — e.g. `\* HMI IP is plant-assigned; confirm at startup.` (do not use a bare `* ` line unless you want a bullet)
- **Pipe tables** (`| col |` + `|---|`) → double borders, bold header, Times 12; Username/Password columns centered (§12 same as §11)
- **`-` bullets** → unordered lists
- **`1.` lines** → ordered steps (Normal text `1. …`, `2. …` — restarts at each `##` topic; stays **not** Heading 3)
- Terms = bullets only

**Re-inject:** `manual_build_draft` re-applies staging; fix `state.json` then **`manual_revise_section`**.

## Editing an already-built section

`manual_revise_section` after operator OK — do not re-run `manual_start_project`.

## Do not bypass tools

Use only **`manual_*` tools** for `draft.docx`.

## Rules

- **Always** stage **one `{SNN_SECTION}` per `manual_stage_section` call** — sequential only, no parallel staging.
- **Default:** wait for operator OK before each `manual_stage_section`; **auto-approve mode:** same pacing, no OK between sections.
- **Default:** wait for operator confirmation before `manual_build_draft`; in auto-approve mode, build when all catalog slugs are staged.
- `S01_SECTION` (Terms) last unless catalog says otherwise.
- **Default: no new top-level sections** — use catalog slugs unless operator requests `manual_add_section`.
