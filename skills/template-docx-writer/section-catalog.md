# Controls manual — section catalog (default seed)

Copy to `templates/section-catalog.md` in the Project Manuals workspace.

**Manual type:** **Operator manual** — for the person **running** the machine (startup, operation, shutdown, safety, daily checks). It is **not** a controls-engineering, programming, or commissioning guide.

**Template:** `{SNN_SECTION}` per chapter. **Tables:** pipe markdown only — script applies double-border styling to every table.

**§12:** Optional `{S12_SECTION}`; stage `# 12.0` + pipe table (System | Purpose | Username | Password | Notes). No `{SYLO_TABLE_ACCESS}` placeholder.

## Audience and voice (required)

Write every staged section for **machine operators** and **operator maintenance** (daily checks, tip change, housekeeping). The reader **does not** have Studio 5000, ladder logic, or the PLC program file.

| Do | Do not |
|----|--------|
| HMI screens, panel labels, pushbutton names, selector positions, E-Stops, guards, weld start/stop | Programming, I/O configuration, network design, “download the project”, tag databases |
| “Press **WELD START**”, “select **Manual** on the HMI”, “rotate the **ROTATION** switch” | “In the L5X…”, “the export file shows…”, “rung logic”, “add-on instruction” |
| What the operator sees, hears, and does at the cell | IP addresses, slot numbers, catalog numbers, and architecture **unless** the operator must use them on the HMI/password table (§12) |

**Tone:** Clear, imperative, shop-floor. Explain **what to do** and **what happens** on the machine — not how the PLC was programmed.

## Sources for the agent (not for the manual text)

Use project files to **understand** behavior; translate into operator language in prose.

| Source | Agent use | In staged manual? |
|--------|-----------|-------------------|
| **L5X / PLC export** | Situation awareness: modes, interlocks, devices, sequence logic | **Never** — operators do not receive this file; **do not name** L5X, RSLogix, Logix Designer, or “the program export” in any section |
| **Electrical drawing (PDF)** | Device labels on panels/pendants (match nameplate/HMI text) | Yes — refer as **electrical drawing** or **panel label**, not sheet-internal drafting jargon |
| **HMI** (screenshots if available) | Menus, faults, mode names the operator actually sees | Yes |
| **Layout / panel photos** | Physical control identification (§9) | Yes |
| **Mechanical / weld docs** | Capacities and limits when confirmed | Yes, in operator terms |

When §11 (control hardware) or networking appears in the template, limit to what **operations or password handoff** need (e.g. HMI login, “call maintenance”) — not a network engineering chapter.

## Document styles

Inject loads the package **`manual-styles.json`** (controls-manual defaults), merges **`templates/manual-styles.json`** in your workspace if present (created on first `manual_start_project`), then optional overrides from the table below. Re-inject sections after changing styles.

| Key | Value |
|-----|-------|
| heading_color | `#0E2841` |
| heading_levels | `1, 2, 3` |
| heading_apply_when_not_default | `true` |
| blank_line_before_heading1 | `false` |
| blank_line_before_heading2 | `true` |
| heading1_page_break_before | `true` |
| trim_leading_trailing_blank_lines | `true` |

**Heading color:** Word theme *Dark Blue, Text 2, Lighter 10%* ≈ `#0E2841`. On inject, levels 1–3 get this color when runs are still black/automatic; custom colors are left as-is.

**Layout:** `blank_line_before_heading1` / `heading2` = `true` if you want a blank Normal line before those headings (default: `false`). `heading1_page_break_before` = `true` starts each injected chapter H1 on a new page (replaces template placeholder page breaks removed during inject).

| Key | Value |
|-----|-------|
| image_caption_enabled | `true` |
| image_caption_font | `Times New Roman` |
| image_caption_size_pt | `10` |
| image_caption_italic | `true` |
| image_caption_prefix | `Figure` |
| image_caption_use_seq | `true` |
| image_caption_word_style | `Caption` |
| image_caption_align | `center` |
| image_caption_blank_line_after | `true` |
| image_align | `center` |

**Figure index:** Headings do **not** appear in Word’s figure list. Use markdown images with **alt text = caption description**; inject adds a **Caption** paragraph (`Figure` + SEQ + description). Update with **References → Table of Figures** (or F9 if your template links it).
