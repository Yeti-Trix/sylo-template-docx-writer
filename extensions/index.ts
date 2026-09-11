import { execFile } from 'node:child_process'
import { writeFileSync, unlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import path from 'node:path'

import type { ExtensionAPI } from '@earendil-works/pi-coding-agent'
import { Type } from 'typebox'

const execFileAsync = promisify(execFile)

const PACKAGE_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const CLI = path.join(PACKAGE_ROOT, 'scripts', 'manual_cli.py')

type ToolContentBlock = { type: 'text'; text: string }

function resolvePython(): string {
  return process.platform === 'win32' ? 'python' : 'python3'
}

async function runManualCli(
  args: string[],
): Promise<{ ok: true; data: unknown } | { ok: false; error: string }> {
  try {
    const { stdout, stderr } = await execFileAsync(resolvePython(), [CLI, ...args], {
      cwd: PACKAGE_ROOT,
      maxBuffer: 8 * 1024 * 1024,
      windowsHide: true,
    })
    const trimmed = stdout.trim()
    if (!trimmed) {
      return { ok: false, error: stderr.trim() || 'manual_cli produced no output' }
    }
    const data = JSON.parse(trimmed) as unknown
    const rec = data as Record<string, unknown>
    if (rec.error) return { ok: false, error: String(rec.error) }
    return { ok: true, data }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    return {
      ok: false,
      error:
        `${message}\n` +
        `Ensure Python deps: pip install -r packages/sylo-template-docx-writer/scripts/requirements.txt`,
    }
  }
}

function toolError(text: string): { content: ToolContentBlock[] } {
  return { content: [{ type: 'text', text }] }
}

function ok(data: unknown, summary: string): { content: ToolContentBlock[] } {
  return {
    content: [
      { type: 'text', text: summary },
      { type: 'text', text: JSON.stringify(data, null, 2) },
    ],
  }
}

export default function syloTemplateDocxWriterExtension(pi: ExtensionAPI): void {
  pi.registerTool({
    name: 'manual_discover_placeholders',
    label: 'Discover manual placeholders',
    description:
      'List {PLACEHOLDER} tags in a Word template or draft.docx. Run before writing sections so the AI knows what the template expects.',
    parameters: Type.Object({
      docx_path: Type.String({ description: 'Path to template or draft .docx' }),
    }),
    async execute(_id, params) {
      const docx = String(params.docx_path ?? '').trim()
      if (!docx) return toolError('docx_path required.')
      const r = await runManualCli(['discover', docx])
      if (!r.ok) return toolError(r.error)
      const data = r.data as { placeholders?: string[] }
      return ok(
        r.data,
        `Found ${data.placeholders?.length ?? 0} placeholder(s). Draft prose in conversation, then inject after operator approval.`,
      )
    },
  })

  pi.registerTool({
    name: 'manual_start_project',
    label: 'Start manual project',
    description:
      'Copy a Word template into projects/<id>/draft.docx, fix Strict OOXML if needed, write state.json. Seeds templates/ (section-catalog.md, manual-styles.json with controls-manual heading/caption defaults, table-formats.json) on first use of the workspace.',
    parameters: Type.Object({
      project_root: Type.String({
        description: 'Pi cwd folder (e.g. Project Manuals) containing projects/',
      }),
      project_id: Type.String({ description: 'Slug e.g. 12345-EXAMPLE' }),
      template_path: Type.String({ description: 'Path to controls template .docx' }),
    }),
    async execute(_id, params) {
      const root = String(params.project_root ?? '').trim()
      const id = String(params.project_id ?? '').trim()
      const tpl = String(params.template_path ?? '').trim()
      if (!root || !id || !tpl) return toolError('project_root, project_id, template_path required.')
      const r = await runManualCli(['start', root, id, tpl])
      if (!r.ok) return toolError(r.error)
      const seeded = (r.data as { workspace?: { seeded_files?: string[] } })?.workspace?.seeded_files ?? []
      const stylesNote = seeded.some((p) => p.includes('manual-styles.json'))
        ? ' Seeded templates/manual-styles.json (controls-manual inject defaults).'
        : ''
      return ok(
        r.data,
        `Project ${id} started.${stylesNote} Edit templates/manual-styles.json to customize; re-inject after changes.`,
      )
    },
  })

  pi.registerTool({
    name: 'manual_fill_cover',
    label: 'Fill cover placeholders',
    description:
      'Replace {S00_COVER_PROJECT_TITLE}, {S00_COVER_JOB_NUMBER}, {S00_COVER_MANUAL_TYPE} in draft.docx without changing template font/size on those lines.',
    parameters: Type.Object({
      draft_path: Type.String({ description: 'Path to draft.docx' }),
      project_title: Type.Optional(Type.String()),
      job_number: Type.Optional(Type.String()),
      manual_type: Type.Optional(Type.String({ description: 'e.g. Controls Manual' })),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      if (!draft) return toolError('draft_path required.')
      const args = ['fill-cover', draft]
      const title = String(params.project_title ?? '').trim()
      const job = String(params.job_number ?? '').trim()
      const mtype = String(params.manual_type ?? '').trim()
      if (title) args.push('--project-title', title)
      if (job) args.push('--job-number', job)
      if (mtype) args.push('--manual-type', mtype)
      if (args.length === 2) return toolError('At least one of project_title, job_number, manual_type required.')
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Cover placeholders filled (formatting preserved).')
    },
  })

  pi.registerTool({
    name: 'manual_inject_section',
    label: 'Inject manual section',
    description:
      'Replace {PLACEHOLDER} in draft.docx with approved prose. Call only after operator approves text in chat. Sub-subsections use ### 4.3.1 (Heading 3), not 1.1 under 4.3.',
    parameters: Type.Object({
      draft_path: Type.String({ description: 'Path to projects/.../draft.docx' }),
      placeholder: Type.String({ description: 'e.g. SECTION_3_CONTENT (without braces)' }),
      content: Type.String({
        description: 'Approved markdown; sub-subsections as ### 4.3.1 (full three-level number).',
      }),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      const ph = String(params.placeholder ?? '').trim()
      const content = String(params.content ?? '')
      if (!draft || !ph) return toolError('draft_path and placeholder required.')
      const tmp = join(tmpdir(), `sylo-manual-${randomUUID()}.txt`)
      writeFileSync(tmp, content, 'utf8')
      try {
        const r = await runManualCli(['inject', draft, ph, '--content-file', tmp])
        if (!r.ok) return toolError(r.error)
        return ok(r.data, `Injected {${ph.toUpperCase()}} into draft.`)
      } finally {
        try {
          unlinkSync(tmp)
        } catch {
          /* ignore */
        }
      }
    },
  })

  pi.registerTool({
    name: 'manual_insert_image',
    label: 'Insert manual image',
    description:
      'Replace an {IMG_...} placeholder with a screenshot or diagram. Adds a Caption line (SEQ Figure) for Word table of figures.',
    parameters: Type.Object({
      draft_path: Type.String(),
      placeholder: Type.String({ description: 'e.g. IMG_6_1_DASHBOARD' }),
      image_path: Type.String({ description: 'Path to image file (save pasted screenshots to project inputs/ first)' }),
      width_inches: Type.Optional(Type.Number({ description: 'Default 6.5', minimum: 1, maximum: 8 })),
      caption: Type.Optional(
        Type.String({
          description: 'Caption description after "Figure N:" (defaults to filename stem)',
        }),
      ),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      const ph = String(params.placeholder ?? '').trim()
      const img = String(params.image_path ?? '').trim()
      if (!draft || !ph || !img) return toolError('draft_path, placeholder, image_path required.')
      const args = ['image', draft, ph, img]
      if (typeof params.width_inches === 'number') args.push('--width', String(params.width_inches))
      const cap = String(params.caption ?? '').trim()
      if (cap) args.push('--caption', cap)
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      return ok(r.data, `Image inserted at {${ph.toUpperCase()}}.`)
    },
  })

  pi.registerTool({
    name: 'manual_render_pdf_page',
    label: 'Render PDF page to project inputs',
    description:
      'Render one PDF page to PNG under projects/<id>/inputs/ for staging (![caption](inputs/...)). ' +
      'Optional crop: clip_norm (0–1 fractions, top-left) or clip_pt (PDF points). Requires PyMuPDF (pip with this package requirements).',
    parameters: Type.Object({
      pdf_path: Type.String({ description: 'Path to source PDF' }),
      page: Type.Number({ description: '1-based page number', minimum: 1 }),
      draft_path: Type.Optional(
        Type.String({ description: 'projects/<id>/draft.docx — inputs/ is sibling folder' }),
      ),
      project_root: Type.Optional(
        Type.String({ description: 'Workspace root (parent of projects/) if no draft_path' }),
      ),
      project_id: Type.Optional(Type.String()),
      dpi: Type.Optional(Type.Number({ description: 'Default 150', minimum: 72, maximum: 600 })),
      name: Type.Optional(Type.String({ description: 'Output filename stem' })),
      clip_norm: Type.Optional(
        Type.String({
          description:
            'Crop region: left,top,right,bottom as 0–1 fractions of page size (e.g. 0.1,0.2,0.9,0.85)',
        }),
      ),
      clip_pt: Type.Optional(
        Type.String({
          description: 'Crop region: x0,y0,x1,y1 in PDF points (top-left origin); use after measuring one full-page render',
        }),
      ),
    }),
    async execute(_id, params) {
      const pdf = String(params.pdf_path ?? '').trim()
      const page = Math.floor(Number(params.page))
      if (!pdf || !Number.isFinite(page) || page < 1) {
        return toolError('pdf_path and page (>= 1) required.')
      }
      const draft = String(params.draft_path ?? '').trim()
      const root = String(params.project_root ?? '').trim()
      const pid = String(params.project_id ?? '').trim()
      if (!draft && !(root && pid)) {
        return toolError('Provide draft_path or project_root + project_id.')
      }
      const args = ['render-pdf', pdf, '--page', String(page)]
      if (draft) args.push('--draft', draft)
      if (root) args.push('--project-root', root)
      if (pid) args.push('--project-id', pid)
      if (typeof params.dpi === 'number') args.push('--dpi', String(params.dpi))
      const nm = String(params.name ?? '').trim()
      if (nm) args.push('--name', nm)
      const cn = String(params.clip_norm ?? '').trim()
      if (cn) args.push('--clip-norm', cn)
      const cp = String(params.clip_pt ?? '').trim()
      if (cp) args.push('--clip-pt', cp)
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      const data = r.data as { relative_markdown?: string }
      return ok(
        r.data,
        `Saved ${data.relative_markdown ?? 'PNG'} — use markdown_snippet in manual_stage_section.`,
      )
    },
  })

  pi.registerTool({
    name: 'manual_append_revision',
    label: 'Append revision row',
    description: 'Add a row to the first table in the document (revision history).',
    parameters: Type.Object({
      draft_path: Type.String(),
      rev: Type.String(),
      date: Type.String(),
      description: Type.String(),
      author: Type.String({ description: 'Initials or name' }),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      if (!draft) return toolError('draft_path required.')
      const r = await runManualCli([
        'revision',
        draft,
        '--rev',
        String(params.rev ?? ''),
        '--date',
        String(params.date ?? ''),
        '--description',
        String(params.description ?? ''),
        '--by',
        String(params.author ?? ''),
      ])
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Revision row appended.')
    },
  })

  pi.registerTool({
    name: 'manual_stage_section',
    label: 'Stage manual section',
    description:
      'Store one operator-approved prose section in state.json staging. Does not touch draft.docx. Body must start with # N.0 Chapter title (Heading 1). Do not add extra top-level sections without operator OK + manual_add_section.',
    parameters: Type.Object({
      state_path: Type.String(),
      placeholder: Type.String({ description: 'Catalog slug e.g. S03_SECTION' }),
      content: Type.String({
        description:
          'Markdown: first line # 3.0 Machine Capacities (H1), then ##/### subsections. Terms: bullets only. No duplicate ## N.0 under same slug.',
      }),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      const ph = String(params.placeholder ?? '').trim()
      const content = String(params.content ?? '')
      if (!sp || !ph || !content.trim()) return toolError('state_path, placeholder, content required.')
      const tmp = join(tmpdir(), `sylo-manual-${randomUUID()}.txt`)
      writeFileSync(tmp, content, 'utf8')
      try {
        const r = await runManualCli(['stage-section', sp, ph, '--body-file', tmp])
        if (!r.ok) return toolError(r.error)
        return ok(r.data, `Staged {${ph.toUpperCase()}}. Move to next catalog section or build draft when ready.`)
      } finally {
        try {
          unlinkSync(tmp)
        } catch {
          /* ignore */
        }
      }
    },
  })

  pi.registerTool({
    name: 'manual_revise_section',
    label: 'Revise built section',
    description:
      'Edit a section already written into draft.docx — re-stage new prose and re-inject it in place (replaces old content between hidden section markers). Use this to fix or rewrite a section without manual_start_project. Requires the section to have been built at least once. Call only after operator approves the new text.',
    parameters: Type.Object({
      state_path: Type.String({ description: 'projects/<id>/state.json' }),
      placeholder: Type.String({ description: 'Catalog slug e.g. S09_SECTION' }),
      content: Type.String({ description: 'Approved replacement section body (markdown)' }),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      const ph = String(params.placeholder ?? '').trim()
      const content = String(params.content ?? '')
      if (!sp || !ph || !content.trim()) return toolError('state_path, placeholder, content required.')
      const tmp = join(tmpdir(), `sylo-manual-${randomUUID()}.txt`)
      writeFileSync(tmp, content, 'utf8')
      try {
        const r = await runManualCli(['revise', sp, ph, '--body-file', tmp])
        if (!r.ok) return toolError(r.error)
        return ok(r.data, `Revised {${ph.toUpperCase()}} in draft.docx.`)
      } finally {
        try {
          unlinkSync(tmp)
        } catch {
          /* ignore */
        }
      }
    },
  })

  pi.registerTool({
    name: 'manual_add_section',
    label: 'Add manual section',
    description:
      'Rare: insert {SNN_SECTION} placeholder only (no H1 in Word). Operator must OK first — most manuals need zero extra chapters. After add, stage body starting with # N.0 Title. For every future project, add {SNN_SECTION} to controls-template.docx instead.',
    parameters: Type.Object({
      state_path: Type.String(),
      section_number: Type.String({ description: 'e.g. 12.0' }),
      title: Type.String({ description: 'e.g. Cybersecurity' }),
      placeholder: Type.String({ description: 'e.g. S12_SECTION' }),
      after_placeholder: Type.String({
        description: 'Existing slug to insert after, e.g. S11_SECTION',
      }),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      const number = String(params.section_number ?? '').trim()
      const title = String(params.title ?? '').trim()
      const ph = String(params.placeholder ?? '').trim()
      const after = String(params.after_placeholder ?? '').trim()
      if (!sp || !number || !title || !ph || !after) {
        return toolError('state_path, section_number, title, placeholder, after_placeholder required.')
      }
      const r = await runManualCli([
        'add-section',
        sp,
        '--number',
        number,
        '--title',
        title,
        '--placeholder',
        ph,
        '--after',
        after,
      ])
      if (!r.ok) return toolError(r.error)
      return ok(
        r.data,
        `Added section ${number} ${title} as {${ph.toUpperCase()}}. Update section-catalog.md and Word TOC (F9) when ready.`,
      )
    },
  })

  pi.registerTool({
    name: 'manual_list_staging',
    label: 'List staged sections',
    description: 'Show which sections are staged in state.json vs still pending.',
    parameters: Type.Object({
      state_path: Type.String(),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      if (!sp) return toolError('state_path required.')
      const r = await runManualCli(['list-staging', sp])
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Staging summary loaded.')
    },
  })

  pi.registerTool({
    name: 'manual_build_draft',
    label: 'Build draft from staging',
    description:
      'Inject all staged prose sections (and §6 HMI if finalized screens exist) into draft.docx. Call when operator confirms all sections are approved and staged.',
    parameters: Type.Object({
      state_path: Type.String(),
      skip_hmi: Type.Optional(Type.Boolean({ description: 'Set true if §6 handled separately' })),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      if (!sp) return toolError('state_path required.')
      const args = ['build-draft', sp]
      if (params.skip_hmi === true) args.push('--skip-hmi')
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Staged content written to draft.docx.')
    },
  })

  pi.registerTool({
    name: 'manual_stage_hmi_screen',
    label: 'Stage HMI screen',
    description:
      'Save one approved §6 operator screen (title, body, image path) to state.json staging. Call after each screenshot+description; does not touch draft.docx yet.',
    parameters: Type.Object({
      state_path: Type.String({ description: 'projects/<id>/state.json' }),
      title: Type.String({ description: 'Screen name e.g. Main Dashboard' }),
      body: Type.String({ description: 'Approved subsection prose' }),
      image_path: Type.Optional(
        Type.String({ description: 'Path to screenshot saved under projects/<id>/inputs/' }),
      ),
      placeholder: Type.Optional(
        Type.String({ description: 'Default S06_SECTION — must match template slug' }),
      ),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      const title = String(params.title ?? '').trim()
      const body = String(params.body ?? '')
      if (!sp || !title || !body.trim()) return toolError('state_path, title, and body required.')
      const tmp = join(tmpdir(), `sylo-hmi-${randomUUID()}.txt`)
      writeFileSync(tmp, body, 'utf8')
      try {
        const args = ['hmi-stage', sp, '--title', title, '--body-file', tmp]
        if (params.image_path) args.push('--image', String(params.image_path))
        if (params.placeholder) args.push('--placeholder', String(params.placeholder))
        const r = await runManualCli(args)
        if (!r.ok) return toolError(r.error)
        return ok(r.data, `Staged HMI screen "${title}". Ask for next screen or finalize when done.`)
      } finally {
        try {
          unlinkSync(tmp)
        } catch {
          /* ignore */
        }
      }
    },
  })

  pi.registerTool({
    name: 'manual_finalize_hmi',
    label: 'Finalize HMI section',
    description:
      'Inject staged §6 HMI screens into draft.docx at {S06_SECTION}. Call when operator confirms all screens are captured.',
    parameters: Type.Object({
      state_path: Type.String(),
      placeholder: Type.Optional(Type.String()),
      width_inches: Type.Optional(Type.Number({ minimum: 1, maximum: 8 })),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      if (!sp) return toolError('state_path required.')
      const args = ['hmi-finalize', sp]
      if (params.placeholder) args.push('--placeholder', String(params.placeholder))
      if (typeof params.width_inches === 'number') args.push('--width', String(params.width_inches))
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      return ok(r.data, '§6 HMI injected into draft.docx from staging.')
    },
  })

  pi.registerTool({
    name: 'manual_hmi_staging',
    label: 'Read HMI staging',
    description: 'List staged §6 screens in state.json (resume mid-HMI without relying on chat history).',
    parameters: Type.Object({
      state_path: Type.String(),
      placeholder: Type.Optional(Type.String()),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      if (!sp) return toolError('state_path required.')
      const args = ['hmi-staging', sp]
      if (params.placeholder) args.push('--placeholder', String(params.placeholder))
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      const data = r.data as { screen_count?: number }
      return ok(r.data, `${data.screen_count ?? 0} HMI screen(s) staged.`)
    },
  })

  pi.registerTool({
    name: 'manual_replace_access_table',
    label: 'Replace access/password table',
    description:
      'Optional: bulk-fill an access table already in draft (first column header System). Normally §12 is a pipe table from manual_stage_section; use this only when operator supplies JSON rows.',
    parameters: Type.Object({
      draft_path: Type.String(),
      rows_json_path: Type.String({
        description:
          'Path to JSON file: [{system, purpose, username, password, notes}, ...]. Write file in project inputs/ first.',
      }),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      const rowsPath = String(params.rows_json_path ?? '').trim()
      if (!draft || !rowsPath) return toolError('draft_path and rows_json_path required.')
      const r = await runManualCli(['access-table', draft, rowsPath])
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Access table replaced (§21).')
    },
  })

  pi.registerTool({
    name: 'manual_append_access_row',
    label: 'Append access table row',
    description: 'Add one row to the marked access table (table-formats.json id: access).',
    parameters: Type.Object({
      draft_path: Type.String(),
      system: Type.String(),
      purpose: Type.String(),
      username: Type.Optional(Type.String()),
      password: Type.Optional(Type.String()),
      notes: Type.Optional(Type.String()),
    }),
    async execute(_id, params) {
      const draft = String(params.draft_path ?? '').trim()
      if (!draft) return toolError('draft_path required.')
      const args = [
        'access-row',
        draft,
        '--system',
        String(params.system ?? ''),
        '--purpose',
        String(params.purpose ?? ''),
        '--username',
        String(params.username ?? ''),
        '--password',
        String(params.password ?? ''),
        '--notes',
        String(params.notes ?? ''),
      ]
      const r = await runManualCli(args)
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Access table row appended.')
    },
  })

  pi.registerTool({
    name: 'manual_project_state',
    label: 'Manual project state',
    description: 'Read state.json checklist (completed / pending placeholders, paths).',
    parameters: Type.Object({
      state_path: Type.String({ description: 'Path to projects/<id>/state.json' }),
    }),
    async execute(_id, params) {
      const sp = String(params.state_path ?? '').trim()
      if (!sp) return toolError('state_path required.')
      const r = await runManualCli(['state', sp])
      if (!r.ok) return toolError(r.error)
      return ok(r.data, 'Project state loaded.')
    },
  })
}
