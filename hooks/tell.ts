// What the model is told: when to cite, how to call the tool, and what each
// call came to — and reading those answers back out of the transcript.
import type { EngineInterface } from 'claude-code'

import type { Mode, Source, Status } from '../types'
import type { Located } from './locate'
import { TOOL, askedOf, calledOf, placeOf } from './source'
import { noted } from './state'
import { EN } from './words'

// ---------- When Claude cites ----------
//
// A tool the model MAY call is not a habit: left to itself it writes «the
// report says …» and cites nothing (Emil, 2026-11-25). Three things make
// citing the default, and the reader's to turn down:
//   - the tool's own description says WHEN to call it, and how eagerly: that
//     is the one text of a plugin's the model reads at every turn. (A section
//     in the system prompt would be the obvious place and is not a plugin's
//     to add: the engine's own security plugin passes a plugin's
//     `prompt.compose` hook over. Measured 2026-11-25, and left alone.)
//   - an offer under the prompt when a turn read documents and cited none,
//     which asks for the sources in one press — it does not depend on the
//     model having remembered anything;
//   - the mode itself, kept across sessions: by itself, when asked, or off.

/** How the description opens, by the reader's mode: WHEN to cite. It says
 *  what (the claims that matter, not every sentence — a look-up takes
 *  seconds) and what is not a source (code being edited). */
const WHEN: Record<Exclude<Mode, 'off'>, string> = {
  always:
    'USE THIS WHENEVER you tell the user what a document says. The user installed «Show your work» to check your claims against the source in one press, so back the claims that matter — figures, findings, dates, anything you quote — from a Word, PowerPoint or Excel file, a PDF, a note or a web page with one call each, several in parallel, before you answer. Do not cite every sentence, and do not cite code you are editing or a file you only skimmed for context. ',
  asked:
    'Use this when the user asks where something is said, asks for sources, or asks you to show your work — not unasked. ',
}

const DESCRIPTION =
  'Cite a passage in a source the user can open: a Word, PowerPoint or Excel file, a PDF, a plain text file (.md, .txt, .log, .tex), a code file (.R, .py, .do, .sql, .ts …), an HTML file — each by absolute path — or a web page (https URL). The words are looked up in the file itself, and the answer opens with the citation’s number, «[3] …». Write that number, as [3], right after the claim it backs and nowhere else: the user never sees the number — where you wrote it they see a link reading the place («p. 7», «line 19»), which opens the source at the words. Quote word for word. For Excel give `cell` — a cell («Sheet!B2»), a row, column or block («Sheet!5:5», «Sheet!C:C», «Sheet!B2:D10») or a sheet alone («Sheet!») — and, as `quote`, the value you say is there (for a sheet alone, its name). Set `show` only when the user asks to see the passage now. A NOT FOUND answer means the claim is unverified: say so, or quote the source’s own words.'

/** The tool as the model sees it, worded by the reader's mode (its name stays
 *  at the call, where `claude plugin validate` reads it) */
export const toolFor = (mode: Exclude<Mode, 'off'>) => ({
  isDeferred: false,
  description: `${WHEN[mode]}${DESCRIPTION}`,
  inputSchema: {
    type: 'object' as const,
    properties: {
      source: { type: 'string', description: 'Absolute path of the file, or the page’s URL' },
      quote: { type: 'string', description: 'The words, verbatim; for Excel, the value cited' },
      cell: {
        type: 'string',
        description: 'Excel only: «Sheet!B2», «Sheet!5:5», «Sheet!C:C», «Sheet!B2:D10» or «Sheet!»',
      },
      page: { type: 'number', description: 'PDF only: the page the words are on' },
      claim: { type: 'string', description: 'The claim these words back, in one sentence' },
      show: { type: 'boolean', description: 'Open the source at the words right now' },
    },
    required: ['source', 'quote'],
  },
})

export const ELSEWHERE = 'Show your work runs on Windows only for now, so it cannot check or open a source on this machine.'

// ---------- What a call came to ----------
//
// The answer is written here and read back by `statusOfAnswer` and
// `sourceOf`: change a phrase in one and the other must follow.

/** The tool's answer: the number first, in the form the model writes it in */
export function answerOf(source: Source, located: Located): string {
  const n = `[${source.n}] `
  // The model is answered in English whatever the reader's language: the
  // place is told as the file has it, the link is the reader's
  const place = placeOf(source, EN)
  const name = calledOf(source)
  const where = place === name ? name : `${place} of ${name}`

  if (located.error !== undefined) {
    return `${n}Could not be checked (${located.error}). Treat the claim as unverified and tell the user.`
  }

  if (source.kind === 'web') {
    return `${n}Noted, NOT checked: a web page is not read. It opens at these words when the user presses the link.`
  }

  if (source.kind === 'excel') {
    return /^(is a sheet|holds)/.test(located.detail)
      ? `${n}${where} ${located.detail}.`
      : `${n}${where} ${located.detail}: ${located.status === 'found' ? 'matches the value cited' : 'does NOT match the value cited'}.`
  }

  return located.status === 'found'
    ? `${n}Found word for word, ${where}${located.detail === '' ? '' : ` (${located.detail})`}.`
    : `${n}NOT FOUND: these words are not in ${name}. Treat the claim as unverified and tell the user.`
}

/** The verdict an answer of `answerOf` gives */
const statusOfAnswer = (text: string): Status =>
  /NOT checked/.test(text)
    ? 'unchecked'
    : /NOT FOUND|does NOT match|holds no cell|Could not be checked/.test(text)
      ? 'missing'
      : 'found'

// ---------- The sources of a session ----------
//
// Two records, since the first one failed in the first session that had the
// plugin INSTALLED rather than written into it (2026-11-25: eight sources
// cited and checked, and the answer drawn with its bare «[1]» … «[8]» — the
// same thing the plugin before this one met a day earlier):
// - what the tool noted in the session's state as each call was answered;
// - the session's own transcript: every answered show_source call, with what
//   was asked and the answer's first line. This one needs nothing to have
//   been noted, and holds for a session resumed later.
// What was noted wins where both have a number; the transcript fills the rest.

/** A source as the transcript holds a call: what the model asked, and what
 *  the tool answered — «[3] Found word for word, p. 3 of report.docx.» */
function sourceOf(tool: string, input: Record<string, unknown>, text: string | undefined): Source | undefined {
  const n = Number(/^\[(\d{1,6})\]/.exec(text ?? '')?.[1] ?? 0)
  const asked = tool === TOOL ? askedOf(input) : undefined

  if (asked === undefined || text === undefined || n < 1) {
    return undefined
  }

  return {
    ...asked,
    n,
    page: Number(/ (?:p\.|line|slide) (\d+) of /.exec(text)?.[1] ?? asked.page),
    status: statusOfAnswer(text),
  }
}

/** Every source of the session: noted ones first, the transcript's for any
 *  number nobody noted, in the order of their numbers */
export async function known($: EngineInterface): Promise<Source[]> {
  const kept = await noted($)
  const all = new Map(kept.map(one => [one.n, one]))

  try {
    const messages = await $.session.messages({ last: 400 })

    if (!('deny' in messages)) {
      for (const message of messages) {
        for (const use of message.toolUses ?? []) {
          const made = sourceOf(use.tool, use.input as Record<string, unknown>, use.text)

          // The newest call of a number is the one kept
          if (made !== undefined && !kept.some(one => one.n === made.n)) {
            all.set(made.n, made)
          }
        }
      }
    }
  } catch {
    // Then only what was noted
  }

  return [...all.values()].sort((a, b) => a.n - b.n)
}
