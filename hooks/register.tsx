// Show your work: a plugin for Claude Code that makes Claude show exactly
// where each claim comes from, so the reader can check it in one press.
//
// Claude cites a source through one tool, `show_source`: a Word, PowerPoint or
// Excel file, a PDF, a text or code file, an HTML file or a web page, with the
// words it rests on. The plugin
//   1. looks the words up IN THE FILE, through the program that owns it, and
//      tells Claude whether they are there (a claim whose words are not there
//      comes back as not found);
//   2. gives the citation a number, which Claude writes after its claim;
//   3. draws that number, in the answer, as a link that reads as the place —
//      «s. 7», «rad 250», «linje 16» — and opens the source AT the words;
//   4. lists every source of the session in a pane, to sort, filter and press.
//
// The number is Claude's and the plugin's. The reader never sees it: where
// Claude wrote «[3]» they see the place (Emil, 2026-11-25: numbers in brackets
// on a button are noise; a link in the sentence is how a paper cites).
//
// Why the answer is drawn here: only a `Markdown` element of the plugin's own
// hands it the press on a link. The link's address is never fetched — it only
// says which source was pressed.
//
// Windows only, for now: Office is driven through COM, Notepad through UI
// Automation, and the script beside this module (scripts/show.ps1) is
// PowerShell. PDFs go through PDF Scholar (github.com/emilmsh/pdf-scholar),
// whose own server script the plugin runs — so its switches decide what is
// shared, here as everywhere.
import { atom, read, update } from 'claude-code'
import type { EngineInterface, Register } from 'claude-code'

import type { Kind, Mode, Source, Status, View } from '../types'

// ---------- State ----------
//
// Change a tag when a value's fields change: it declines what an earlier
// version stored, which outlives a reload.

const sources = atom({ plugin: 'show-your-work', key: 'sources' } as const, [], { shape: 'sources-1' })
const reader = atom({ plugin: 'show-your-work', key: 'reader' } as const, '', { shape: 'reader-1' })
const view = atom(
  { plugin: 'show-your-work', key: 'view' } as const,
  { sort: 'cited', filter: 'all', group: 'none' },
  { shape: 'view-2' },
)

const uncited = atom({ plugin: 'show-your-work', key: 'uncited' } as const, [], { shape: 'uncited-1' })

/** The most kept: a session that cites more than this has long since scrolled
 *  the first ones out of reach */
const MOST = 500

const PANE = 'sources'

/** «[3]» in an answer — not «[3, 5]» or «[17–19]», a paper's own citations
 *  quoted along; not a markdown link's «[3](…)»; and not a «[3]» the answer
 *  only talks ABOUT, in quotes or as code */
const CITED = /(?<![«"“`])[ \t]*\[(\d{1,6})\](?![(»"”`])/g

/** A citation's link target inside an answer. Nothing is ever fetched from
 *  it: the press is the plugin's own, and the number says which source was
 *  pressed. (Markdown draws only http, https and file links.) */
const HREF = 'https://source.invalid/'

const KINDS: readonly Kind[] = ['word', 'powerpoint', 'excel', 'pdf', 'text', 'html', 'web']

const EXTENSIONS: Record<Exclude<Kind, 'web'>, readonly string[]> = {
  word: ['docx', 'docm', 'doc', 'rtf'],
  powerpoint: ['pptx', 'pptm', 'ppt'],
  excel: ['xlsx', 'xlsm', 'xls', 'csv'],
  pdf: ['pdf'],
  html: ['html', 'htm'],
  // Prose goes to Notepad, code to VS Code where it is installed: the script
  // tells them apart, here they are one kind — text looked up by line
  text: [
    'md', 'markdown', 'txt', 'text', 'log', 'tex',
    'r', 'rmd', 'qmd', 'py', 'do', 'ado', 'sql', 'js', 'jsx', 'ts', 'tsx', 'json', 'yaml', 'yml',
    'toml', 'css', 'sh', 'ps1', 'c', 'h', 'cpp', 'cs', 'java', 'go', 'rs', 'jl', 'm', 'sas',
  ],
}

// ---------- The reader's language ----------

type Words = {
  page: (n: number) => string
  line: (n: number) => string
  slide: (n: number) => string
  rows: (span: string) => string
  columns: (span: string) => string
  sheet: (name: string) => string
  pane: string
  none: string
  noneFit: string
  count: (shown: number, all: number) => string
  sort: string
  sorts: Record<View['sort'], string>
  group: string
  groups: Record<View['group'], string>
  filter: string
  all: string
  missing: string
  unchecked: string
  kinds: Record<Kind, string>
  opening: (name: string) => string
  failed: (name: string) => string
  command: string
  opened: string
  notHere: string
  mode: string
  modes: Record<Mode, string>
  nextSession: string
  uncited: (names: string) => string
  cite: string
}

const EN: Words = {
  page: n => `p. ${n}`,
  line: n => `line ${n}`,
  slide: n => `slide ${n}`,
  rows: span => `row ${span}`,
  columns: span => `column ${span}`,
  sheet: name => `sheet ${name}`,
  pane: 'Sources',
  none: 'No sources yet.',
  noneFit: 'No sources fit the filter.',
  count: (shown, all) => (shown === all ? `${all} sources` : `${shown} of ${all} sources`),
  sort: 'Sort',
  sorts: { cited: 'In the order cited', file: 'By file', kind: 'By kind' },
  group: 'Group',
  groups: { none: 'One list', file: 'By file', kind: 'By kind' },
  filter: 'Show',
  all: 'All',
  missing: 'Not found',
  unchecked: 'not checked',
  kinds: {
    word: 'Word',
    powerpoint: 'PowerPoint',
    excel: 'Excel',
    pdf: 'PDF',
    text: 'Text and code',
    html: 'HTML files',
    web: 'Web pages',
  },
  opening: name => `Opening ${name} …`,
  failed: name => `${name}: could not be opened`,
  command: 'Open the pane with the sources Claude has cited in this session',
  opened: 'The Sources pane is open.',
  notHere: 'The Sources pane cannot be shown here.',
  mode: 'Cite',
  modes: { always: 'By itself', asked: 'When I ask', off: 'Off' },
  nextSession: 'Applies fully from your next session.',
  uncited: names => `Claude read ${names} and showed no sources.`,
  cite: 'Show the sources',
}

const NB: Words = {
  page: n => `s. ${n}`,
  line: n => `linje ${n}`,
  slide: n => `lysbilde ${n}`,
  rows: span => `rad ${span}`,
  columns: span => `kolonne ${span}`,
  sheet: name => `ark ${name}`,
  pane: 'Kilder',
  none: 'Ingen kilder ennå.',
  noneFit: 'Ingen kilder passer filteret.',
  count: (shown, all) => (shown === all ? `${all} kilder` : `${shown} av ${all} kilder`),
  sort: 'Sorter',
  sorts: { cited: 'I den rekkefølgen de ble vist til', file: 'Etter fil', kind: 'Etter type' },
  group: 'Grupper',
  groups: { none: 'Én liste', file: 'Etter fil', kind: 'Etter type' },
  filter: 'Vis',
  all: 'Alle',
  missing: 'Ikke funnet',
  unchecked: 'ikke kontrollert',
  kinds: {
    word: 'Word',
    powerpoint: 'PowerPoint',
    excel: 'Excel',
    pdf: 'PDF',
    text: 'Tekst og kode',
    html: 'HTML-filer',
    web: 'Nettsider',
  },
  opening: name => `Åpner ${name} …`,
  failed: name => `${name}: kunne ikke åpnes`,
  command: 'Åpne panelet med kildene Claude har vist til i denne sesjonen',
  opened: 'Kildepanelet er åpnet.',
  notHere: 'Kildepanelet kan ikke vises her.',
  mode: 'Siter',
  modes: { always: 'Av seg selv', asked: 'Når jeg ber om det', off: 'Av' },
  nextSession: 'Gjelder fullt ut fra neste sesjon.',
  uncited: names => `Claude leste ${names} og viste ingen kilder.`,
  cite: 'Vis kildene',
}

/** The language the machine is set to, which is as close as a plugin gets to
 *  the reader's. English where there is no entry here, or it cannot be asked. */
function words(): Words {
  let language = 'en'

  try {
    language = new Intl.DateTimeFormat().resolvedOptions().locale.toLowerCase().split('-')[0] ?? 'en'
  } catch {
    // No Intl here: English
  }

  return ['nb', 'nn', 'no'].includes(language) ? NB : EN
}

// ---------- A source's kind, name and place ----------

function kindOf(path: string): Kind | undefined {
  if (/^https?:\/\//i.test(path)) {
    return 'web'
  }

  const extension = /\.([a-z0-9]+)$/i.exec(path)?.[1]?.toLowerCase() ?? ''

  return (Object.keys(EXTENSIONS) as Exclude<Kind, 'web'>[]).find(kind =>
    EXTENSIONS[kind].includes(extension),
  )
}

const nameOf = (path: string): string => path.split(/[\\/]/).pop() ?? path

const hostOf = (url: string): string => /^https?:\/\/([^/]+)/i.exec(url)?.[1] ?? url

/** An Excel place in words: «AN250», «row 250», «column C–E», «sheet
 *  Assumptions»; with the sheet in front when asked, «Results · row 5» */
function excelPlace(cell: string, isWithSheet: boolean, the: Words): string {
  const at = cell.lastIndexOf('!')
  const isAddress = /^\$?[a-z]{0,3}\$?\d*(:\$?[a-z]{0,3}\$?\d*)?$/i.test(cell.trim())
  const sheet = at > 0 ? cell.slice(0, at).replace(/^'|'$/g, '') : isAddress ? '' : cell.trim()
  const where = at >= 0 ? cell.slice(at + 1).trim() : isAddress ? cell.trim() : ''
  const rows = /^(\d+):(\d+)$/.exec(where)
  const columns = /^\$?([a-z]+):\$?([a-z]+)$/i.exec(where)
  const span = (from: string, to: string): string => (from === to ? from : `${from}–${to}`)
  const place =
    where === ''
      ? ''
      : rows
        ? the.rows(span(rows[1] ?? '', rows[2] ?? ''))
        : columns
          ? the.columns(span((columns[1] ?? '').toUpperCase(), (columns[2] ?? '').toUpperCase()))
          : where

  if (place === '') {
    return the.sheet(sheet)
  }

  return isWithSheet && sheet !== '' ? `${sheet} · ${place}` : place
}

/** Where the words are, as the pane's button says it: the place inside the
 *  source, or the source's own name where it has no places */
function placeOf(source: Source, the: Words): string {
  if (source.kind === 'excel') return excelPlace(source.cell, true, the)
  if (source.kind === 'web') return hostOf(source.path)
  if (source.kind === 'html' || source.page <= 0) return nameOf(source.path)
  if (source.kind === 'powerpoint') return the.slide(source.page)
  if (source.kind === 'text') return the.line(source.page)

  return the.page(source.page)
}

/** What stands IN THE SENTENCE where the model wrote «[3]»: the place, as a
 *  paper cites it. An answer resting on several files names the file too. */
function inline(source: Source, isNamed: boolean, the: Words): string {
  const place = source.kind === 'excel' ? excelPlace(source.cell, false, the) : placeOf(source, the)
  const name = nameOf(source.path)
  const text =
    isNamed && source.kind !== 'web' && place !== name
      ? `${name.replace(/\.[a-z0-9]+$/i, '')}, ${place}`
      : place

  // Nothing of a file's name may end the link early or start a new one
  return `${source.status === 'missing' ? '✗ ' : ''}${text}`.replace(/[[\]\\]/g, '')
}

// ---------- Looking a source up, and showing it ----------

type Located = { status: Status; page: number; detail: string; error?: string }

/** The script beside this module, with one request: its answer is one JSON
 *  line on standard output */
async function script(
  $: EngineInterface,
  request: Record<string, unknown>,
  timeoutMs: number,
): Promise<Record<string, unknown> | undefined> {
  const ran = await $.process.run(
    [
      'powershell.exe',
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      "& (Join-Path $env:CLAUDE_PLUGIN_ROOT 'scripts\\show.ps1')",
    ],
    { env: { ...(await $.process.hookEnv()), SYW_REQUEST: JSON.stringify(request) }, timeoutMs },
  )

  try {
    const answer: unknown = JSON.parse(ran.stdout.trim().split(/\r?\n/).pop() ?? '')

    return typeof answer === 'object' && answer !== null && !Array.isArray(answer)
      ? (answer as Record<string, unknown>)
      : undefined
  } catch {
    return undefined
  }
}

/** Everything but a PDF: Office, text, code, HTML and the web, through the script */
async function byScript($: EngineInterface, source: Source, isShown: boolean): Promise<Located> {
  let answer: Record<string, unknown> | undefined

  try {
    answer = await script(
      $,
      { kind: source.kind, path: source.path, quote: source.quote, cell: source.cell, show: isShown },
      60_000,
    )
  } catch {
    // The program did not answer in time: it is busy with a dialog of its
    // own, or the reader is editing a cell (Excel refuses automation then)
    return { status: 'missing', page: 0, detail: '', error: 'the program did not answer in a minute' }
  }

  if (answer === undefined) {
    return { status: 'missing', page: 0, detail: '', error: 'the lookup gave no answer' }
  }

  const text = typeof answer.text === 'string' ? answer.text : ''
  const formula = typeof answer.formula === 'string' && answer.formula.startsWith('=') ? answer.formula : ''
  const detail =
    source.kind !== 'excel'
      ? answer.where === 'notes'
        ? 'in the speaker notes'
        : ''
      : answer.whole === 'sheet'
        ? 'is a sheet of the workbook'
        : answer.whole === 'range'
          ? answer.found === true
            ? `holds «${text}» in ${typeof answer.at === 'string' ? answer.at : 'one of its cells'}`
            : 'holds no cell with the value cited'
          : `shows «${text}»${formula === '' ? '' : ` (formula ${formula})`}`

  return {
    status: source.kind === 'web' ? 'unchecked' : answer.found === true ? 'found' : 'missing',
    page: typeof answer.page === 'number' ? answer.page : 0,
    detail,
    ...(typeof answer.error === 'string' ? { error: answer.error } : {}),
  }
}

/** A PDF, through PDF Scholar. The plugin runs the reader's own server script
 *  itself, one call per process, the way an AI tool starts it — so it does
 *  not matter whether this session has PDF Scholar connected, and the server
 *  starts the app when it is not running. PDF Scholar's switches still
 *  decide: with «Read together» off the server refuses, and says so.
 *  Every process is started for the SAME session (`--session`) and told the
 *  citation's number here (`number`), so PDF Scholar's list is one list and
 *  its cards carry these numbers. PDF Scholar 0.87.0 and older read
 *  neither: there every call is a list of its own, numbered 1. */
async function byPdfScholar($: EngineInterface, source: Source, isShown: boolean): Promise<Located> {
  const none = (error: string): Located => ({ status: 'missing', page: 0, detail: '', error })
  const local = await $.env.get('LOCALAPPDATA')
  const roaming = await $.env.get('APPDATA')
  const exe = `${local ?? ''}\\Programs\\pdfx\\PDF Scholar.exe`
  const server = `${roaming ?? ''}\\PDF Scholar\\mcp\\pdf-scholar-mcp.cjs`

  if (local === undefined || roaming === undefined || !(await $.fs.exists(exe))) {
    return none('PDF Scholar is not installed on this machine (github.com/emilmsh/pdf-scholar)')
  }

  if (!(await $.fs.exists(server))) {
    return none('PDF Scholar is not sharing with AI tools: turn on a tool and «Read together» in its settings menu → AI tools')
  }

  let session = ''

  await update($, reader, id => {
    session =
      id !== '' ? id : Array.from({ length: 16 }, () => Math.floor(Math.random() * 16).toString(16)).join('')

    return session
  })

  const rpc = [
    {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'show-your-work', version: '1' },
      },
    },
    { jsonrpc: '2.0', method: 'notifications/initialized' },
    {
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/call',
      params: {
        name: 'show_passage',
        arguments: {
          path: source.path,
          page: source.page || 1,
          quote: source.quote,
          ...(source.claim === '' ? {} : { claim: source.claim }),
          ...(source.n > 0 ? { number: source.n } : {}),
          show: isShown,
        },
      },
    },
  ]

  try {
    const ran = await $.process.run([exe, server, '--client', 'claude', '--session', session], {
      env: { ELECTRON_RUN_AS_NODE: '1' },
      stdin: `${rpc.map(one => JSON.stringify(one)).join('\n')}\n`,
      // The server waits up to 75 s for a passage in an app it had to start
      timeoutMs: 90_000,
    })
    const result = ran.stdout
      .split(/\r?\n/)
      .filter(line => line.trim().startsWith('{'))
      .map(line => JSON.parse(line) as { id?: number; result?: { content?: { text?: string }[]; isError?: boolean } })
      .find(one => one.id === 2)?.result
    const text = result?.content?.[0]?.text ?? ''

    if (result === undefined) {
      return none('PDF Scholar gave no answer')
    }

    if (result.isError === true) {
      return none(text)
    }

    // The reader pressed a link to get here: the window is theirs to work in
    // at once. PDF Scholar itself only comes up ABOVE the program in front,
    // without the keyboard — right when a tool shows something unasked,
    // wrong after a press (it read as frozen until clicked).
    if (isShown) {
      await script($, { kind: 'raise', process: 'PDF Scholar' }, 20_000).catch(() => undefined)
    }

    return {
      status: /NOT FOUND|could not be checked|Only part of the quote/.test(text) ? 'missing' : 'found',
      page: Number(/on page (\d+)/.exec(text)?.[1] ?? source.page),
      detail: '',
    }
  } catch {
    return none('PDF Scholar did not answer')
  }
}

const locate = (source: Source, isShown: boolean, $: EngineInterface): Promise<Located> =>
  source.kind === 'pdf' ? byPdfScholar($, source, isShown) : byScript($, source, isShown)

/** A press: open the source at the words. Starting Word or PowerPoint cold
 *  takes longer than a press may hold (10 s), so the press only STARTS this
 *  and says that it has; a failure is said when it is known. */
function show($: EngineInterface, source: Source): void {
  const the = words()
  const name = source.kind === 'web' ? hostOf(source.path) : nameOf(source.path)

  $.ui.toast(the.opening(name))

  void (async () => {
    try {
      const answer = await locate(source, true, $)

      if (answer.error !== undefined) {
        $.ui.toast(`${name}: ${answer.error}`)
      }
    } catch {
      $.ui.toast(the.failed(name))
    }
  })()
}

// ---------- What the model is told ----------

const DESCRIPTION =
  'Cite a passage in a source the user can open: a Word, PowerPoint or Excel file, a PDF, a plain text file (.md, .txt, .log, .tex), a code file (.R, .py, .do, .sql, .ts …), an HTML file — each by absolute path — or a web page (https URL). The words are looked up in the file itself, and the answer opens with the citation’s number, «[3] …». Write that number, as [3], right after the claim it backs and nowhere else: the user never sees the number — where you wrote it they see a link reading the place («p. 7», «line 19»), which opens the source at the words. Quote word for word. For Excel give `cell` — a cell («Sheet!B2»), a row, column or block («Sheet!5:5», «Sheet!C:C», «Sheet!B2:D10») or a sheet alone («Sheet!») — and, as `quote`, the value you say is there (for a sheet alone, its name). Call it for every claim you rest on a source; set `show` only when the user asks to see the passage now. A NOT FOUND answer means the claim is unverified: say so, or quote the source’s own words.'

/** The tool's answer: the number first, in the form the model writes it in */
function answerOf(source: Source, located: Located): string {
  const n = `[${source.n}] `
  // The model is answered in English whatever the reader's language: the
  // place is told as the file has it, the link is the reader's
  const place = placeOf(source, EN)
  const name = source.kind === 'web' ? hostOf(source.path) : nameOf(source.path)
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

// ---------- The pane ----------

/** The sources the pane lists, as the reader has it arranged */
function arranged(list: readonly Source[], { sort, filter }: View): Source[] {
  const kept = list.filter(one =>
    filter === 'all' ? true : filter === 'missing' ? one.status === 'missing' : one.kind === filter,
  )
  const byName = (a: Source, b: Source): number =>
    nameOf(a.path).localeCompare(nameOf(b.path), undefined, { numeric: true }) || a.n - b.n

  if (sort === 'file') {
    return [...kept].sort(byName)
  }

  if (sort === 'kind') {
    return [...kept].sort((a, b) => KINDS.indexOf(a.kind) - KINDS.indexOf(b.kind) || byName(a, b))
  }

  return [...kept].sort((a, b) => a.n - b.n)
}

type Group = { key: string; title: string; items: Source[] }

/** The arranged sources under a heading per file or per kind — files by
 *  name, kinds in their fixed order; each group keeps the list's own order.
 *  Ungrouped, one group with no heading. */
function grouped(shown: readonly Source[], by: View['group'], the: Words): Group[] {
  if (by === 'none') {
    return [{ key: 'all', title: '', items: [...shown] }]
  }

  const groups = new Map<string, Group>()

  for (const source of shown) {
    const key = by === 'kind' ? source.kind : source.path
    const title =
      by === 'kind' ? the.kinds[source.kind] : source.kind === 'web' ? hostOf(source.path) : nameOf(source.path)
    const group = groups.get(key) ?? { key, title, items: [] }

    group.items.push(source)
    groups.set(key, group)
  }

  return [...groups.values()].sort((a, b) =>
    by === 'kind'
      ? KINDS.indexOf(a.key as Kind) - KINDS.indexOf(b.key as Kind)
      : a.title.localeCompare(b.title, undefined, { numeric: true }),
  )
}

const isSort = (value: string): value is View['sort'] => ['cited', 'file', 'kind'].includes(value)

const isGroup = (value: string): value is View['group'] => ['none', 'file', 'kind'].includes(value)

const isFilter = (value: string): value is View['filter'] =>
  value === 'all' || value === 'missing' || (KINDS as readonly string[]).includes(value)

/** Everything this plugin does to a program is Windows' (scripts/show.ps1,
 *  PDF Scholar's install path). On another system it offers NOTHING — no tool
 *  for the model to call and be let down by — until that system has a way of
 *  its own. A machine that cannot be asked is taken for Windows: the script
 *  then says what went wrong by itself. */
async function isWindows($: EngineInterface): Promise<boolean> {
  try {
    return (await $.session.machine()).sysname === 'Windows_NT'
  } catch {
    return true
  }
}

// ---------- When Claude cites ----------
//
// A tool the model MAY call is not a habit: left to itself it writes «the
// report says …» and cites nothing (Emil, 2026-11-25). Three things make
// citing the default, and the reader's to turn down:
//   - a standing instruction in the system prompt, while the mode is 'always';
//   - an offer under the prompt when a turn read documents and cited none,
//     which asks for the sources in one press — it does not depend on the
//     model having remembered anything;
//   - the mode itself, kept across sessions: by itself, when asked, or off.

/** What the model is told at the start of every session while citing is on.
 *  It says WHEN (a claim about a document the reader can open), WHAT (the
 *  claims that matter, not every sentence — a look-up takes seconds) and what
 *  a miss means. */
const RULE = [
  'The user has installed «Show your work» so that they can check what you say against the source in one press.',
  'When you tell the user what a document says — a Word, PowerPoint or Excel file, a PDF, a note, or a web page they can open — back the claims that matter with the show_source tool: figures, findings, dates, and anything you quote.',
  'Call it once per claim, several calls in parallel, with the source’s own words verbatim, and write the number it answers with, like [3], right after the claim. The user sees a link there that opens the source at the words.',
  'Do not cite every sentence, and do not cite code you are editing or a file you only skimmed for context.',
  'If show_source answers NOT FOUND, the words are not in the source: correct the claim or say it is unverified. Never write a page, slide or cell from memory in place of calling the tool.',
].join(' ')

/** The reader's choice, kept across sessions; 'always' until they choose */
async function modeOf($: EngineInterface): Promise<Mode> {
  try {
    const kept = await $.store.get('mode')

    return kept === 'asked' || kept === 'off' ? kept : 'always'
  } catch {
    return 'always'
  }
}

/** The kinds a turn is offered sources for: documents, never the code and
 *  notes a coding session reads all day */
const DOCUMENT = /[^\s"'<>|*?]+\.(docx?|pptx?|xlsx?|pdf)\b/gi

const ELSEWHERE = 'Show your work runs on Windows only for now, so it cannot check or open a source on this machine.'

export const register: Register = on => {
  on('session.start', async ($, e, next) => {
    if (!(await isWindows($))) {
      return next(e)
    }

    await $.tool.register({
      name: 'show_source',
      isDeferred: false,
      description: DESCRIPTION,
      inputSchema: {
        type: 'object',
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
    await $.command.register({ name: 'sources', description: words().command })

    return next(e)
  })

  on('prompt.compose', async ($, e, next) => {
    const composed = await next(e)

    if ((await modeOf($)) !== 'always' || !(await isWindows($))) {
      return composed
    }

    return {
      ...composed,
      sections: [...composed.sections, { id: 'show-your-work:cite', text: RULE, scope: 'session' as const }],
    }
  })

  // A new turn starts clean: what the last one read is no longer on offer
  on('prompt.submit', async ($, e, next) => {
    await update($, uncited, () => [])

    return next(e)
  })

  // Every tool call is looked at for the documents it names — a Read, a
  // script that opens one — and nothing else of it is kept
  on('tool.call', async ($, e, next) => {
    if (e.tool !== 'mcp__show-your-work__show_source') {
      const named = [...JSON.stringify(e).matchAll(DOCUMENT)].map(match => nameOf(match[0].replace(/\\\\/g, '/')))

      if (named.length > 0) {
        await update($, uncited, list => [...new Set([...list, ...named])].slice(-20))
      }
    }

    return next(e)
  })

  on('ui.render', { component: 'AbovePrompt' }, async ($, e, next) => {
    const read_ = await read($, uncited)

    if (e.props.hasSurvey || read_.length === 0 || (await modeOf($)) === 'off' || !(await isWindows($))) {
      return next(e)
    }

    const { Box, Button, Text } = $.ui.resolve(e)
    const the = words()
    const names = read_.length > 3 ? `${read_.slice(0, 3).join(', ')} …` : read_.join(', ')

    return (
      <Box flexDirection="row" columnGap={1}>
        <Text dimColor>{the.uncited(names)}</Text>
        <Button
          key="cite"
          label={the.cite}
          onPress={async () => {
            await update($, uncited, () => [])
            await $.prompt.submit({
              text: 'Show me where each claim in your last answer is said: cite them with show_source, and tell me if any is not in the source.',
              asUser: true,
            })
          }}
        />
      </Box>
    )
  })

  // The pane opens by itself at the first citation; this brings it back
  // after it has been closed, and opens it before there is any
  on('command.run', { command: 'sources' }, async $ => {
    const the = words()
    const opened = await $.ui.open({ id: PANE, title: the.pane })

    return { text: opened.isPlaced ? the.opened : the.notHere }
  })

  on('tool.run', { tool: 'mcp__show-your-work__show_source' }, async ($, e) => {
    if (!(await isWindows($))) {
      return { deny: ELSEWHERE }
    }

    const input = e as unknown as Record<string, unknown>
    const path = typeof input.source === 'string' ? input.source.trim() : ''
    const kind = kindOf(path)
    const quote = typeof input.quote === 'string' ? input.quote.trim() : ''

    if ((await modeOf($)) === 'off') {
      return { deny: 'The user has turned Show your work off. Cite in plain text, and say that the links are off.' }
    }

    if (kind === undefined || quote === '') {
      return {
        deny: 'Give `source` (a Word, PowerPoint, Excel, PDF, HTML, text or code file by absolute path, or an https URL) and `quote`.',
      }
    }

    if (kind === 'excel' && typeof input.cell !== 'string') {
      return { deny: 'For an Excel file give `cell`, as «Sheet!B2».' }
    }

    const asked: Source = {
      n: 0,
      kind,
      path,
      quote,
      cell: typeof input.cell === 'string' ? input.cell.trim() : '',
      page: typeof input.page === 'number' && input.page >= 1 ? Math.floor(input.page) : 0,
      claim: typeof input.claim === 'string' ? input.claim.trim() : '',
      status: 'unchecked',
    }
    let made = asked

    // The number FIRST: a source in a program with a list of its own (PDF
    // Scholar) is told it, so its card and this list read the same. The same
    // words in the same source again are the same citation, and a number is
    // never given twice — the next is one past the highest ever given.
    await update($, sources, list => {
      const same = list.find(
        one => one.path === asked.path && one.quote === asked.quote && one.cell === asked.cell,
      )

      made = { ...asked, n: same?.n ?? list.reduce((most, one) => Math.max(most, one.n), 0) + 1 }

      return same === undefined ? [...list, made].slice(-MOST) : list
    })

    const located = await locate(made, input.show === true, $)

    // Cited: the turn is no longer one that read documents and showed nothing
    await update($, uncited, () => [])

    made = { ...made, page: located.page || asked.page, status: located.status }
    await update($, sources, list => list.map(one => (one.n === made.n ? made : one)))
    // Nowhere to draw a pane (a phone, a run with no window) is no failure
    await $.ui.open({ id: PANE, title: words().pane })

    return { result: answerOf(made, located) }
  })

  on('ui.render', { component: 'Pane', requestId: PANE }, async ($, e) => {
    const { Box, Text, Button, Select } = $.ui.resolve(e)
    const the = words()
    const list = await read($, sources)
    const how = await read($, view)
    const shown = arranged(list, how)
    const kinds = KINDS.filter(kind => list.some(one => one.kind === kind))
    const lost = list.filter(one => one.status === 'missing').length
    const mode = await modeOf($)
    const modes = (
      <Select
        key="mode"
        label={the.mode}
        value={mode}
        options={(['always', 'asked', 'off'] as const).map(value => ({ value, label: the.modes[value] }))}
        onSelect={value => {
          if (value === 'always' || value === 'asked' || value === 'off') {
            void $.store.set('mode', value).then(() => {
              $.ui.toast(the.nextSession)
              $.ui.invalidate('ui.render')
            })
          }
        }}
      />
    )

    if (list.length === 0) {
      return (
        <Box flexDirection="column">
          {modes}
          <Text> </Text>
          <Text dimColor>{the.none}</Text>
        </Box>
      )
    }

    return (
      <Box flexDirection="column">
        {modes}
        <Box flexDirection="row" flexWrap="wrap" columnGap={2}>
          <Select
            key="sort"
            label={the.sort}
            value={how.sort}
            options={(['cited', 'file', 'kind'] as const).map(value => ({ value, label: the.sorts[value] }))}
            onSelect={value => {
              if (isSort(value)) {
                void update($, view, now => ({ ...now, sort: value }))
              }
            }}
          />
          <Select
            key="filter"
            label={the.filter}
            value={how.filter}
            options={[
              { value: 'all', label: `${the.all} (${list.length})` },
              ...kinds.map(kind => ({
                value: kind,
                label: `${the.kinds[kind]} (${list.filter(one => one.kind === kind).length})`,
              })),
              ...(lost > 0 || how.filter === 'missing' ? [{ value: 'missing', label: `${the.missing} (${lost})` }] : []),
            ]}
            onSelect={value => {
              if (isFilter(value)) {
                void update($, view, now => ({ ...now, filter: value }))
              }
            }}
          />
          <Select
            key="group"
            label={the.group}
            value={how.group}
            options={(['none', 'file', 'kind'] as const).map(value => ({ value, label: the.groups[value] }))}
            onSelect={value => {
              if (isGroup(value)) {
                void update($, view, now => ({ ...now, group: value }))
              }
            }}
          />
        </Box>
        <Text dimColor>{the.count(shown.length, list.length)}</Text>
        <Text> </Text>
        {shown.length === 0 && <Text dimColor>{the.noneFit}</Text>}
        {grouped(shown, how.group, the).map(group => (
          <Box flexDirection="column" key={`group-${group.key}`}>
            {group.title !== '' && (
              <Text bold>
                {group.title} ({group.items.length})
              </Text>
            )}
            {group.items.map(source => (
              <Box flexDirection="column" key={`row-${source.n}`}>
                <Box flexDirection="row" columnGap={1}>
                  <Button
                    key={`source-${source.n}`}
                    label={`${source.status === 'missing' ? '✗ ' : ''}${placeOf(source, the)}`}
                    dimColor={source.status === 'missing'}
                    onPress={() => show($, source)}
                  />
                  <Text dimColor>
                    {[
                      // The file's name, unless the heading or the button already says it
                      how.group === 'file' || source.kind === 'web' || placeOf(source, the) === nameOf(source.path)
                        ? ''
                        : nameOf(source.path),
                      source.status === 'missing' ? the.missing.toLowerCase() : '',
                      source.status === 'unchecked' ? the.unchecked : '',
                    ]
                      .filter(part => part !== '')
                      .join(' · ')}
                  </Text>
                </Box>
                {source.claim !== '' && <Text>{source.claim}</Text>}
                <Text dimColor>
                  «{source.quote.length > 160 ? `${source.quote.slice(0, 160)} …` : source.quote}»
                </Text>
                <Text> </Text>
              </Box>
            ))}
          </Box>
        ))}
      </Box>
    )
  })

  on('ui.render', { component: 'AssistantMessage' }, async ($, e, next) => {
    const cited = [...e.props.text.matchAll(CITED)].map(match => Number(match[1]))

    if (cited.length === 0) {
      return next(e)
    }

    const list = await read($, sources)
    const mine = [...new Set(cited)]
      .map(n => list.find(one => one.n === n))
      .filter((one): one is Source => one !== undefined)

    if (mine.length === 0) {
      return next(e)
    }

    const { Markdown } = $.ui.resolve(e)
    const the = words()
    const isNamed = new Set(mine.map(source => source.path)).size > 1
    const byHref = new Map(mine.map(source => [`${HREF}${source.n}`, source]))
    // «[3]» in the sentence becomes a link that reads as the place; a number
    // that is no source of ours stays as written
    const text = e.props.text.replace(CITED, (whole, digits: string) => {
      const source = mine.find(one => one.n === Number(digits))

      return source === undefined ? whole : ` [${inline(source, isNamed, the)}](${HREF}${source.n})`
    })

    // The answer is drawn here, not by `next`: only a Markdown of the
    // plugin's own hands it the press on a link
    return (
      <Markdown
        key="answer"
        text={text}
        pressableLinks={[...byHref.keys()]}
        onLinkPress={link => {
          const source = byHref.get(link.href)

          if (source !== undefined) {
            show($, source)
          }
        }}
      />
    )
  })
}
