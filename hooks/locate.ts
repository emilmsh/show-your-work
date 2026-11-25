// Looking a source up, and showing it: the one seam where the plugin reaches
// programs on the machine. Everything here is Windows' (scripts/show.ps1, PDF
// Scholar's install path); see CLAUDE.md § Platforms.
import type { EngineInterface } from 'claude-code'

import type { Source, Status } from '../types'
import { calledOf } from './source'
import { readerSession } from './state'
import { words } from './words'

/** What a look-up came to. `detail` is what the model is told beside the
 *  verdict (an Excel cell's value, «in the speaker notes»); `error` that the
 *  words could not be checked at all. */
export type Located = { status: Status; page: number; detail: string; error?: string }

const failure = (error: string): Located => ({ status: 'missing', page: 0, detail: '', error })

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

/** What the script found in an Excel workbook, in the words the model reads */
function excelDetail(answer: Record<string, unknown>): string {
  const text = typeof answer.text === 'string' ? answer.text : ''
  const formula = typeof answer.formula === 'string' && answer.formula.startsWith('=') ? answer.formula : ''

  if (answer.whole === 'sheet') {
    return 'is a sheet of the workbook'
  }

  if (answer.whole === 'range') {
    return answer.found === true
      ? `holds «${text}» in ${typeof answer.at === 'string' ? answer.at : 'one of its cells'}`
      : 'holds no cell with the value cited'
  }

  return `shows «${text}»${formula === '' ? '' : ` (formula ${formula})`}`
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
    return failure('the program did not answer in a minute')
  }

  if (answer === undefined) {
    return failure('the lookup gave no answer')
  }

  return {
    status: source.kind === 'web' ? 'unchecked' : answer.found === true ? 'found' : 'missing',
    page: typeof answer.page === 'number' ? answer.page : 0,
    detail:
      source.kind === 'excel' ? excelDetail(answer) : answer.where === 'notes' ? 'in the speaker notes' : '',
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
  const local = await $.env.get('LOCALAPPDATA')
  const roaming = await $.env.get('APPDATA')
  const exe = `${local ?? ''}\\Programs\\pdfx\\PDF Scholar.exe`
  const server = `${roaming ?? ''}\\PDF Scholar\\mcp\\pdf-scholar-mcp.cjs`

  if (local === undefined || roaming === undefined || !(await $.fs.exists(exe))) {
    return failure('PDF Scholar is not installed on this machine (github.com/emilmsh/pdf-scholar)')
  }

  if (!(await $.fs.exists(server))) {
    return failure('PDF Scholar is not sharing with AI tools: turn on a tool and «Read together» in its settings menu → AI tools')
  }

  const session = await readerSession($)
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
      return failure('PDF Scholar gave no answer')
    }

    if (result.isError === true) {
      return failure(text)
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
    return failure('PDF Scholar did not answer')
  }
}

// A function DECLARATION, as is every function here that takes $: past an
// import the engine follows $ only into those
export function locate(source: Source, isShown: boolean, $: EngineInterface): Promise<Located> {
  return source.kind === 'pdf' ? byPdfScholar($, source, isShown) : byScript($, source, isShown)
}

/** A press: open the source at the words. Starting Word or PowerPoint cold
 *  takes longer than a press may hold (10 s), so the press only STARTS this
 *  and says that it has; a failure is said when it is known. */
export function show($: EngineInterface, source: Source): void {
  const the = words()
  const name = calledOf(source)

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
