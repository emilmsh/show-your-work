// Run: claude plugin test .   (Claude Code 2.1.330 or later)
//
// Nothing here starts Word, a browser or PDF Scholar: each test's own hooks
// stand for the engine beneath the plugin, and `process.run` is answered with
// what the lookup script (scripts/show.ps1) or PDF Scholar's server would say.
// What the script itself does is checked by hand on a Windows machine with
// Office — see the README.
import { expect, test } from 'claude-code/testing'

const TOOL = 'mcp__show-your-work__show_source'

const ANSWER = { plugin: 'show-your-work', component: 'AssistantMessage' } as const

const PANE = {
  plugin: 'show-your-work',
  component: 'Pane',
  requestId: 'sources',
  props: {
    title: 'Sources',
    isFocused: false,
    bodyColumns: 60,
    placement: 'dock',
    scroll: { offset: 0, bodyRows: 30 },
    view: {},
  },
} as const

const REPORT = 'C:\\work\\report.docx'
const MODEL = 'C:\\work\\model.xlsx'
const PAPER = 'C:\\work\\paper.pdf'

/** A test's own `on`: its hooks stand for the engine beneath the plugin */
type On = Parameters<Extract<Parameters<typeof test>[1], (...args: never[]) => unknown>>[1]

type Run = { argv: readonly string[]; request: Record<string, unknown>; stdin: string }

/** The tools the plugin has registered, by name, in every test of this file's run */
const registered: string[] = []

/** What the session's own record holds of earlier calls; a test sets it */
let transcript: { role: 'assistant'; text: string; toolUses: { tool_use_id: string; tool: string; input: Record<string, unknown>; text: string }[] }[] = []

/** …and what each told the model about itself */
const described: string[] = []

/** A program that ran and wrote this */
const ran = (stdout: string) => ({
  value: { exitCode: 0, stdout, stderr: '', isStdoutTruncated: false, isStderrTruncated: false },
})

/** The machine beneath the plugin. `script` answers a lookup by its request;
 *  PDF Scholar's server answers every passage as found on the page cited. */
function machine(
  on: On,
  script: (request: Record<string, unknown>) => Record<string, unknown> = () => ({ found: true, page: 3 }),
  sysname = 'Windows_NT',
): Run[] {
  const runs: Run[] = []

  on('session.messages', () => ({ value: transcript as never }))
  on('session.machine', () => ({
    value: { sysname, release: '10.0.26100', machine: 'x86_64', packageManagers: [], runtimes: [] },
  }))
  on('session.start', (_, e) => ({ cwd: e.cwd }))
  on('tool.register', (_, e) => {
    registered.push(e.name)
    described.push(e.description)

    return { value: { tool: TOOL } }
  })
  on('command.register', () => ({ value: { command: 'sources' } }))
  on('ui.open', () => ({ value: { isPlaced: true } }))
  on('ui.toast', () => ({ value: undefined }))
  on('process.hookEnv', () => ({ value: { CLAUDE_PLUGIN_ROOT: 'C:\\plugins\\show-your-work' } }))
  on('env.get', (_, e) => ({ value: `C:\\Users\\reader\\${e.name}` }))
  on('fs.exists', () => ({ value: true }))
  on('process.run', (_, e) => {
    const env = e.init?.env ?? {}
    const request = JSON.parse(env.SYW_REQUEST ?? '{}') as Record<string, unknown>
    const stdin = e.init?.stdin ?? ''

    runs.push({ argv: e.argv, request, stdin })

    if (e.argv[0] !== 'powershell.exe') {
      const call = stdin
        .split('\n')
        .filter(line => line !== '')
        .map(line => JSON.parse(line) as { id?: number; params?: { arguments?: { page?: number; number?: number } } })
        .find(one => one.id === 2)
      const text = `[${call?.params?.arguments?.number ?? 1}] Found word for word on page ${call?.params?.arguments?.page ?? 1}. It is in the user’s list of references in PDF Scholar.`

      // Electron writes a line break before any code runs: the answer is
      // found by its id, not by its place
      return ran(
        `\r\n${JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} })}\n${JSON.stringify({ jsonrpc: '2.0', id: 2, result: { content: [{ type: 'text', text }] } })}\n`,
      )
    }

    return ran(`${JSON.stringify(script(request))}\n`)
  })

  return runs
}

/** The tool as the model calls it: its answer's text, or why it was refused */
async function cite(
  $: Parameters<Extract<Parameters<typeof test>[1], (...args: never[]) => unknown>>[0],
  input: Record<string, unknown>,
): Promise<string> {
  const answer = (await $.tool.run({ tool: TOOL, ...input } as never)) as { result?: unknown; deny?: unknown }

  return String(answer.result ?? `DENIED: ${String(answer.deny)}`)
}

test('a citation is looked up in the file, numbered and answered', async ($, on) => {
  const runs = machine(on, request => (request.quote === 'not in the file' ? { found: false } : { found: true, page: 3 }))

  await $.session.start({ cwd: '/', surface: 'desktop', isInteractive: true })

  expect(await cite($, { source: REPORT, quote: 'The main finding holds' })).toBe(
    '[1] Found word for word, p. 3 of report.docx.',
  )
  expect(await cite($, { source: REPORT, quote: 'not in the file' })).toMatch(/^\[2\] NOT FOUND: these words are not in report\.docx\./)
  // The same words in the same file again are the same citation
  expect(await cite($, { source: REPORT, quote: 'The main finding holds' })).toMatch(/^\[1\] /)
  expect(await cite($, { source: 'C:\\work\\notes.md', quote: 'a third' })).toMatch(/^\[3\] Found word for word, line 3 of notes\.md\./)

  // What went to the script: the kind by the file's extension, and never
  // «show» unless the model asked for it
  expect(runs.map(run => [run.request.kind, run.request.show])).toEqual([
    ['word', false],
    ['word', false],
    ['word', false],
    ['text', false],
  ])
})

test('Excel is cited by cell, row, column or sheet, and a web page is never called checked', async ($, on) => {
  machine(on, request => {
    if (request.kind === 'web') return { found: false }
    if (request.cell === 'Results!B2') return { found: true, text: '95376000,0', formula: '=A1*48000' }
    if (request.cell === 'Results!B3') return { found: false, text: '12' }
    if (request.cell === 'Results!5:5') return { found: true, whole: 'range', text: '71 %', at: 'D5' }

    return { found: true, whole: 'sheet' }
  })

  expect(await cite($, { source: MODEL, cell: 'Results!B2', quote: '95 376 000' })).toBe(
    '[1] Results · B2 of model.xlsx shows «95376000,0» (formula =A1*48000): matches the value cited.',
  )
  expect(await cite($, { source: MODEL, cell: 'Results!B3', quote: '13' })).toMatch(/does NOT match the value cited\.$/)
  expect(await cite($, { source: MODEL, cell: 'Results!5:5', quote: '71 %' })).toBe(
    '[3] Results · row 5 of model.xlsx holds «71 %» in D5.',
  )
  expect(await cite($, { source: MODEL, cell: 'Assumptions!', quote: 'Assumptions' })).toBe(
    '[4] sheet Assumptions of model.xlsx is a sheet of the workbook.',
  )
  expect(await cite($, { source: 'https://example.org/report', quote: 'some words' })).toMatch(/^\[5\] Noted, NOT checked/)
})

test('a source the plugin cannot open is refused with what to give', async ($, on) => {
  const runs = machine(on)

  expect(await cite($, { source: 'C:\\work\\picture.png', quote: 'words' })).toMatch(/^DENIED: Give `source`/)
  expect(await cite($, { source: REPORT, quote: '  ' })).toMatch(/^DENIED: Give `source`/)
  expect(await cite($, { source: MODEL, quote: '12' })).toMatch(/^DENIED: For an Excel file give `cell`/)
  expect(runs.length).toBe(0)
})

test('a lookup that fails is told as unverified, never as found', async ($, on) => {
  machine(on, () => ({ error: 'no such sheet or cell' }))

  expect(await cite($, { source: MODEL, cell: 'Missing!A1', quote: '1' })).toBe(
    '[1] Could not be checked (no such sheet or cell). Treat the claim as unverified and tell the user.',
  )
})

test('the number in an answer is drawn as a link that reads as the place', async ($, on) => {
  const runs = machine(on, request => (request.quote === 'gone' ? { found: false } : { found: true, page: 7 }))

  await cite($, { source: REPORT, quote: 'first' })
  await cite($, { source: REPORT, quote: 'gone' })

  for (const surface of ['terminal', 'desktop'] as const) {
    const ui = await $.ui.mount({
      ...ANSWER,
      surface,
      props: {
        // [9] is no source of ours; «[1]» is talked about, not cited; a
        // paper's own [17–19] and a markdown link are left alone
        text: 'One claim [1]. Another [2]. Not ours [9], nor «[1]», [17–19] or [1](https://example.org).',
        isFirstOfReply: true,
      },
    })
    const text = String((await ui.find({ type: 'Markdown', key: 'answer' }))?.props.text)

    // «p. 7» on an English machine, «s. 7» on a Norwegian one; the number is gone
    expect(text).toMatch(/^One claim \[report, (p|s)\. 7\]\(https:\/\/source\.invalid\/1 "C:\/work\/report\.docx"\)\. /)
    expect(text).toMatch(/Another \[✗ report\.docx\]\(https:\/\/source\.invalid\/2 "C:\/work\/report\.docx"\)\./)
    expect(text).toMatch(/Not ours \[9\], nor «\[1\]», \[17–19\] or \[1\]\(https:\/\/example\.org\)\.$/)

    await ui.press({ key: 'answer', link: { href: 'https://source.invalid/1' } })
    await ui.unmount()
  }

  // A press asks the script to SHOW, the same source, once per surface
  expect(runs.filter(run => run.request.show === true).map(run => run.request.quote)).toEqual(['first', 'first'])
})

test('a link always names its file, a long name by its head and tail, with the path as its title', async ($, on) => {
  machine(on)

  await cite($, { source: 'C:\\work\\Evaluation of the school meals pilot 2025 final v3.docx', quote: 'words' })
  await cite($, { source: 'https://example.org/report', quote: 'some words' })

  const ui = await $.ui.mount({ ...ANSWER, surface: 'desktop', props: { text: 'A [1] and B [2].', isFirstOfReply: true } })
  const text = String((await ui.find({ type: 'Markdown', key: 'answer' }))?.props.text)

  expect(text).toMatch(
    /^A \[Evaluation of the…final v3, (p|s)\. 3\]\(https:\/\/source\.invalid\/1 "C:\/work\/Evaluation of the school meals pilot 2025 final v3\.docx"\) and B \[example\.org\]\(https:\/\/source\.invalid\/2\)\.$/,
  )
  await ui.unmount()
})

test('an answer that rests on several files names the file in each link', async ($, on) => {
  machine(on)

  await cite($, { source: REPORT, quote: 'first' })
  await cite($, { source: MODEL, cell: 'Results!C:C', quote: '12' })

  const ui = await $.ui.mount({
    ...ANSWER,
    surface: 'desktop',
    props: { text: 'A [1] and B [2].', isFirstOfReply: true },
  })
  const text = String((await ui.find({ type: 'Markdown', key: 'answer' }))?.props.text)

  expect(text).toMatch(/^A \[report, (p|s)\. 3\]\(https:\/\/source\.invalid\/1 "[^"]+"\) and B \[model, (column|kolonne) C\]\(https:\/\/source\.invalid\/2 "[^"]+"\)\.$/)
  await ui.unmount()
})

test('an answer that cites nothing of ours is left to the engine', async ($, on) => {
  machine(on)
  on('ui.render', ($, e) => {
    const { Text } = $.ui.resolve(e)

    return <Text>engine</Text>
  })

  const ui = await $.ui.mount({
    ...ANSWER,
    surface: 'desktop',
    props: { text: 'A paper cites [3] and [4].', isFirstOfReply: true },
  })

  expect((await ui.find({ type: 'Text', text: /engine/ }))?.text).toBe('engine')
  expect(await ui.find({ type: 'Markdown', key: 'answer' })).toBe(undefined)
  await ui.unmount()
})

test('the pane lists every source, and sorts and filters them', async ($, on) => {
  const runs = machine(on, request => (request.quote === 'gone' ? { found: false } : { found: true, page: 2 }))

  await cite($, { source: REPORT, quote: 'in the report', claim: 'The report says so' })
  await cite($, { source: MODEL, cell: 'Results!B2', quote: 'gone' })
  await cite($, { source: 'C:\\work\\appendix.docx', quote: 'in the appendix' })

  const ui = await $.ui.mount({ ...PANE, surface: 'desktop' })
  const labels = async (): Promise<string[]> =>
    (await ui.findAll({ type: 'Button' })).map(button => String(button.props.label))

  // In the order cited, the missing one marked
  expect(await labels()).toEqual([expect.stringMatching(/^(p|s)\. 2$/), '✗ Results · B2', expect.stringMatching(/^(p|s)\. 2$/)])
  expect((await ui.find({ type: 'Text', text: /The report says so/ }))?.text).toBe('The report says so')
  expect((await ui.find({ type: 'Text', text: /«in the report»/ }))?.text).toBe('«in the report»')

  await ui.select({ key: 'filter', value: 'missing' })
  expect(await labels()).toEqual(['✗ Results · B2'])

  await ui.select({ key: 'filter', value: 'word' })
  await ui.select({ key: 'sort', value: 'file' })
  // appendix.docx before report.docx
  expect((await ui.findAll({ type: 'Button' })).map(button => button.key)).toEqual(['source-3', 'source-1'])

  // Grouped: a heading per file with its count, files by name, and a row
  // under its file's heading no longer repeats the file's name
  await ui.select({ key: 'filter', value: 'all' })
  await ui.select({ key: 'group', value: 'file' })
  expect((await ui.findAll({ type: 'Text', text: /\(\d\)$/ })).map(heading => heading.text)).toEqual([
    'appendix.docx (1)',
    'model.xlsx (1)',
    'report.docx (1)',
  ])
  expect((await ui.findAll({ type: 'Button' })).map(button => button.key)).toEqual(['source-3', 'source-2', 'source-1'])
  expect(await ui.find({ type: 'Text', text: /^report\.docx$/ })).toBe(undefined)

  await ui.select({ key: 'group', value: 'kind' })
  expect((await ui.findAll({ type: 'Text', text: /\(\d\)$/ })).map(heading => heading.text)).toEqual([
    'Word (2)',
    'Excel (1)',
  ])

  await ui.press({ key: 'source-3' })
  expect(runs.at(-1)?.request).toEqual({
    kind: 'word',
    path: 'C:\\work\\appendix.docx',
    quote: 'in the appendix',
    cell: '',
    show: true,
  })
  await ui.unmount()
})

test('a PDF goes to PDF Scholar: one session, and this plugin’s own numbers', async ($, on) => {
  const runs = machine(on)

  await cite($, { source: REPORT, quote: 'first' })
  expect(await cite($, { source: PAPER, page: 6, quote: 'we use sine and cosine functions', claim: 'Sinusoids' })).toBe(
    '[2] Found word for word, p. 6 of paper.pdf.',
  )
  await cite($, { source: PAPER, page: 7, quote: 'self-attention layers are faster' })

  const reader = runs.filter(run => run.argv[0] !== 'powershell.exe')
  const call = (run: Run): { name: string; arguments: Record<string, unknown> } =>
    (JSON.parse(run.stdin.split('\n')[2] ?? '{}') as { params: { name: string; arguments: Record<string, unknown> } }).params

  // PDF Scholar's own executable in Node mode on its own server script, as
  // the tool «claude», for ONE named session across both processes
  expect(reader.length).toBe(2)
  expect(reader[0]?.argv.slice(0, 4)).toEqual([
    'C:\\Users\\reader\\LOCALAPPDATA\\Programs\\pdfx\\PDF Scholar.exe',
    'C:\\Users\\reader\\APPDATA\\PDF Scholar\\mcp\\pdf-scholar-mcp.cjs',
    '--client',
    'claude',
  ])
  expect(reader[0]?.argv[4]).toBe('--session')
  expect(reader[0]?.argv[5]).toMatch(/^[0-9a-f]{16}$/)
  expect(reader[1]?.argv[5]).toBe(reader[0]?.argv[5])

  // The citation's number here is the card's number there
  expect(call(reader[0] as Run)).toEqual({
    name: 'show_passage',
    arguments: { path: PAPER, page: 6, quote: 'we use sine and cosine functions', claim: 'Sinusoids', number: 2, show: false },
  })
  expect(call(reader[1] as Run).arguments.number).toBe(3)
})

test('on a system that is not Windows the plugin offers nothing and starts nothing', async ($, on) => {
  const runs = machine(on, undefined, 'Darwin')

  registered.length = 0
  await $.session.start({ cwd: '/', surface: 'terminal', isInteractive: true })

  // No tool for the model to call and be let down by
  expect(registered).toEqual([])
  // …and one that is called all the same says why, and touches no program
  expect(await cite($, { source: REPORT, quote: 'words' })).toMatch(/^DENIED: Show your work runs on Windows only for now/)
  expect(runs.length).toBe(0)
})

const BAND = {
  plugin: 'show-your-work',
  component: 'AbovePrompt',
  props: {
    hasSurvey: false,
    isCollapsed: false,
    isWorking: false,
    maxRows: 4,
    bodyColumns: 80,
    scroll: { offset: 0, bodyRows: 4 },
    view: {},
  },
} as const

/** The reader's choice of when Claude cites, as the store across sessions holds it */
function chosen(on: On, mode: () => string | undefined, submitted: string[] = []): void {
  on('store.get', () => ({ value: mode() }))
  on('store.set', () => ({ value: undefined }))
  on('prompt.compose', () => ({ sections: [{ id: 'core', text: 'You are Claude.', scope: 'shared' as const }] }))
  on('prompt.submit', (_, e) => {
    submitted.push(e.text)

    return { text: e.text, context: [] } as never
  })
  on('tool.call', () => ({ result: 'ok' }))
}

test('citing is the default: the tool says when to call it, unless the reader turned it down', async ($, on) => {
  let mode: string | undefined
  machine(on)
  chosen(on, () => mode)

  const start = async (): Promise<string | undefined> => {
    registered.length = 0
    described.length = 0
    await $.session.start({ cwd: '/', surface: 'desktop', isInteractive: true })

    return described[0]
  }

  // Nothing chosen yet: Claude cites by itself
  expect(await start()).toMatch(/^USE THIS WHENEVER you tell the user what a document says/)
  mode = 'asked'
  expect(await start()).toMatch(/^Use this when the user asks where something is said/)
  // Off: no tool at all, and one called all the same refuses and starts nothing
  mode = 'off'
  expect(await start()).toBe(undefined)
  expect(await cite($, { source: REPORT, quote: 'words' })).toMatch(/^DENIED: The user has turned Show your work off/)
})

test('a turn that read documents and cited none is offered its sources in one press', async ($, on) => {
  const asked: string[] = []
  machine(on)
  chosen(on, () => 'asked', asked)

  const offer = async (): Promise<string | undefined> => {
    const ui = await $.ui.mount({ ...BAND, surface: 'desktop' })
    const text = (await ui.find({ type: 'Text', text: /report\.docx/ }))?.text

    await ui.unmount()

    return text
  }

  // Code and notes are read all day and offered nothing
  await $.tool.call({ tool: 'Read', file_path: 'C:\\work\\main.ts' } as never)
  expect(await offer()).toBe(undefined)

  await $.tool.call({ tool: 'Read', file_path: REPORT } as never)
  await $.tool.call({ tool: 'Bash', command: 'python read.py "C:\\work\\model.xlsx"' } as never)
  expect(await offer()).toMatch(/report\.docx, model\.xlsx/)

  // One press asks for the sources, as the reader's own words
  const ui = await $.ui.mount({ ...BAND, surface: 'desktop' })
  await ui.press({ key: 'cite' })
  await ui.unmount()
  expect(asked.at(-1)).toMatch(/cite them with show_source/)
  expect(await offer()).toBe(undefined)

  // And a turn that cites is offered nothing
  await $.tool.call({ tool: 'Read', file_path: REPORT } as never)
  await cite($, { source: REPORT, quote: 'words' })
  expect(await offer()).toBe(undefined)
})

test('a source nobody noted is read from the transcript: the link is drawn all the same', async ($, on) => {
  // The first session with the plugin installed (2026-11-25): eight sources
  // cited and checked, nothing in the session's state when the answer was
  // drawn, and the answer shown with its bare numbers.
  machine(on)
  transcript = [
    {
      role: 'assistant',
      text: '',
      toolUses: [
        { tool_use_id: 'a', tool: 'Bash', input: { command: 'ls' }, text: '[1] not ours' },
        {
          tool_use_id: 'b',
          tool: TOOL,
          input: { source: REPORT, quote: 'Attendance rose by 4.2 percentage points' },
          text: '[1] Found word for word, p. 3 of report.docx.',
        },
        {
          tool_use_id: 'c',
          tool: TOOL,
          input: { source: MODEL, cell: 'Results!B2', quote: '12 000 000' },
          text: '[2] Results · B2 of model.xlsx shows «9537600»: does NOT match the value cited.',
        },
      ],
    },
  ]

  const ui = await $.ui.mount({
    ...ANSWER,
    surface: 'desktop',
    props: { text: 'Attendance rose [1]. It cost twelve million [2].', isFirstOfReply: true },
  })
  const text = String((await ui.find({ type: 'Markdown', key: 'answer' }))?.props.text)

  expect(text).toMatch(/^Attendance rose \[report, (p|s)\. 3\]\(https:\/\/source\.invalid\/1 "[^"]+"\)\. /)
  expect(text).toMatch(/It cost twelve million \[✗ model, B2\]\(https:\/\/source\.invalid\/2 "[^"]+"\)\.$/)
  await ui.unmount()

  // The pane lists them too
  const pane = await $.ui.mount({ ...PANE, surface: 'desktop' })
  expect((await pane.findAll({ type: 'Button' })).map(button => button.key)).toEqual(['source-1', 'source-2'])
  await pane.unmount()
  transcript = []
})
