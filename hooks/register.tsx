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
// Windows and macOS, each by a script of its own behind one request and one
// answer (scripts/show.ps1: Office through COM, Notepad through UI
// Automation; scripts/show.js: Office through Apple events). PDFs go through
// PDF Scholar (github.com/emilmsh/pdf-scholar), whose own server script the
// plugin runs — so its switches decide what is shared, here as everywhere —
// and on Linux a PDF is the one kind shown (`kindsOn`).
//
// This module holds the hooks; the rest is beside it:
//   words.ts   the reader's language       source.ts  a source's kind, name and place
//   state.ts   what is kept, the mode      locate.ts  looking a source up, and showing it
//   tell.ts    what the model is told, and reading it back from the transcript
//   pane.tsx   the Sources pane
import type { Register } from 'claude-code'

import type { Source } from '../types'
import { locate, show } from './locate'
import { PANE, pane } from './pane'
import { TOOL, askedOf, inline, kindOf, kindsOn, nameOf, titled } from './source'
import { modeOf, numbered, reading, renoted, systemOf, unshown } from './state'
import { ELSEWHERE, PDF_ONLY, answerOf, known, toolFor } from './tell'
import { words } from './words'

/** «[3]» in an answer — not «[3, 5]» or «[17–19]», a paper's own citations
 *  quoted along; not a markdown link's «[3](…)»; and not a «[3]» the answer
 *  only talks ABOUT, in quotes or as code */
const CITED = /(?<![«"“`])[ \t]*\[(\d{1,6})\](?![(»"”`])/g

/** A citation's link target inside an answer. Nothing is ever fetched from
 *  it: the press is the plugin's own, and the number says which source was
 *  pressed. (Markdown draws only http, https and file links.) */
const HREF = 'https://source.invalid/'

/** The kinds a turn is offered sources for: documents, never the code and
 *  notes a coding session reads all day */
const DOCUMENT = /[^\s"'<>|*?]+\.(docx?|pptx?|xlsx?|pdf)\b/gi

export const register: Register = on => {
  on('session.start', async ($, e, next) => {
    const mode = await modeOf($)

    const system = await systemOf($)

    // Off, or on a system nothing can be shown on: no tool for the model to
    // reach for
    if (mode === 'off' || kindsOn(system).length === 0) {
      return next(e)
    }

    await $.tool.register({ name: 'show_source', ...toolFor(mode, system) })
    await $.command.register({ name: 'sources', description: words().command })

    return next(e)
  })

  // A new turn starts clean: what the last one read is no longer on offer
  on('prompt.submit', async ($, e, next) => {
    await reading($)

    return next(e)
  })

  // Every tool call is looked at for the documents it names — a Read, a
  // script that opens one — and nothing else of it is kept. Only documents
  // this system can show: sources are not offered for the others.
  on('tool.call', async ($, e, next) => {
    if (e.tool !== TOOL) {
      const found = [...JSON.stringify(e).matchAll(DOCUMENT)].map(match => nameOf(match[0].replace(/\\\\/g, '/')))
      const shown = found.length === 0 ? [] : kindsOn(await systemOf($))
      const named = found.filter(name => shown.includes(kindOf(name) ?? 'web'))

      if (named.length > 0) {
        await reading($, named)
      }
    }

    return next(e)
  })

  on('ui.render', { component: 'AbovePrompt' }, async ($, e, next) => {
    const read_ = await unshown($)

    if (e.props.hasSurvey || read_.length === 0 || (await modeOf($)) === 'off') {
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
            await reading($)
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
    const shown = kindsOn(await systemOf($))

    if (shown.length === 0) {
      return { deny: ELSEWHERE }
    }

    if ((await modeOf($)) === 'off') {
      return { deny: 'The user has turned Show your work off. Cite in plain text, and say that the links are off.' }
    }

    const input = e as unknown as Record<string, unknown>
    const asked = askedOf(input)

    if (asked === undefined) {
      return {
        deny: 'Give `source` (a Word, PowerPoint, Excel, PDF, HTML, text or code file by absolute path, or an https URL) and `quote`.',
      }
    }

    if (!shown.includes(asked.kind)) {
      return { deny: PDF_ONLY }
    }

    if (asked.kind === 'excel' && typeof input.cell !== 'string') {
      return { deny: 'For an Excel file give `cell`, as «Sheet!B2».' }
    }

    const citation = await numbered($, asked)
    const located = await locate(citation, input.show === true, $)

    // Cited: the turn is no longer one that read documents and showed nothing
    await reading($)

    const made = { ...citation, page: located.page || asked.page, status: located.status }

    await renoted($, made)
    // Nowhere to draw a pane (a phone, a run with no window) is no failure
    await $.ui.open({ id: PANE, title: words().pane })

    return { result: answerOf(made, located) }
  })

  on('ui.render', { component: 'Pane', requestId: PANE }, ($, e) => pane($, e))

  on('ui.render', { component: 'AssistantMessage' }, async ($, e, next) => {
    const cited = new Set([...e.props.text.matchAll(CITED)].map(match => Number(match[1])))

    if (cited.size === 0) {
      return next(e)
    }

    const mine = new Map(
      (await known($)).filter(one => cited.has(one.n)).map((one): [number, Source] => [one.n, one]),
    )

    if (mine.size === 0) {
      return next(e)
    }

    const { Markdown } = $.ui.resolve(e)
    const the = words()
    // «[3]» in the sentence becomes a link that reads as the place; a number
    // that is no source of ours stays as written
    const text = e.props.text.replace(CITED, (whole, digits: string) => {
      const source = mine.get(Number(digits))

      return source === undefined ? whole : ` [${inline(source, the)}](${HREF}${source.n}${titled(source)})`
    })

    // The answer is drawn here, not by `next`: only a Markdown of the
    // plugin's own hands it the press on a link
    return (
      <Markdown
        key="answer"
        text={text}
        pressableLinks={[...mine.keys()].map(n => `${HREF}${n}`)}
        onLinkPress={link => {
          const source = mine.get(Number(link.href.slice(HREF.length)))

          if (source !== undefined && link.href.startsWith(HREF)) {
            show($, source)
          }
        }}
      />
    )
  })
}
