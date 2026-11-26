# Show your work — a Claude Code plugin

Claude cites a source through one tool, the plugin looks the words up in the
file, and the citation becomes a link in the answer that opens the source at
the words. Owner: Emil (communicates in Norwegian; code, comments and commit
messages in English). User-facing strings in Norwegian and English, chosen by
the machine's language (`words()` in `hooks/words.ts`).

Sister project: [PDF Scholar](https://github.com/emilmsh/pdf-scholar), which
shows the PDFs. Its checkout is beside this one.

## Layout

- `hooks/register.tsx` — the hooks: the tool, the numbering, the link in the
  answer, the offer above the prompt. The engine loads the file named in
  `hooks/hooks.json` and follows its imports; nothing is built or bundled.
- Beside it: `words.ts` (the reader's language), `source.ts` (a source's
  kind, name and place, and the model's call read into one), `state.ts`
  (every value kept, and the mode), `locate.ts` (looking a source up and
  showing it — the Windows seam), `tell.ts` (what the model is told, and
  reading it back out of the transcript), `pane.tsx` (the Sources pane).
- `scripts/show.ps1` — everything that touches a program on Windows.
  `scripts/show.js` — the same on macOS, behind the same request and answer.
- `types/index.d.ts` — the contract for what the plugin keeps in session
  state. Change a value's fields and you change its `shape` tag too.
- `hooks/register.test.tsx` — the tests. They start no program.
- `.claude-plugin/` — the manifest, the marketplace file (this repository is
  its own marketplace) and the icon.

## Commands

Use Claude Code 2.1.330 or later (`claude` on the PATH may be older).

- `claude plugin validate .`
- `claude plugin test .`
- Type-check with the `tsconfig.json` here, once the engine has laid its types
  in `.claude-plugin/types/` (it does when the plugin has been loaded once).
- The look-up itself (`scripts/show.ps1`), against the real programs, nothing
  shown on screen. Run both after any change to the script:
  - `scripts/test/make-word.ps1 <out.docx>` then `scripts/test/check-word.ps1
    <out.docx>`: a Word report with a table of contents, fields, footnotes, a
    table, links, a tracked deletion and hidden text. Each marked passage must
    come back on the page Word's own search puts it on, with the words at the
    place being the words asked for (`exact`). The demo's `report.docx` has
    none of those and cannot catch a wrong place: on 2026-11-26 a real report
    answered page 19 for words on page 25.
  - `scripts/test/check-kinds.ps1`: PowerPoint, Excel, text and code from
    `demo/`, with words that are there and words that are not.
  - `python scripts/test/make_hard.py <folder>` then
    `scripts/test/check-cases.ps1 <folder>` (`-Kind excel` for one kind): 83
    cases on files with what real ones have - grouped shapes, table rows,
    soft line breaks; formatted numbers in every way they are written,
    merged cells, sheet names with spaces and apostrophes; text in three
    encodings, Markdown quoted without its marks; HTML entities, inline
    tags, comments. A third of the cases are words that are NOT there. The
    first run found 24 wrong answers. Add a case here for every look-up that
    goes wrong in use.
  - A PDF is looked up by PDF Scholar: its own `npm run test:passage`.
- The Mac's look-up (`scripts/show.js`), as far as it goes without a Mac:
  `node scripts/test/check-cases.mjs <folder> [word.docx]` runs the same
  `cases.json` (text, HTML, PowerPoint) and the Word fixture's passages
  through the half of the script that reads files. Excel, Word's page and
  everything shown need a Mac: `sh scripts/test/check-mac.sh [show]`.

## Platforms: one seam, and nothing system-shaped outside it

**Windows is the reference.** macOS is held to it kind for kind; Linux shows
PDFs only. What each system shows, and where it differs, is the table in the
README (§ Which systems): a difference that is not in that table is a bug.
The macOS script was written on Windows (2026-11-26) and is **not released
until it has run on a Mac** (`docs/MAC-TEST.md`); the branch is `mac-linux`.

What a system decides lives in exactly three places:

1. `scripts/show.ps1` (Windows) — Office through COM, Notepad through UI
   Automation, VS Code and the browser by their programs, window raising
   through Win32.
2. `scripts/show.js` (macOS, run by `osascript -l JavaScript`) — Word, Excel
   and PowerPoint through Apple events, VS Code by its `vscode://` address,
   the browser through `open location`. A Word or PowerPoint quote is looked
   up in the FILE (a zip of XML); the program is asked only for the page and
   to select. Its upper half touches no program and runs under Node too.
3. `hooks/locate.ts` — which script a system runs (`SCRIPT`) and where PDF
   Scholar is (`pdfScholar`: `command.json` beside its server script first,
   then the usual install places).

Which kinds a system shows is `kindsOn` in `hooks/source.ts`, and the system
itself `systemOf` in `hooks/state.ts`. Everything else in `hooks/` — the
tool's contract, the numbering, the labels, the link in the answer, the pane,
the state — must stay free of anything a system decides. The rules:

- **A new source kind is a contract first, a system's way second.** Add the
  kind, its label and its answer in `source.ts`, `words.ts`, `tell.ts` and the
  tests, and only then the way to look it up and show it — in BOTH scripts, in
  the same change. If the look-up could be done without a program (reading a
  text file), do it without one.
- **The script's request and answer are the seam.** `{ kind, path, quote,
  cell, show }` in, one JSON line `{ found, page, … }` or `{ error }` out,
  the same fields on every system. A system gets its own script behind it;
  it never gets a branch inside another's.
- **A rule for judging a quote is changed in both scripts at once**: the
  pattern (`Pattern` / `pattern`), the readings of a number (`Number` /
  `readings`), the Markdown and HTML reading. `check-cases.ps1` and
  `check-cases.mjs` run the SAME `cases.json` against each, so a rule that
  moved in one alone fails there.
- **No path is split, joined or guessed in `hooks/`** beyond taking a
  file's name for a label, and PDF Scholar's own folders in `pdfScholar`. A
  path is the model's, passed through.
- **On a system with no way of its own the plugin offers nothing**
  (`kindsOn` is empty): no tool, so the model is never let down by one. On
  Linux the tool is worded for PDFs alone. Never ship a half-working kind.
- **PDF Scholar must work the same without this plugin.** Nothing in PDF
  Scholar may depend on it, and what the plugin asks of PDF Scholar goes
  through its public MCP tools — the same on every system PDF Scholar runs on.

## Releasing

Users receive a change only when `version` in `.claude-plugin/plugin.json`
changes: a new commit alone reaches nobody who has the plugin installed.

1. `claude plugin validate .` and `claude plugin test .`.
2. Press a link for one file of each kind on a Windows machine with Office,
   VS Code and PDF Scholar — once with the file open, once with the program
   closed. The tests stand in for the script; only a hand can check it.
3. Bump `version` in `.claude-plugin/plugin.json` (nowhere else: the
   marketplace entry carries none, on purpose). Commit as
   `vX.Y.Z: <one line>`, tag `vX.Y.Z`, push with the tag.

Auto-update is OFF by default for a third-party marketplace; the README tells
the reader how to turn it on. Do not promise that updates arrive by
themselves.

## Gotchas

- `scripts/show.ps1` must be plain ASCII: Windows PowerShell reads a script
  without a byte-order mark as ANSI. A no-break space in a pattern is written
  ` `.
- `scripts/show.js` is kept ASCII too. A tool that writes a file can turn an
  escaped character in a pattern into the character itself without saying so
  (2026-11-26: every no-break space and byte-order mark in the new script):
  search the file for anything outside ASCII after writing it.
- The script answers with ONE line on standard output. A method that returns
  a value (`Select()`, `Activate()`) must be cast to `[void]`, or its `True`
  becomes part of the answer.
- A document can be open with no window on screen (a check that was stopped
  half way). «Shown» must make the window visible, not only activate it.
- Excel's `Intersect` cannot be called with two arguments on an instance the
  script started itself; `GetActiveObject` can hand back an invisible Excel.
- Start a program by its own executable, never through a wrapper that waits
  (`code.cmd` held the script for 42 seconds).
- A press may hold for ten seconds. Start the work and return.
- A plugin must not be named like an MCP server whose tools it calls: both
  would own `mcp__<name>__…`.
- Across files the engine follows `$` only into a FUNCTION DECLARATION
  (`export function locate($ …)`), never into a `const f = ($) => …`; and
  `read`/`update` must name an atom defined in the same file. That is why
  every read and write of a kept value lives in `state.ts`.
