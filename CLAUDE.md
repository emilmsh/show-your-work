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

## Platforms: one seam, and nothing Windows-shaped outside it

The plugin is **Windows only for now**, and that is said wherever it is
offered (README, plugin description, PDF Scholar's documents).

What is Windows' own lives in exactly two places:

1. `scripts/show.ps1` — Office through COM, Notepad through UI Automation,
   VS Code and the browser by their programs, window raising through Win32.
2. `byPdfScholar` in `hooks/locate.ts` — PDF Scholar's install path
   (`%LOCALAPPDATA%\Programs\pdfx`) and its data folder.

Everything else in `hooks/` — the tool's contract, the numbering, the
kinds, the labels, the link in the answer, the pane, the state — must stay
free of anything a system decides. The rules that keep it so:

- **A new source kind is a contract first, a Windows way second.** Add the
  kind, its label and its answer in `source.ts`, `words.ts`, `tell.ts` and the tests, and only
  then the way to look it up and show it in `show.ps1`. If the look-up could
  be done without a program (reading a text file), do it without one.
- **The script's request and answer are the seam.** `{ kind, path, quote,
  cell, show }` in, one JSON line `{ found, page, … }` or `{ error }` out.
  Another system gets its own script behind the same request and answer; it
  never gets a branch inside the Windows script.
- **No path is split, joined or guessed in `hooks/`** beyond taking a
  file's name for a label. A path is the model's, passed through.
- **On another system the plugin offers nothing** (`isWindows`): no tool, so
  the model is never let down by one. Keep that until the system has a way of
  its own; never ship a half-working one.
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
