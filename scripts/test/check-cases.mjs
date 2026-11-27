// Holds scripts/show.js - the look-up macOS uses - to the cases the Windows
// script is held to (check-cases.ps1), for every kind it judges without a
// program: text, HTML, PowerPoint, and whether a Word file holds the words.
// Runs under Node on any system; nothing is shown on screen.
//
//   node scripts/test/check-cases.mjs <folder made by make_hard.py> [word.docx]
//
// With a Word file made by make-word.ps1 (its passages are in <file>.json):
// every passage must be found in it. The PAGE is Word's to say, on a Mac.
// Exit code = cases that failed.
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const [folder, wordFile] = process.argv.slice(2)
const show = join(dirname(dirname(fileURLToPath(import.meta.url))), 'show.js')
// On a Mac, SYW_OSASCRIPT=1 runs the script the way the plugin does there
// (JavaScript for Automation) rather than under Node: the same cases, read by
// the Mac's own means
const runner = process.env.SYW_OSASCRIPT === '1' ? ['/usr/bin/osascript', '-l', 'JavaScript', show] : [process.execPath, show]
const ask = request =>
  JSON.parse(
    execFileSync(runner[0], runner.slice(1), {
      env: { ...process.env, SYW_REQUEST: JSON.stringify({ ...request, show: false }) },
      encoding: 'utf8',
    })
      .trim()
      .split(/\r?\n/)
      .pop(),
  )
const json = path => JSON.parse(readFileSync(path, 'utf8').replace(/^﻿/, ''))
let failed = 0
let ran = 0

for (const one of json(join(folder, 'cases.json'))) {
  if (one.kind === 'excel') {
    continue
  }

  ran++

  const answer = ask({ kind: one.kind, path: join(folder, one.file), quote: one.quote, cell: one.cell })
  const found = answer.found === true

  if (found === one.found && (one.page === 0 || answer.page === one.page)) {
    continue
  }

  failed++
  console.log(`FAIL  ${one.file}  [${one.why}]`)
  console.log(`      quote: ${one.quote}`)
  console.log(`      found=${found} (want ${one.found})  place=${answer.page} (want ${one.page})  ${answer.error ?? ''}`)
}

if (wordFile) {
  for (const one of json(`${wordFile}.json`).marks) {
    ran++

    const answer = ask({ kind: 'word', path: wordFile, quote: one.quote })

    if (answer.found === true) {
      continue
    }

    failed++
    console.log(`FAIL  ${wordFile}  quote: ${one.quote}  ${answer.error ?? ''}`)
  }
}

console.log(`${failed} of ${ran} failed`)
process.exit(failed)
