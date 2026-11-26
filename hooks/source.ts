// A source: what kind it is, what it is called, where in it the words are,
// and how the model's call to the tool becomes one.
import type { Kind, Source, System } from '../types'
import type { Words } from './words'

/** The tool, as the engine names it in a call and in the transcript */
export const TOOL = 'mcp__show-your-work__show_source'

export const KINDS: readonly Kind[] = ['word', 'powerpoint', 'excel', 'pdf', 'text', 'html', 'web']

/** What can be shown on a system. Windows and macOS: every kind, each by the
 *  system's own script. Linux: PDFs alone, which PDF Scholar shows there as
 *  everywhere — Office has no Linux, and nothing is offered half-working. */
export const kindsOn = (system: System): readonly Kind[] =>
  system === 'windows' || system === 'mac' ? KINDS : system === 'linux' ? ['pdf'] : []

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

export function kindOf(path: string): Kind | undefined {
  if (/^https?:\/\//i.test(path)) {
    return 'web'
  }

  const extension = /\.([a-z0-9]+)$/i.exec(path)?.[1]?.toLowerCase() ?? ''

  return (Object.keys(EXTENSIONS) as Exclude<Kind, 'web'>[]).find(kind =>
    EXTENSIONS[kind].includes(extension),
  )
}

export const nameOf = (path: string): string => path.split(/[\\/]/).pop() ?? path

export const hostOf = (url: string): string => /^https?:\/\/([^/]+)/i.exec(url)?.[1] ?? url

/** What a source is called where it is named: a file by its name, a web page
 *  by its host */
export const calledOf = (source: Source): string =>
  source.kind === 'web' ? hostOf(source.path) : nameOf(source.path)

/** An Excel place in words: «AN250», «row 250», «column C–E», «sheet
 *  Assumptions»; with the sheet in front when asked, «Results · row 5» */
export function excelPlace(cell: string, isWithSheet: boolean, the: Words): string {
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
export function placeOf(source: Source, the: Words): string {
  if (source.kind === 'excel') return excelPlace(source.cell, true, the)
  if (source.kind === 'web') return hostOf(source.path)
  if (source.kind === 'html' || source.page <= 0) return nameOf(source.path)
  if (source.kind === 'powerpoint') return the.slide(source.page)
  if (source.kind === 'text') return the.line(source.page)

  return the.page(source.page)
}

/** A long name by its head and its tail, where names that begin alike differ */
const shortened = (name: string): string =>
  name.length > 30 ? `${name.slice(0, 18).trimEnd()}…${name.slice(-9).trimStart()}` : name

/** What stands IN THE SENTENCE where the model wrote «[3]»: the file and the
 *  place, as a paper cites it. The file is ALWAYS named: «p. 17» alone says
 *  nothing in a session that reads many documents (Emil, 2026-11-25 — an
 *  answer about four countries' reports, every link a bare page). A web page
 *  is its host already. */
export function inline(source: Source, the: Words): string {
  const place = source.kind === 'excel' ? excelPlace(source.cell, false, the) : placeOf(source, the)
  const name = nameOf(source.path)
  const text =
    source.kind === 'web' || place === name
      ? shortened(place)
      : `${shortened(name.replace(/\.[a-z0-9]+$/i, ''))}, ${place}`

  // Nothing of a file's name may end the link early or start a new one
  return `${source.status === 'missing' ? '✗ ' : ''}${text}`.replace(/[[\]\\]/g, '')
}

/** The link's title, shown on hover: the whole path, which the link's text
 *  has no room for. No quote or line break may end the title early. */
export const titled = (source: Source): string =>
  source.kind === 'web' ? '' : ` "${source.path.replace(/\\/g, '/').replace(/["\r\n]/g, '')}"`

/** A source as the model asked for it, before it has a number or has been
 *  looked up; undefined when the call names no source of a known kind, or no
 *  words. The call's input is the model's, so every field is checked. */
export function askedOf(input: Record<string, unknown>): Source | undefined {
  const text = (value: unknown): string => (typeof value === 'string' ? value.trim() : '')
  const path = text(input.source)
  const quote = text(input.quote)
  const kind = kindOf(path)

  if (kind === undefined || quote === '') {
    return undefined
  }

  return {
    n: 0,
    kind,
    path,
    quote,
    cell: text(input.cell),
    page: typeof input.page === 'number' && input.page >= 1 ? Math.floor(input.page) : 0,
    claim: text(input.claim),
    status: 'unchecked',
  }
}
