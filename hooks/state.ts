// What the plugin keeps, and what it asks of the machine. Every read and
// write of the session's values is in this file: the engine lists the values
// a module touches from the atoms in the file that touches them.
import { atom, read, update } from 'claude-code'
import type { EngineInterface } from 'claude-code'

import type { Mode, Source, View } from '../types'

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

/** The sources the tool has noted in this session, in the order numbered */
export async function noted($: EngineInterface): Promise<readonly Source[]> {
  return read($, sources)
}

/** Notes a citation and gives it its number. The number comes FIRST, before
 *  the look-up: a source in a program with a list of its own (PDF Scholar) is
 *  told it, so its card and this list read the same. The same words in the
 *  same source again are the same citation, and a number is never given
 *  twice — the next is one past the highest ever given. */
export async function numbered($: EngineInterface, asked: Source): Promise<Source> {
  let made = asked

  await update($, sources, list => {
    const same = list.find(one => one.path === asked.path && one.quote === asked.quote && one.cell === asked.cell)

    made = { ...asked, n: same?.n ?? list.reduce((most, one) => Math.max(most, one.n), 0) + 1 }

    return same === undefined ? [...list, made].slice(-MOST) : list
  })

  return made
}

/** Puts what the look-up found in place of the citation with its number */
export async function renoted($: EngineInterface, made: Source): Promise<void> {
  await update($, sources, list => list.map(one => (one.n === made.n ? made : one)))
}

/** This session's id in PDF Scholar's list: made at the first PDF citation */
export async function readerSession($: EngineInterface): Promise<string> {
  let session = ''

  await update($, reader, id => {
    session =
      id !== '' ? id : Array.from({ length: 16 }, () => Math.floor(Math.random() * 16).toString(16)).join('')

    return session
  })

  return session
}

/** How the reader has the pane arranged */
export async function arrangement($: EngineInterface): Promise<View> {
  return read($, view)
}

export async function arrange($: EngineInterface, change: Partial<View>): Promise<void> {
  await update($, view, now => ({ ...now, ...change }))
}

/** The documents this turn read and has not cited */
export async function unshown($: EngineInterface): Promise<readonly string[]> {
  return read($, uncited)
}

/** Adds documents a turn read; none (the default) empties the list */
export async function reading($: EngineInterface, names: readonly string[] = []): Promise<void> {
  await update($, uncited, list => (names.length === 0 ? [] : [...new Set([...list, ...names])].slice(-20)))
}

export const MODES: readonly Mode[] = ['always', 'asked', 'off']

export const isMode = (value: unknown): value is Mode => (MODES as readonly unknown[]).includes(value)

/** The reader's choice, kept across sessions; 'always' until they choose */
export async function modeOf($: EngineInterface): Promise<Mode> {
  try {
    const kept = await $.store.get('mode')

    return isMode(kept) ? kept : 'always'
  } catch {
    return 'always'
  }
}

/** Everything this plugin does to a program is Windows' (scripts/show.ps1,
 *  PDF Scholar's install path). On another system it offers NOTHING — no tool
 *  for the model to call and be let down by — until that system has a way of
 *  its own. A machine that cannot be asked is taken for Windows: the script
 *  then says what went wrong by itself. */
export async function isWindows($: EngineInterface): Promise<boolean> {
  try {
    return (await $.session.machine()).sysname === 'Windows_NT'
  } catch {
    return true
  }
}
