// The reader's language: every word the plugin shows, in English and
// Norwegian. The model is always answered in English (see tell.ts).
import type { Kind, Mode, View } from '../types'

export type Words = {
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

export const EN: Words = {
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
export function words(): Words {
  let language = 'en'

  try {
    language = new Intl.DateTimeFormat().resolvedOptions().locale.toLowerCase().split('-')[0] ?? 'en'
  } catch {
    // No Intl here: English
  }

  return ['nb', 'nn', 'no'].includes(language) ? NB : EN
}
