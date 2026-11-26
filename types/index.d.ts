/** The kind of source, which is also which program shows it */
export type Kind = 'word' | 'powerpoint' | 'excel' | 'pdf' | 'text' | 'html' | 'web'

/** What looking the words up in the source came to. A web page is never
 *  fetched, so its words are never checked. */
export type Status = 'found' | 'missing' | 'unchecked'

/** One cited source: what the model pointed at, and where it was found */
export type Source = {
  /** The citation's number in this session, from 1; never given twice */
  n: number
  kind: Kind
  /** An absolute file path, or a web address */
  path: string
  /** The words cited, verbatim; for Excel, the value cited */
  quote: string
  /** Excel only: «Sheet!B2», «Sheet!5:5», «Sheet!C:C», «Sheet!» */
  cell: string
  /** A page (Word, PDF), a slide (PowerPoint) or a line (a text or code
   *  file); 0 where there is none */
  page: number
  /** The claim the words back, as the model put it */
  claim: string
  status: Status
}

/** How the reader has the Sources pane arranged */
export type View = {
  sort: 'cited' | 'file' | 'kind'
  filter: 'all' | 'missing' | Kind
  /** Under a heading per file or per kind, or as one list */
  group: 'none' | 'file' | 'kind'
}

/** The system the plugin runs on; 'other' is one it can show nothing on */
export type System = 'windows' | 'mac' | 'linux' | 'other'

/** How eagerly Claude cites: by itself, only when asked, or not at all */
export type Mode = 'always' | 'asked' | 'off'

declare module 'claude-code' {
  interface PluginState {
    'show-your-work': {
      sources: Shaped<Source[]>
      /** The id PDF Scholar keeps this session's references under: sixteen
       *  hex characters, made at the first PDF citation; '' until then */
      reader: Shaped<string>
      view: Shaped<View>
      /** The documents this turn has read and not cited: file names, in the
       *  order read. Emptied when a turn starts and when a source is cited. */
      uncited: Shaped<string[]>
    }
  }
}
