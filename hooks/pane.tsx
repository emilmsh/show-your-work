// The Sources pane: every source of the session, to arrange and press.
import type { EngineInterface, RenderInput } from 'claude-code'

import type { Kind, Source, View } from '../types'
import { show } from './locate'
import { KINDS, calledOf, nameOf, placeOf } from './source'
import { MODES, arrange, arrangement, isMode, modeOf } from './state'
import { known } from './tell'
import type { Words } from './words'
import { words } from './words'

export const PANE = 'sources'

const SORTS: readonly View['sort'][] = ['cited', 'file', 'kind']

const GROUPS: readonly View['group'][] = ['none', 'file', 'kind']

const isSort = (value: string): value is View['sort'] => (SORTS as readonly string[]).includes(value)

const isGroup = (value: string): value is View['group'] => (GROUPS as readonly string[]).includes(value)

const isFilter = (value: string): value is View['filter'] =>
  value === 'all' || value === 'missing' || (KINDS as readonly string[]).includes(value)

/** The sources the pane lists, as the reader has it arranged */
function arranged(list: readonly Source[], { sort, filter }: View): Source[] {
  const kept = list.filter(one =>
    filter === 'all' ? true : filter === 'missing' ? one.status === 'missing' : one.kind === filter,
  )
  const byName = (a: Source, b: Source): number =>
    nameOf(a.path).localeCompare(nameOf(b.path), undefined, { numeric: true }) || a.n - b.n

  if (sort === 'file') {
    return [...kept].sort(byName)
  }

  if (sort === 'kind') {
    return [...kept].sort((a, b) => KINDS.indexOf(a.kind) - KINDS.indexOf(b.kind) || byName(a, b))
  }

  return [...kept].sort((a, b) => a.n - b.n)
}

type Group = { key: string; title: string; items: Source[] }

/** The arranged sources under a heading per file or per kind — files by
 *  name, kinds in their fixed order; each group keeps the list's own order.
 *  Ungrouped, one group with no heading. */
function grouped(shown: readonly Source[], by: View['group'], the: Words): Group[] {
  if (by === 'none') {
    return [{ key: 'all', title: '', items: [...shown] }]
  }

  const groups = new Map<string, Group>()

  for (const source of shown) {
    const key = by === 'kind' ? source.kind : source.path
    const group = groups.get(key) ?? { key, title: by === 'kind' ? the.kinds[source.kind] : calledOf(source), items: [] }

    group.items.push(source)
    groups.set(key, group)
  }

  return [...groups.values()].sort((a, b) =>
    by === 'kind'
      ? KINDS.indexOf(a.key as Kind) - KINDS.indexOf(b.key as Kind)
      : a.title.localeCompare(b.title, undefined, { numeric: true }),
  )
}

/** The pane itself */
export async function pane($: EngineInterface, e: RenderInput<'Pane'>) {
  const { Box, Text, Button, Select } = $.ui.resolve(e)
  const the = words()
  const list = await known($)
  const how = await arrangement($)
  const shown = arranged(list, how)
  const kinds = KINDS.filter(kind => list.some(one => one.kind === kind))
  const lost = list.filter(one => one.status === 'missing').length
  const modes = (
    <Select
      key="mode"
      label={the.mode}
      value={await modeOf($)}
      options={MODES.map(value => ({ value, label: the.modes[value] }))}
      onSelect={value => {
        if (isMode(value)) {
          void $.store.set('mode', value).then(() => {
            $.ui.toast(the.nextSession)
            $.ui.invalidate('ui.render')
          })
        }
      }}
    />
  )

  if (list.length === 0) {
    return (
      <Box flexDirection="column">
        {modes}
        <Text> </Text>
        <Text dimColor>{the.none}</Text>
      </Box>
    )
  }

  return (
    <Box flexDirection="column">
      {modes}
      <Box flexDirection="row" flexWrap="wrap" columnGap={2}>
        <Select
          key="sort"
          label={the.sort}
          value={how.sort}
          options={SORTS.map(value => ({ value, label: the.sorts[value] }))}
          onSelect={value => {
            if (isSort(value)) {
              void arrange($, { sort: value })
            }
          }}
        />
        <Select
          key="filter"
          label={the.filter}
          value={how.filter}
          options={[
            { value: 'all', label: `${the.all} (${list.length})` },
            ...kinds.map(kind => ({
              value: kind,
              label: `${the.kinds[kind]} (${list.filter(one => one.kind === kind).length})`,
            })),
            ...(lost > 0 || how.filter === 'missing' ? [{ value: 'missing', label: `${the.missing} (${lost})` }] : []),
          ]}
          onSelect={value => {
            if (isFilter(value)) {
              void arrange($, { filter: value })
            }
          }}
        />
        <Select
          key="group"
          label={the.group}
          value={how.group}
          options={GROUPS.map(value => ({ value, label: the.groups[value] }))}
          onSelect={value => {
            if (isGroup(value)) {
              void arrange($, { group: value })
            }
          }}
        />
      </Box>
      <Text dimColor>{the.count(shown.length, list.length)}</Text>
      <Text> </Text>
      {shown.length === 0 && <Text dimColor>{the.noneFit}</Text>}
      {grouped(shown, how.group, the).map(group => (
        <Box flexDirection="column" key={`group-${group.key}`}>
          {group.title !== '' && (
            <Text bold>
              {group.title} ({group.items.length})
            </Text>
          )}
          {group.items.map(source => {
            const place = placeOf(source, the)

            return (
              <Box flexDirection="column" key={`row-${source.n}`}>
                <Box flexDirection="row" columnGap={1}>
                  <Button
                    key={`source-${source.n}`}
                    label={`${source.status === 'missing' ? '✗ ' : ''}${place}`}
                    dimColor={source.status === 'missing'}
                    onPress={() => show($, source)}
                  />
                  <Text dimColor>
                    {[
                      // The file's name, unless the heading or the button already says it
                      how.group === 'file' || source.kind === 'web' || place === nameOf(source.path)
                        ? ''
                        : nameOf(source.path),
                      source.status === 'missing' ? the.missing.toLowerCase() : '',
                      source.status === 'unchecked' ? the.unchecked : '',
                    ]
                      .filter(part => part !== '')
                      .join(' · ')}
                  </Text>
                </Box>
                {source.claim !== '' && <Text>{source.claim}</Text>}
                <Text dimColor>
                  «{source.quote.length > 160 ? `${source.quote.slice(0, 160)} …` : source.quote}»
                </Text>
                <Text> </Text>
              </Box>
            )
          })}
        </Box>
      ))}
    </Box>
  )
}
