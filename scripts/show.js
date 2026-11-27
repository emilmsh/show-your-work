// Show your work on macOS: the same request and the same answer as
// scripts/show.ps1 on Windows, by the means a Mac has.
//
//   osascript -l JavaScript scripts/show.js        (JavaScript for Automation)
//
// The request is JSON in the environment (SYW_REQUEST):
//   { kind, path, quote, cell, show }      or      { kind: "raise", process }
// The answer is ONE line of JSON on standard output: { found, page, ... } or
// { error }.
//
//   Word          the words are looked up in the file itself (a .docx is a zip
//                 of XML); Word is asked for the PAGE and selects the words
//   PowerPoint    looked up in the file; PowerPoint goes to the slide
//   Excel         Excel's own automation, as on Windows
//   text, code    read here; shown in VS Code at line and column when it is
//                 installed, else in the default text editor without the place
//   HTML, web     shown in the default browser, by a text fragment
//
// What reads a file and judges a quote is plain JavaScript and runs under
// Node too, on any system: `node scripts/test/check-cases.mjs` holds it to
// the cases the Windows script is held to. What drives a program is the
// Mac's own and can only be checked on a Mac (scripts/test/check-mac.sh).
// Keep the two apart: nothing above "Programs" may touch `Application`.

var isMac = typeof ObjC !== 'undefined'

// ---------- The machine: reading a file, a zip's member, running a command ----------

var host = isMac ? macHost() : nodeHost()

function macHost() {
  ObjC.import('Foundation')

  var me = Application.currentApplication()

  me.includeStandardAdditions = true

  var quoted = function (text) {
    return "'" + String(text).replace(/'/g, "'\\''") + "'"
  }
  var shell = function (command) {
    return me.doShellScript(command, { alteringLineEndings: false })
  }

  return {
    me: me,
    quoted: quoted,
    shell: shell,
    request: function () {
      return ObjC.unwrap($.NSProcessInfo.processInfo.environment.objectForKey('SYW_REQUEST'))
    },
    // By its mark when it has one (UTF-16), else as UTF-8 when it IS that,
    // else as Windows' Western code page: what a file with a Norwegian letter
    // in it is when it is not UTF-8
    text: function (path) {
      var data = $.NSData.dataWithContentsOfFile(path)

      if (data.isNil()) {
        throw new Error('the file cannot be read')
      }

      var as = function (encoding) {
        return ObjC.unwrap($.NSString.alloc.initWithDataEncoding(data, encoding))
      }
      var head =
        data.length >= 2 ? ObjC.unwrap(data.subdataWithRange($.NSMakeRange(0, 2)).base64EncodedStringWithOptions(0)) : ''

      if (head === '//4=' || head === '/v8=') {
        return String(as(10)).replace(/^\ufeff/, '')
      }

      var utf8 = as(4)

      return typeof utf8 === 'string' ? utf8.replace(/^\ufeff/, '') : String(as(12))
    },
    member: function (zip, name) {
      try {
        return shell('/usr/bin/unzip -p ' + quoted(zip) + ' ' + quoted(name))
      } catch (error) {
        return ''
      }
    },
    // A .doc or an .rtf, which is no zip: the system's own converter
    converted: function (path) {
      return shell('/usr/bin/textutil -convert txt -stdout ' + quoted(path))
    },
  }
}

function nodeHost() {
  var fs = require('node:fs')
  var run = require('node:child_process').execFileSync
  // Windows' own tar reads a zip; Git's does not
  var unzip =
    process.platform === 'win32'
      ? function (zip, name) {
          return run(process.env.SystemRoot + '\\System32\\tar.exe', ['-xOf', zip, name], { stdio: ['ignore', 'pipe', 'ignore'] })
        }
      : function (zip, name) {
          return run('unzip', ['-p', zip, name], { stdio: ['ignore', 'pipe', 'ignore'] })
        }

  return {
    request: function () {
      return process.env.SYW_REQUEST
    },
    text: function (path) {
      var bytes = fs.readFileSync(path)

      if (bytes.length >= 2 && ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))) {
        return new TextDecoder(bytes[0] === 0xff ? 'utf-16le' : 'utf-16be').decode(bytes).replace(/^\ufeff/, '')
      }

      try {
        return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^\ufeff/, '')
      } catch (error) {
        return new TextDecoder('windows-1252').decode(bytes)
      }
    },
    member: function (zip, name) {
      try {
        return unzip(zip, name).toString('utf8')
      } catch (error) {
        return ''
      }
    },
    converted: function () {
      throw new Error('a .doc or .rtf file is read on a Mac only')
    },
  }
}

// ---------- Judging a quote: the same rules as show.ps1 ----------

// Whitespace of any kind matches any other; a footnote, cell or field mark, or
// a hyphen that only shows at a line's end, may sit between any two
// characters. A quotation mark matches any other of its kind and a hyphen any
// other hyphen: whoever quotes types the plain one for the typeset one.
function pattern(quote) {
  var marks = '[\\x01\\x02\\x07\\x0b\\x13-\\x15\\x1f\\u00ad\\u200b]'
  var kinds = ["['\\u2018\\u2019\\u201a\\u2032]", '["\\u201c\\u201d\\u201e\\u00ab\\u00bb\\u2033]', '[-\\u2010\\u2011\\x1e]']
  var one = function (character) {
    for (var i = 0; i < kinds.length; i++) {
      if (new RegExp(kinds[i]).test(character)) {
        return kinds[i]
      }
    }

    return character.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&')
  }
  var parts = String(quote)
    .replace(/[\x1f\u00ad\u200b]/g, '')
    .trim()
    .split(/[\s\u00a0\u202f]+/)
    .map(function (word) {
      return Array.from(word).map(one).join(marks + '*')
    })

  return new RegExp(parts.join('[\\s\\u00a0\\u202f\\x01\\x02\\x07\\x0b\\x13-\\x15]+'))
}

// A number as somebody wrote it, in every way it can be read: "1 987",
// "1,987" and "1.987" are all 1987 to someone, and the last two are also
// 1.987. Each reading is { value, decimals, percent }; none when the text is
// no number. A typeset minus is a minus.
function readings(text) {
  var compact = String(text)
    .replace(/[\s\u00a0\u202f]/g, '')
    .replace(/\u2212/g, '-')
    .replace(/%$/, '')
  var percent = /%$/.test(String(text).trim())
  var read = []
  var grouped =
    /^(-?\d{1,3})((?:[.,]\d{3})+)$/.exec(compact) ||
    /^(-?\d{1,3})((?:\.\d{3})+),(\d+)$/.exec(compact) ||
    /^(-?\d{1,3})((?:,\d{3})+)\.(\d+)$/.exec(compact)

  // Groups of three, with one separator between them, not both
  if (grouped && !(grouped[2].indexOf('.') >= 0 && grouped[2].indexOf(',') >= 0)) {
    var decimals = grouped[3] || ''
    var whole = grouped[1] + grouped[2].replace(/[.,]/g, '')

    read.push({ value: Number(decimals === '' ? whole : whole + '.' + decimals), decimals: decimals.length, percent: percent })
  }

  // One separator at most: the decimal sign
  var plain = /^-?\d+(?:[.,](\d+))?$/.exec(compact)

  if (plain) {
    read.push({ value: Number(compact.replace(',', '.')), decimals: plain[1] ? plain[1].length : 0, percent: percent })
  }

  return read
}

// A cell holds the value cited when it SHOWS it (as text, blanks aside), or
// when its number is the number cited, in one of the ways that can be read,
// to as many decimals as were cited
function sameValue(shown, value, quote) {
  var norm = function (text) {
    return String(text == null ? '' : text)
      .replace(/[\s\u00a0\u202f]/g, '')
      .replace(/\u2212/g, '-')
  }

  if (norm(shown) === norm(quote)) {
    return true
  }

  if (typeof value !== 'number') {
    return false
  }

  return readings(quote).some(function (one) {
    return Math.abs((one.percent ? value * 100 : value) - one.value) < 0.5 * Math.pow(10, -one.decimals) + 1e-9
  })
}

function decoded(text) {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z][a-z0-9]*);/gi, function (whole, name) {
    if (name[0] === '#') {
      var code = name[1] === 'x' || name[1] === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10)

      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole
    }

    var known = ENTITIES[name]

    return known === undefined ? whole : String.fromCodePoint(known)
  })
}

var ENTITIES = (function () {
  var table = { amp: 38, lt: 60, gt: 62, quot: 34, apos: 39 }
  var runs = [
    [160, 'nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn sup2 sup3 acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde auml aring aelig ccedil egrave eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde ouml divide oslash ugrave uacute ucirc uuml yacute thorn yuml'],
    [913, 'Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu Nu Xi Omicron Pi Rho'],
    [931, 'Sigma Tau Upsilon Phi Chi Psi Omega'],
    [945, 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigmaf sigma tau upsilon phi chi psi omega'],
  ]
  var single =
    'OElig 338 oelig 339 Scaron 352 scaron 353 Yuml 376 fnof 402 circ 710 tilde 732 ensp 8194 emsp 8195 thinsp 8201 zwnj 8204 zwj 8205 ndash 8211 mdash 8212 lsquo 8216 rsquo 8217 sbquo 8218 ldquo 8220 rdquo 8221 bdquo 8222 dagger 8224 Dagger 8225 bull 8226 hellip 8230 permil 8240 prime 8242 Prime 8243 lsaquo 8249 rsaquo 8250 euro 8364 trade 8482 larr 8592 uarr 8593 rarr 8594 darr 8595 harr 8596 part 8706 prod 8719 sum 8721 minus 8722 radic 8730 infin 8734 asymp 8776 ne 8800 equiv 8801 le 8804 ge 8805'.split(' ')

  runs.forEach(function (run) {
    run[1].split(' ').forEach(function (name, i) {
      table[name] = run[0] + i
    })
  })

  for (var i = 0; i < single.length; i += 2) {
    table[single[i]] = Number(single[i + 1])
  }

  return table
})()

// ---------- Reading the files ----------

// A text file: the words are looked up in the file itself and the place is a
// LINE. Markdown is quoted as it READS: without the marks for bold, italics
// and code, and a link by its words - tried only when the words as written
// are not there, and nothing is taken out but inside a line, so the line holds.
function inText(path, quote) {
  var text = host.text(path)
  var looked = pattern(quote)
  var match = looked.exec(text)

  if (!match && /\.(md|markdown|qmd|rmd)$/i.test(path)) {
    try {
      var plain = text
        .replace(/!?\[([^\]\r\n]*)\]\([^)\r\n]*\)/g, '$1')
        .replace(new RegExp('\\*+(?=\\S)|(?<=\\S)\\*+|`+|(?<![\\p{L}\\p{N}_\\\\])_+(?=\\S)|(?<=\\S)_+(?![\\p{L}\\p{N}_])', 'gu'), '')
      var again = looked.exec(plain)

      if (again) {
        match = again
        text = plain
      }
    } catch (error) {
      // A system whose JavaScript has no look-behind: the words as written only
    }
  }

  if (!match) {
    return { found: false }
  }

  var line = text.slice(0, match.index).split('\n').length

  return { found: true, page: line, column: match.index - (text.lastIndexOf('\n', Math.max(match.index - 1, 0)) + 1) + 1 }
}

// An HTML file's text as it reads: no comment, script, style sheet or head; a
// tag that sets text apart (a paragraph, a line break, a cell, a list item) is
// a blank, and one inside a sentence or a word (bold, a link) is nothing at
// all. Then every entity, named or numbered, is read.
function inHtml(path, quote) {
  var inline = 'a|abbr|b|bdi|bdo|cite|code|del|dfn|em|font|i|ins|kbd|mark|q|s|samp|small|span|strike|strong|sub|sup|tt|u|var|wbr'
  var text = host
    .text(path)
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(script|style|head|template)\b[\s\S]*?<\/\1\s*>/gi, ' ')
    .replace(new RegExp('</?(?:' + inline + ')(?:\\s[^>]*)?>', 'gi'), '')
    .replace(/<[^>]+>/g, ' ')

  return { found: pattern(quote).test(decoded(text)) }
}

// The text of one stretch of Office XML: its paragraphs, each run's text, a
// tab, and a line break inside a paragraph as the mark Office itself uses.
// `p` is the namespace: "w" in a Word file, "a" in a PowerPoint file.
function paragraphs(xml, p) {
  var paragraph = new RegExp('<' + p + ':p(?: [^>]*)?>([\\s\\S]*?)</' + p + ':p>', 'g')
  var piece = new RegExp(
    '<' + p + ':t(?: [^>]*)?>([\\s\\S]*?)</' + p + ':t>|<' + p + ':(br|cr|tab|noBreakHyphen|softHyphen)\\b[^>]*/>',
    'g',
  )
  var marks = { br: '\x0b', cr: '\x0b', tab: '\t', noBreakHyphen: '\x1e', softHyphen: '\x1f' }
  var out = []
  var one

  while ((one = paragraph.exec(xml)) !== null) {
    var text = ''
    var bit

    piece.lastIndex = 0

    while ((bit = piece.exec(one[1])) !== null) {
      text += bit[2] ? marks[bit[2]] : decoded(bit[1])
    }

    out.push(text)
  }

  return out.join('\r')
}

// A Word file's words: the body, then the footnotes and the endnotes, as
// Word keeps them - each a text of its own
function inWord(path, quote) {
  var looked = pattern(quote)
  var stories = [['body', 'word/document.xml'], ['footnotes', 'word/footnotes.xml'], ['endnotes', 'word/endnotes.xml']]

  if (!/\.(docx|docm)$/i.test(path)) {
    return { found: looked.test(host.converted(path)), where: 'body' }
  }

  for (var i = 0; i < stories.length; i++) {
    var xml = host.member(path, stories[i][1])

    if (i === 0 && xml === '') {
      throw new Error('the file cannot be read as a Word document')
    }

    if (looked.test(paragraphs(xml, 'w'))) {
      return { found: true, where: stories[i][0] }
    }
  }

  return { found: false }
}

// A presentation's words, slide by slide in the order they are shown: every
// text box whole (a group's too, however deep), every table cell, each table
// ROW once more as a line - a row is what is quoted from a table - and then
// the slide's speaker notes
function inPowerPoint(path, quote) {
  if (!/\.(pptx|pptm|ppsx)$/i.test(path)) {
    throw new Error('a .ppt file cannot be read here: save it as .pptx')
  }

  var looked = pattern(quote)
  var targets = function (rels) {
    var found = {}
    var relation = /<Relationship\b[^>]*>/g
    var one

    while ((one = relation.exec(rels)) !== null) {
      var id = /\bId="([^"]+)"/.exec(one[0])
      var target = /\bTarget="([^"]+)"/.exec(one[0])
      var type = /\bType="([^"]+)"/.exec(one[0])

      if (id && target) {
        found[id[1]] = { target: target[1], type: type ? type[1] : '' }
      }
    }

    return found
  }
  // "slides/slide1.xml" beside the presentation, "../notesSlides/n.xml"
  // beside a slide, or a whole path from the zip's root
  var under = function (folder, target) {
    if (target[0] === '/') {
      return target.slice(1)
    }

    var parts = folder.split('/')

    target.split('/').forEach(function (part) {
      if (part === '..') {
        parts.pop()
      } else if (part !== '.') {
        parts.push(part)
      }
    })

    return parts.join('/')
  }
  var texts = function (xml) {
    var all = []
    var body = /<(p|a):txBody>([\s\S]*?)<\/\1:txBody>/g
    var row = /<a:tr\b[\s\S]*?<\/a:tr>/g
    var one

    while ((one = body.exec(xml)) !== null) {
      all.push(paragraphs(one[2], 'a'))
    }

    while ((one = row.exec(xml)) !== null) {
      var cells = []
      var cell = /<a:txBody>([\s\S]*?)<\/a:txBody>/g
      var each

      while ((each = cell.exec(one[0])) !== null) {
        cells.push(paragraphs(each[1], 'a'))
      }

      all.push(cells.join('\t'))
    }

    return all
  }
  var hits = function (list) {
    return list.some(function (text) {
      return looked.test(text)
    })
  }
  var presentation = host.member(path, 'ppt/presentation.xml')

  if (presentation === '') {
    throw new Error('the file cannot be read as a presentation')
  }

  var files = targets(host.member(path, 'ppt/_rels/presentation.xml.rels'))
  var order = /<p:sldId\b[^>]*\br:id="([^"]+)"/g
  var slide
  var n = 0

  while ((slide = order.exec(presentation)) !== null) {
    n++

    if (!files[slide[1]]) {
      continue
    }

    var name = under('ppt', files[slide[1]].target)

    if (hits(texts(host.member(path, name)))) {
      return { found: true, page: n, where: 'slide' }
    }

    var folder = name.slice(0, name.lastIndexOf('/'))
    var beside = targets(host.member(path, folder + '/_rels/' + name.slice(folder.length + 1) + '.rels'))
    var notes = Object.keys(beside).filter(function (id) {
      return /\/notesSlide$/.test(beside[id].type)
    })[0]

    if (notes) {
      // The notes page holds the slide's picture and number too: only the
      // notes themselves are the notes
      var shapes = host.member(path, under(folder, beside[notes].target)).match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
      var written = shapes.filter(function (shape) {
        return /<p:ph\b[^>]*\btype="body"/.test(shape)
      })

      if (hits(texts(written.join('')))) {
        return { found: true, page: n, where: 'notes' }
      }
    }
  }

  return { found: false }
}

// "B2", "B2:D10", "5:5", "C:E" as { r1, c1, r2, c2 } (1-based; 0 = open)
function rectangle(address) {
  var column = function (letters) {
    var n = 0

    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64)
    }

    return n
  }
  var ends = String(address).replace(/\$/g, '').trim().split(':')
  var parse = function (end) {
    var parts = /^([a-z]{0,3})(\d*)$/i.exec(end)

    return parts && end !== '' ? { c: parts[1] === '' ? 0 : column(parts[1]), r: parts[2] === '' ? 0 : Number(parts[2]) } : null
  }
  var from = parse(ends[0])
  var to = ends.length > 1 ? parse(ends[1]) : from

  if (!from || !to || ends.length > 2) {
    return null
  }

  return { r1: from.r, c1: from.c, r2: to.r, c2: to.c }
}

function letters(column) {
  var text = ''

  while (column > 0) {
    text = String.fromCharCode(65 + ((column - 1) % 26)) + text
    column = Math.floor((column - 1) / 26)
  }

  return text
}

// A sheet's name as Excel writes it in a formula: in quotation marks when it
// has a space, and an apostrophe of its own doubled
function sheetNamed(name) {
  var quotedName = /^'(.*)'$/.exec(name.trim())

  return quotedName ? quotedName[1].replace(/''/g, "'") : name.trim()
}

// ---------- Programs (macOS only below this line) ----------

function samePath(theirs, mine) {
  var a = String(theirs || '').normalize('NFC')
  var b = String(mine).normalize('NFC')

  // Older Office answers "Macintosh HD:Users:me:file.docx"
  if (a.indexOf('/') < 0 && a.indexOf(':') >= 0) {
    a = '/' + a.split(':').slice(1).join('/')
  }

  return a.toLowerCase() === b.toLowerCase()
}

function opened(app, list, path) {
  var all = []

  // No list at all (null) is what a program answers while a window of its
  // own stands in front of every document: a sign-in or licence screen
  try {
    all = list() || []
  } catch (error) {
    return null
  }

  for (var i = 0; i < all.length; i++) {
    var names = []

    try {
      names.push(all[i].posixFullName())
    } catch (error) {}

    try {
      names.push(all[i].fullName())
    } catch (error) {}

    if (names.some(function (name) { return samePath(name, path) })) {
      return all[i]
    }
  }

  return null
}

// The file in its program: where it is already open, or opened the way a
// double-click opens it (which a sandboxed Office accepts without asking for
// the folder) - out of sight when it is only to be checked
function inProgram(name, list, path, show) {
  var app
  var wasRunning

  try {
    app = Application(name)
    wasRunning = app.running()
  } catch (error) {
    throw new Error(name + ' is not installed')
  }

  var file = wasRunning ? opened(app, list(app), path) : null
  var wasOpen = file !== null

  if (!wasOpen) {
    host.shell('/usr/bin/open ' + (show ? '' : '-g -j ') + '-a ' + host.quoted(name) + ' ' + host.quoted(path))

    for (var i = 0; i < 150 && file === null; i++) {
      delay(0.2)
      file = opened(app, list(app), path)
    }

    if (file === null) {
      // Also what a program without a licence comes to: it starts on an
      // offer to buy it, and opens nothing behind it
      throw new Error(name + ' did not open the file')
    }
  }

  return { app: app, file: file, wasOpen: wasOpen, wasRunning: wasRunning }
}

function shut(held, close) {
  if (!held.wasOpen) {
    try {
      close(held.file)
    } catch (error) {}

    if (!held.wasRunning) {
      try {
        held.app.quit()
      } catch (error) {}
    }
  }
}

function showWord(req, show) {
  var res = inWord(req.path, req.quote)

  // Not in the file: Word is not started to say so again
  if (!res.found && !show) {
    return res
  }

  var held

  try {
    held = inProgram('Microsoft Word', function (app) { return function () { return app.documents() } }, req.path, show)
  } catch (error) {
    // Found in the file, and no Word to ask for the page
    res.program = String(error.message || error)

    return res
  }

  var word = held.app
  var doc = held.file

  res.wasOpen = held.wasOpen

  try {
    // WHERE in the document the words are. A place in the text is not a place
    // in the document (a field's code takes room in the one and not in the
    // other), so the place is ASKED for, as on Windows: the text in front of
    // a position only grows with the position, and the position with exactly
    // the match's text in front of it is searched out.
    var body = doc.textObject
    var text = String(body.content() || '')
    var match = pattern(req.quote).exec(text)

    if (match) {
      var base = body.startOfContent()
      var last = body.endOfContent()
      var span = function (a, b) {
        return word.createRange(doc, { start: a, end: b })
      }
      var before = function (position) {
        return position <= base ? 0 : String(span(base, position).content() || '').length
      }
      var lo = base
      var hi = last

      while (lo < hi) {
        var mid = Math.ceil((lo + hi) / 2)

        if (before(mid) <= match.index) {
          lo = mid
        } else {
          hi = mid - 1
        }
      }

      var from = lo

      lo = from + 1
      hi = Math.min(last, from + 2 * match[0].length + 2000)

      while (lo < hi) {
        var middle = Math.floor((lo + hi) / 2)

        if (String(span(from, middle).content() || '').length >= match[0].length) {
          hi = middle
        } else {
          lo = middle + 1
        }
      }

      var range = span(from, lo)

      // Said only when it holds: the words at the place are the words found.
      // Else the place is left UNSAID - no page, nothing selected.
      res.exact = String(range.content() || '') === match[0]

      if (res.exact) {
        var page

        try {
          page = range.getRangeInformation({ informationType: 'active end page number' })
        } catch (error) {
          page = word.getRangeInformation(range, { informationType: 'active end page number' })
        }

        if (Number(page) > 0) {
          res.page = Number(page)
        }

        if (show) {
          try {
            range.select()
          } catch (error) {
            word.select(range)
          }
        }
      }
    }
  } catch (error) {
    res.place = String(error.message || error)
  }

  if (show) {
    word.activate()
    res.raised = 'foreground'
  } else {
    shut(held, function (file) { file.close({ saving: 'no' }) })
  }

  return res
}

function showPowerPoint(req, show) {
  var res = inPowerPoint(req.path, req.quote)

  if (!show) {
    return res
  }

  var held = inProgram('Microsoft PowerPoint', function (app) { return function () { return app.presentations() } }, req.path, true)

  res.wasOpen = held.wasOpen
  held.app.activate()
  res.raised = 'foreground'

  if (res.found) {
    // The slide; PowerPoint on a Mac selects no words on another's word
    try {
      held.file.documentWindows[0].view.goToSlide({ number: res.page })
    } catch (error) {
      try {
        held.app.activeWindow.view.goToSlide({ number: res.page })
      } catch (again) {
        res.place = String(again.message || again)
      }
    }
  }

  return res
}

function showExcel(req, show) {
  var held = inProgram('Microsoft Excel', function (app) { return function () { return app.workbooks() } }, req.path, show)
  var excel = held.app
  var book = held.file
  var res = { found: false, wasOpen: held.wasOpen }

  // The place is a cell ("Sheet!B2"), a row, a column or a block ("Sheet!5:5",
  // "Sheet!C:C", "Sheet!B2:D10"), or a sheet alone ("Sheet!" or its name).
  // A cell is checked against the value cited; a row, column or block holds
  // the citation when ONE of its cells does; a sheet alone only has to exist.
  try {
    var cell = String(req.cell || '')
    var at = cell.lastIndexOf('!')
    var address = cell.slice(at + 1).trim()
    var sheet

    if (at > 0) {
      sheet = book.worksheets.byName(sheetNamed(cell.slice(0, at)))
      res.sheet = sheet.name()
    } else {
      try {
        sheet = book.worksheets.byName(sheetNamed(cell))
        res.sheet = sheet.name()
        address = ''
      } catch (error) {
        sheet = book.worksheets[0]
        res.sheet = sheet.name()
      }
    }

    var grid = function (value) {
      return Array.isArray(value) ? (Array.isArray(value[0]) ? value : [value]) : [[value]]
    }
    var where = rectangle(address)
    var selected = null

    if (address === '') {
      res.found = true
      res.whole = 'sheet'
    } else if (!where) {
      throw new Error('no such cell')
    } else if (where.r1 > 0 && where.c1 > 0 && where.r1 === where.r2 && where.c1 === where.c2) {
      var one = sheet.ranges.byName(letters(where.c1) + where.r1)

      // A cell of a merged block is the block
      try {
        if (one.mergeCells() === true) {
          var block = rectangle(one.mergeArea.getAddress())

          if (block) {
            one = sheet.ranges.byName(letters(block.c1) + block.r1)
          }
        }
      } catch (error) {}

      res.text = String(one.stringValue())
      res.formula = String(one.formula())
      res.found = sameValue(res.text, one.value(), req.quote)
      selected = one
    } else {
      res.whole = 'range'

      // The part of it that holds anything
      var used = rectangle(sheet.usedRange.getAddress())
      var r1 = Math.max(where.r1 || 1, used.r1)
      var r2 = Math.min(where.r2 || 1048576, used.r2)
      var c1 = Math.max(where.c1 || 1, used.c1)
      var c2 = Math.min(where.c2 || 16384, used.c2)

      selected = sheet.ranges.byName(address.replace(/\$/g, ''))

      if (r1 <= r2 && c1 <= c2) {
        // 5000 cells at most, as on Windows
        r2 = Math.min(r2, r1 + Math.max(1, Math.floor(5000 / (c2 - c1 + 1))) - 1)

        var part = sheet.ranges.byName(letters(c1) + r1 + ':' + letters(c2) + r2)
        var shown = grid(part.stringValue())
        var values = grid(part.value())

        for (var r = 0; r < shown.length && !res.found; r++) {
          for (var c = 0; c < shown[r].length && !res.found; c++) {
            if (sameValue(shown[r][c], values[r] ? values[r][c] : undefined, req.quote)) {
              res.found = true
              res.text = String(shown[r][c])
              res.at = letters(c1 + c) + (r1 + r)
            }
          }
        }
      }
    }

    if (show) {
      excel.activate()

      try {
        sheet.activateObject()

        if (selected) {
          selected.select()
        }

        // A row, column or block is selected whole, and the cell in it that
        // holds the value is the one the view goes to
        if (res.at) {
          sheet.ranges.byName(res.at).activateObject()
        }
      } catch (error) {
        res.place = String(error.message || error)
      }
    }
  } catch (error) {
    res.error = 'no such sheet or cell'
    res.detail = String(error.message || error)
  }

  if (show) {
    res.raised = 'foreground'
  } else {
    shut(held, function (file) { file.close({ saving: 'no' }) })
  }

  return res
}

function showText(req, show) {
  var res = inText(req.path, req.quote)
  var column = res.column || 1

  delete res.column

  if (!show) {
    return res
  }

  // VS Code takes a line and a column in its address; the cursor lands at the
  // first word. Without it the file opens in the default text editor, and the
  // place is the reader's to find.
  try {
    host.shell('/usr/bin/open -Ra "Visual Studio Code"')
    host.shell('/usr/bin/open ' + host.quoted('vscode://file' + encodeURI(req.path).replace(/[?#]/g, encodeURIComponent) + ':' + (res.page || 1) + ':' + column))
    res.editor = 'code'
  } catch (error) {
    host.shell('/usr/bin/open -t ' + host.quoted(req.path))
    res.editor = 'default, without the place'
  }

  res.raised = 'foreground'

  return res
}

// An HTML file, or a web page. A web page is never fetched, so its words are
// never checked. Either is shown in the default browser, which scrolls to the
// words and marks them itself (a text fragment, "#:~:text=").
function showHtml(req, show) {
  var isWeb = req.kind === 'web'
  var res = isWeb ? { found: false } : inHtml(req.path, req.quote)

  if (!show) {
    return res
  }

  var words = String(req.quote).trim().split(/\s+/)
  var escaped = function (text) {
    return encodeURIComponent(text).replace(/[-!'()*]/g, function (c) {
      return '%' + c.charCodeAt(0).toString(16).toUpperCase()
    })
  }
  var fragment =
    words.length <= 12
      ? escaped(words.join(' '))
      : escaped(words.slice(0, 5).join(' ')) + ',' + escaped(words.slice(-5).join(' '))
  var url = isWeb
    ? String(req.path).split('#')[0] + '#:~:text=' + fragment
    : 'file://' + encodeURI(req.path).replace(/[?#]/g, encodeURIComponent) + '#:~:text=' + fragment

  try {
    host.me.openLocation(url)
    res.browser = 'default'
  } catch (error) {
    host.shell('/usr/bin/open ' + host.quoted(isWeb ? String(req.path) : req.path))
    res.browser = 'default, without the words'
  }

  return res
}

// A window of a program that showed something on its own word (PDF Scholar,
// through its own server): brought to the front the way the others are
function showRaise(req) {
  try {
    Application(req.process).activate()
  } catch (error) {
    throw new Error(req.process + ' is not installed')
  }

  return { found: true, raised: 'foreground' }
}

// ---------- The request ----------

function main() {
  var res

  try {
    var req = JSON.parse(host.request() || '{}')
    var show = req.show === true

    if (!isMac) {
      // Under Node: what can be judged without a program, and nothing shown
      if (show) throw new Error('nothing can be shown from here')

      res =
        req.kind === 'word' ? inWord(req.path, req.quote)
        : req.kind === 'powerpoint' ? inPowerPoint(req.path, req.quote)
        : req.kind === 'text' ? inText(req.path, req.quote)
        : req.kind === 'html' ? inHtml(req.path, req.quote)
        : { error: 'unknown kind' }
      delete res.column
    } else {
      res =
        req.kind === 'word' ? showWord(req, show)
        : req.kind === 'powerpoint' ? showPowerPoint(req, show)
        : req.kind === 'excel' ? showExcel(req, show)
        : req.kind === 'text' ? showText(req, show)
        : req.kind === 'html' || req.kind === 'web' ? showHtml(req, show)
        : req.kind === 'raise' ? showRaise(req)
        : { error: 'unknown kind' }
    }
  } catch (error) {
    res = { error: String((error && error.message) || error).trim() }
  }

  return JSON.stringify(res)
}

// osascript calls run() and prints what it returns; Node prints it itself
function run() {
  return main()
}

if (!isMac) {
  console.log(main())
}
