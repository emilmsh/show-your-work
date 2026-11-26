// Runs the look-up (scripts/show.js) on the demo set, on a Mac: words that are
// there must be found at the place they are, and words that are NOT there must
// be refused. Started by check-mac.sh, which says how.
//
// Every answer is printed whole: what went wrong on a Mac nobody here has is
// read out of this output.
ObjC.import('Foundation')

function run(argv) {
  var me = Application.currentApplication()

  me.includeStandardAdditions = true

  var root = ObjC.unwrap($.NSProcessInfo.processInfo.environment.objectForKey('SYW_ROOT'))
  var quoted = function (text) {
    return "'" + String(text).replace(/'/g, "'\\''") + "'"
  }
  var sh = function (command) {
    try {
      return me.doShellScript(command, { alteringLineEndings: false })
    } catch (error) {
      return 'failed: ' + String(error.message || error)
    }
  }
  var ask = function (request) {
    var out = sh('SYW_REQUEST=' + quoted(JSON.stringify(request)) + ' /usr/bin/osascript -l JavaScript ' + quoted(root + '/scripts/show.js'))

    try {
      return JSON.parse(out.trim().split('\n').pop())
    } catch (error) {
      return { error: 'no answer: ' + out }
    }
  }
  var lines = []
  var say = function (line) {
    lines.push(line)
  }
  var version = function (app) {
    return sh('/usr/bin/mdls -raw -name kMDItemVersion ' + quoted('/Applications/' + app + '.app'))
  }

  say('macOS ' + sh('/usr/bin/sw_vers -productVersion'))
  ;['Microsoft Word', 'Microsoft Excel', 'Microsoft PowerPoint', 'Visual Studio Code', 'PDF Scholar'].forEach(function (app) {
    say(app + ': ' + version(app))
  })
  say('')

  // kind, file, quote, cell, found?, place (0 = not asked)
  var cases = [
    ['word', 'report.docx', 'Attendance rose by 4.2 percentage points in the pilot schools', '', true, 1],
    ['word', 'report.docx', 'We find no effect on test scores', '', true, 0],
    ['word', 'report.docx', 'Attendance rose by 5.2 percentage points', '', false, 0],
    ['powerpoint', 'slides.pptx', 'Attendance rose by 4.2 percentage points.', '', true, 3],
    ['powerpoint', 'slides.pptx', 'Recommendation', '', true, 4],
    ['powerpoint', 'slides.pptx', 'Attendance rose by 5.2 percentage points.', '', false, 0],
    ['excel', 'model.xlsx', '1987', 'Assumptions!B2', true, 0],
    ['excel', 'model.xlsx', '1 987', 'Assumptions!B2', true, 0],
    ['excel', 'model.xlsx', '9 537 600', 'Results!B2', true, 0],
    ['excel', 'model.xlsx', '25,26', 'Results!B3', true, 0],
    ['excel', 'model.xlsx', '4.2', 'Results!B:B', true, 0],
    ['excel', 'model.xlsx', 'Results', 'Results!', true, 0],
    ['excel', 'model.xlsx', '1988', 'Assumptions!B2', false, 0],
    ['excel', 'model.xlsx', '4.3', 'Results!B:B', false, 0],
    ['excel', 'model.xlsx', '1987', 'Nowhere!B2', false, 0],
    ['text', 'notes.md', 'Attendance rose most among pupils from low-income families.', '', true, 12],
    ['text', 'notes.md', 'We have attendance records for all 170 schools', '', false, 0],
    ['text', 'analysis.R', 'filter(n_distinct(year) == 7)', '', true, 10],
    ['text', 'analysis.R', 'filter(n_distinct(year) == 8)', '', false, 0],
  ]
  var failed = 0

  say('--- Looked up, nothing shown ---')
  cases.forEach(function (one) {
    var answer = ask({ kind: one[0], path: root + '/demo/' + one[1], quote: one[2], cell: one[3], show: false })
    var found = answer.found === true
    var ok = found === one[4] && (one[5] === 0 || answer.page === one[5])

    if (!ok) {
      failed++
    }

    say((ok ? 'ok    ' : 'FAIL  ') + one[0] + ' ' + one[3] + '  want found=' + one[4] + (one[5] ? ' place=' + one[5] : '') + '  \u00ab' + one[2].slice(0, 40) + '\u00bb')
    say('      ' + JSON.stringify(answer))
  })
  say(failed + ' of ' + cases.length + ' failed')

  if (argv[0] === 'show') {
    // One of each kind, opened at the words, with time to look at each
    var shown = [
      ['word', 'report.docx', 'Attendance rose by 4.2 percentage points in the pilot schools', ''],
      ['excel', 'model.xlsx', '9 537 600', 'Results!B2'],
      ['excel', 'model.xlsx', '4.2', 'Results!B:B'],
      ['powerpoint', 'slides.pptx', 'Recommendation', ''],
      ['text', 'analysis.R', 'filter(n_distinct(year) == 7)', ''],
      ['text', 'notes.md', 'Attendance rose most among pupils from low-income families.', ''],
      ['web', 'https://example.com/', 'Example Domain', ''],
    ]

    say('')
    say('--- Shown: each opens at the words ---')
    shown.forEach(function (one) {
      var path = one[0] === 'web' ? one[1] : root + '/demo/' + one[1]

      say(one[0] + ' ' + one[1] + ' ' + one[3])
      say('      ' + JSON.stringify(ask({ kind: one[0], path: path, quote: one[2], cell: one[3], show: true })))
      delay(4)
    })
  }

  return lines.join('\n')
}
