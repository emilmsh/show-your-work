# Show your work: find a quoted passage in a source file, say where it is,
# and - when asked to SHOW - open the file in its own program AT the words
# and bring that window up. Windows only.
#
#   Word, PowerPoint, Excel   the program's own automation (COM)
#   text files                read here; shown in Notepad through UI Automation
#   code files                read here; shown in VS Code at line and column
#   HTML files, web pages     shown in the default browser, by a text fragment
#
# The request is JSON in the environment (SYW_REQUEST):
#   { kind, path, quote, cell, show }      or      { kind: "raise", process }
# The answer is ONE line of JSON on standard output: { found, page, ... } or
# { error }. Nothing else may be written there - a method that returns a
# value must be cast to [void], or its "True" becomes part of the answer.
#
# ASCII only: Windows PowerShell reads a script without a byte-order mark as
# ANSI, so a no-break space in a pattern is written \u00a0, never typed.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$req = $env:SYW_REQUEST | ConvertFrom-Json
$show = $req.show -eq $true

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class SywWin {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
  // Every file some program has open and has said so: { name, object }.
  // GetActiveObject answers for ONE instance of a program, and not always
  // the one on screen; this finds the file in whichever instance holds it.
  public static List<object[]> OpenFiles() {
    var found = new List<object[]>();
    IRunningObjectTable rot; IBindCtx ctx; IEnumMoniker all;
    if (GetRunningObjectTable(0, out rot) != 0 || CreateBindCtx(0, out ctx) != 0) return found;
    rot.EnumRunning(out all);
    var one = new IMoniker[1];
    while (all.Next(1, one, IntPtr.Zero) == 0) {
      try {
        string name; object o;
        one[0].GetDisplayName(ctx, null, out name);
        if (name == null || name.StartsWith("!")) continue;
        if (rot.GetObject(one[0], out o) == 0 && o != null) found.Add(new object[] { name, o });
      } catch { }
    }
    return found;
  }
}
"@
# A program opens a window where it last stood, and that can be on a screen
# that is no longer there (Word, after the laptop left its dock: the window
# opened above the top edge and could not be reached). Such a window is moved
# onto the main screen, its size kept as far as the screen allows.
function OnScreen($h) {
  $r = New-Object SywWin+RECT
  if (-not [SywWin]::GetWindowRect($h, [ref]$r)) { return }
  $box = New-Object Drawing.Rectangle($r.L, $r.T, ($r.R - $r.L), ($r.B - $r.T))
  foreach ($s in [Windows.Forms.Screen]::AllScreens) {
    $seen = [Drawing.Rectangle]::Intersect($s.WorkingArea, $box)
    if ($seen.Width -ge 200 -and $seen.Height -ge 100) { return }
  }
  $area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
  $w = [Math]::Min([Math]::Max($box.Width, 600), $area.Width); $t = [Math]::Min([Math]::Max($box.Height, 400), $area.Height)
  # 0x14 = keep the stacking order, no activation
  [void][SywWin]::SetWindowPos($h, [IntPtr]::Zero, [int]($area.X + ($area.Width - $w) / 2), [int]($area.Y + ($area.Height - $t) / 2), $w, $t, 0x14)
}
function Raise($hwnd, [string]$process) {
  $h = [IntPtr]::Zero
  if ($hwnd) { $h = [IntPtr][long]$hwnd }
  else { $p = Get-Process $process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1; if ($p) { $h = $p.MainWindowHandle } }
  if ($h -eq [IntPtr]::Zero) { return $false }
  if ([SywWin]::IsIconic($h)) { [void][SywWin]::ShowWindow($h, 9) }
  OnScreen $h
  [void][SywWin]::SetForegroundWindow($h)
  if ([SywWin]::GetForegroundWindow() -eq $h) { return 'foreground' }
  # Windows refuses the foreground to a program the user is not working in
  # (measured: refused while the machine is in use, granted after hours idle).
  # Topmost on and off again puts the window above the others all the same;
  # the keyboard stays where it was. 0x13 = no size, no move, no activation.
  [void][SywWin]::SetWindowPos($h, [IntPtr](-1), 0, 0, 0, 0, 0x13)
  [void][SywWin]::SetWindowPos($h, [IntPtr](-2), 0, 0, 0, 0, 0x13)
  'above'
}
function App([string]$progId) {
  try { @{ app = [Runtime.InteropServices.Marshal]::GetActiveObject($progId); own = $false } }
  catch { @{ app = New-Object -ComObject $progId; own = $true } }
}
# Whitespace of any kind matches any other; a footnote, cell or field mark, or
# a hyphen that only shows at a line's end, may sit between any two
# characters. A quotation mark matches any other of its kind and a hyphen any
# other hyphen: whoever quotes types the plain one for the typeset one.
function Pattern([string]$q) {
  $marks = '[\x01\x02\x07\x0b\x13-\x15\x1f\u00ad\u200b]'
  $kinds = @('[''\u2018\u2019\u201a\u2032]', '["\u201c\u201d\u201e\u00ab\u00bb\u2033]', '[-\u2010\u2011\x1e]')
  $one = {
    param([string]$c)
    foreach ($kind in $kinds) { if ([regex]::IsMatch($c, $kind)) { return $kind } }
    [regex]::Escape($c)
  }
  $q = [regex]::Replace($q, '[\x1f\u00ad\u200b]', '')
  $parts = ($q.Trim() -split '[\s\u00a0\u202f]+') | ForEach-Object {
    (($_.ToCharArray() | ForEach-Object { & $one ([string]$_) }) -join "$marks*")
  }
  $parts -join '[\s\u00a0\u202f\x01\x02\x07\x0b\x13-\x15]+'
}

function Show-Word {
  $a = App 'Word.Application'; $app = $a.app
  $doc = $null
  foreach ($d in $app.Documents) { if ($d.FullName -eq $req.path) { $doc = $d } }
  $opened = $doc -eq $null
  if ($opened) {
    if ($show) { $doc = $app.Documents.Open($req.path); if (-not $app.Visible) { $app.Visible = $true } }
    else { $m = [Type]::Missing; $doc = $app.Documents.Open($req.path, $false, $true, $false, $m, $m, $m, $m, $m, $m, $m, $false) }
  }
  # A document can be open with NO window on screen: one that was opened out
  # of sight to check a quote and never closed (the check was stopped half
  # way), or one another program holds. Shown means seen - measured: a press
  # "worked", the answer said so, and nothing came up.
  if ($show) {
    if (-not $app.Visible) { $app.Visible = $true }
    foreach ($w in $doc.Windows) { if (-not $w.Visible) { $w.Visible = $true } }
  }
  # The body first, then the footnotes and the endnotes: each is a text of its
  # own in Word (a "story"), with places of its own
  $story = $null; $mt = [regex]::Match('', 'x')
  foreach ($kind in 1, 2, 3) {
    if (-not $mt.Success) {
      try { $story = $doc.StoryRanges.Item($kind); $mt = [regex]::Match([string]$story.Text, (Pattern $req.quote)) } catch { }
    }
  }
  $res = @{ found = $mt.Success; wasOpen = -not $opened }
  if ($mt.Success) {
    $base = $story.Start; $last = $story.End
    $span = { param([int]$a, [int]$b) $r = $story.Duplicate; $r.SetRange($a, $b); $r }
    # WHERE in the document the words are. A place in .Text is not a place in
    # the document: a field's code (every entry of a table of contents, every
    # cross-reference, link and citation), a table cell's end and more take
    # room in the document and none, or another amount, in its text. Counted
    # by hand the place came out pages too early in any real report (measured
    # 2026-11-26: a quote on page 25 answered as page 19, and the press went
    # to other words). So the place is ASKED for: the text in front of a
    # position only grows with the position, and the position with exactly
    # the match's text in front of it is searched out.
    $before = { param([int]$p) if ($p -le $base) { 0 } else { ([string](& $span $base $p).Text).Length } }
    $lo = $base; $hi = $last
    while ($lo -lt $hi) {
      $mid = [int][Math]::Ceiling(($lo + $hi) / 2)
      if ((& $before $mid) -le $mt.Index) { $lo = $mid } else { $hi = $mid - 1 }
    }
    $from = $lo
    # The end, from there: the shortest stretch that holds all the words
    $lo = $from + 1; $hi = [Math]::Min($last, $from + 2 * $mt.Length + 2000)
    while ($lo -lt $hi) {
      $mid = [int][Math]::Floor(($lo + $hi) / 2)
      if (([string](& $span $from $mid).Text).Length -ge $mt.Length) { $hi = $mid } else { $lo = $mid + 1 }
    }
    $rng = & $span $from $lo
    # Said only when it holds: the words at the place are the words found
    $res.exact = ([string]$rng.Text) -eq $mt.Value
    if (-not $res.exact) {
      # Not pinned down that way: Word's own search, on the words' opening
      # (it takes 255 characters at most, and no line break). Failing that
      # too the place is left UNSAID - no page, nothing selected - which is
      # better than the wrong one.
      $res.at = ([string]$rng.Text)
      $head = ([regex]::Split($mt.Value, '[\x01\x02\x07\x0b\r\n\x13-\x15]')[0])
      $head = $head.Substring(0, [Math]::Min(60, $head.Length))
      $rng = $story.Duplicate
      if ($head.Length -lt 8 -or -not $rng.Find.Execute($head, $true)) { $rng = $null }
    }
    if ($rng) { $res.page = $rng.Information(3) }
    if ($show -and -not $rng) { $doc.Activate() }
    elseif ($show) { $doc.Activate(); [void]$rng.Select(); $app.ActiveWindow.ScrollIntoView($rng, $true) }
  }
  if ($show) { $res.raised = Raise $app.ActiveWindow.Hwnd 'WINWORD' }
  elseif ($opened) { $doc.Close(0); if ($a.own) { $app.Quit() } }
  $res
}

function Show-PowerPoint {
  $a = App 'PowerPoint.Application'; $app = $a.app
  $pres = $null
  foreach ($x in $app.Presentations) { if ($x.FullName -eq $req.path) { $pres = $x } }
  $opened = $pres -eq $null
  if ($opened) { if ($show) { $pres = $app.Presentations.Open($req.path) } else { $pres = $app.Presentations.Open($req.path, -1, 0, 0) } }
  $pattern = Pattern $req.quote
  $hit = $null
  foreach ($s in $pres.Slides) {
    # Every text of the slide: { where, its range (none for a table row read
    # as one line), its text }. A group's shapes are reached through the
    # group, however deep; a table's cells one by one, and each ROW once more
    # as a line, since a row is what is quoted from a table.
    $ranges = New-Object Collections.ArrayList
    $queue = New-Object Collections.Queue
    foreach ($sh in $s.Shapes) { $queue.Enqueue($sh) }
    while ($queue.Count -gt 0) {
      $sh = $queue.Dequeue()
      if ($sh.Type -eq 6) { foreach ($in in $sh.GroupItems) { $queue.Enqueue($in) }; continue }
      if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { [void]$ranges.Add(@('slide', $sh.TextFrame.TextRange, $sh.TextFrame.TextRange.Text)) }
      if ($sh.HasTable) {
        foreach ($row in $sh.Table.Rows) {
          $line = @()
          foreach ($cell in $row.Cells) { $tr = $cell.Shape.TextFrame.TextRange; $line += $tr.Text; [void]$ranges.Add(@('slide', $tr, $tr.Text)) }
          [void]$ranges.Add(@('slide', $null, ($line -join "`t")))
        }
      }
    }
    foreach ($sh in $s.NotesPage.Shapes) {
      if ($sh.HasTextFrame -and $sh.TextFrame.HasText -and $sh.PlaceholderFormat.Type -eq 2) { [void]$ranges.Add(@('notes', $sh.TextFrame.TextRange, $sh.TextFrame.TextRange.Text)) }
    }
    foreach ($r in $ranges) {
      $mt = [regex]::Match([string]$r[2], $pattern)
      if ($mt.Success) { $hit = @{ slide = $s.SlideIndex; where = $r[0]; range = $r[1]; m = $mt }; break }
    }
    if ($hit) { break }
  }
  $res = @{ found = $hit -ne $null; wasOpen = -not $opened }
  if ($hit) {
    $res.page = $hit.slide; $res.where = $hit.where
    if ($show) {
      # Open without a window (a check that was never closed): give it one
      if ($pres.Windows.Count -eq 0) { [void]$pres.NewWindow() }
      $win = $pres.Windows.Item(1); $win.Activate(); $win.View.GotoSlide($hit.slide)
      if ($hit.where -eq 'slide' -and $hit.range) { try { $hit.range.Characters($hit.m.Index + 1, $hit.m.Length).Select() } catch { $res.selectError = $true } }
    }
  }
  if ($show) { $res.raised = Raise $null 'POWERPNT' }
  elseif ($opened) { $pres.Saved = -1; $pres.Close(); if ($a.own) { $app.Quit() } }
  $res
}

# A number as somebody wrote it, in every way it can be read: "1 987",
# "1,987" and "1.987" are all 1987 to someone, and the last two are also
# 1.987. Each reading is { value, decimals, percent }; none when the text is
# no number. A typeset minus is a minus.
function Number([string]$t) {
  $c = (($t -replace '[\s\u00a0\u202f]', '') -replace '\u2212', '-') -replace '%$', ''
  $percent = $t.Trim().EndsWith('%')
  $inv = [Globalization.CultureInfo]::InvariantCulture
  $read = @()
  # Groups of three: no decimals, or decimals after the OTHER separator
  if ($c -match '^(-?\d{1,3})((?:[.,]\d{3})+)$' -or $c -match '^(-?\d{1,3})((?:\.\d{3})+),(\d+)$' -or $c -match '^(-?\d{1,3})((?:,\d{3})+)\.(\d+)$') {
    $head = $Matches[1]; $groups = $Matches[2]; $dec = ''; if ($Matches[3]) { $dec = $Matches[3] }
    # One separator between the groups, not both
    if (-not ($groups.Contains('.') -and $groups.Contains(','))) {
      $whole = $head + ($groups -replace '[.,]', '')
      $text = $whole; if ($dec -ne '') { $text = "$whole.$dec" }
      $read += @{ value = [double]::Parse($text, $inv); decimals = $dec.Length; percent = $percent }
    }
  }
  # One separator at most: the decimal sign
  if ($c -match '^-?\d+([.,](\d+))?$') {
    $dec = 0; if ($Matches[2]) { $dec = $Matches[2].Length }
    $read += @{ value = [double]::Parse(($c -replace ',', '.'), $inv); decimals = $dec; percent = $percent }
  }
  , $read
}
function Show-Excel {
  # Excel runs as several instances, and one of them can be an invisible one
  # some other program holds. The workbook is used WHERE IT IS OPEN; a closed
  # one is opened in an Excel that is on screen, and only with none of those
  # is a new Excel started.
  $app = $null; $wb = $null; $own = $false
  $files = [SywWin]::OpenFiles()
  foreach ($f in $files) { if ($f[0] -eq $req.path) { try { $x = $f[1].Application; if ($x.Name -eq 'Microsoft Excel') { $wb = $f[1]; $app = $x } } catch { } } }
  if (-not $wb -and -not $show) {
    # Only checked, never shown: in an Excel of this script's own, out of
    # sight. Opened in the Excel on screen the workbook would come up in
    # front of the reader and go away again.
    $app = New-Object -ComObject Excel.Application; $app.DisplayAlerts = $false; $own = $true
  } elseif (-not $wb) {
    foreach ($f in $files) { if (-not $app) { try { $x = $f[1].Application; if ($x.Name -eq 'Microsoft Excel' -and $x.Visible) { $app = $x } } catch { } } }
    if (-not $app) { try { $x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); if ($x.Visible) { $app = $x } } catch { } }
    if (-not $app) { $app = New-Object -ComObject Excel.Application; $own = $true }
  }
  $a = @{ app = $app; own = $own }
  $opened = $wb -eq $null
  if ($opened) { $wb = $app.Workbooks.Open($req.path, 0, (-not $show)) }
  $res = @{ found = $false; wasOpen = -not $opened }
  # The place is a cell ("Sheet!B2"), a row, a column or a block ("Sheet!5:5",
  # "Sheet!C:C", "Sheet!B2:D10"), or a sheet alone ("Sheet!" or its name).
  # A cell is checked against the value cited; a row, column or block holds
  # the citation when ONE of its cells does; a sheet alone only has to exist.
  $at = $req.cell.LastIndexOf('!')
  $claim = Number $req.quote
  $norm = { param($t) (($t -replace '[\s\u00a0\u202f]', '') -replace '\u2212', '-') }
  # A cell holds the value cited when it SHOWS it (as text, blanks aside), or
  # when its number is the number cited, in one of the ways that can be read,
  # to as many decimals as were cited. A cell of a merged block is the block.
  $same = {
    param($c)
    if ($c.MergeCells) { $c = $c.MergeArea.Cells.Item(1, 1) }
    if ((& $norm ([string]$c.Text)) -eq (& $norm ([string]$req.quote))) { return $true }
    if ($c.Value2 -is [double]) {
      foreach ($one in $claim) {
        $v = $c.Value2; if ($one.percent) { $v = $v * 100 }
        if ([math]::Abs($v - $one.value) -lt (0.5 * [math]::Pow(10, -$one.decimals) + 1e-9)) { return $true }
      }
    }
    $false
  }
  # A sheet's name as Excel writes it in a formula: in quotation marks when it
  # has a space, and an apostrophe of its own doubled
  $sheetOf = { param([string]$n) $n = $n.Trim(); if ($n -match "^'(.*)'$") { $Matches[1] -replace "''", "'" } else { $n } }
  try {
    $addr = $req.cell.Substring($at + 1).Trim()
    if ($at -gt 0) { $ws = $wb.Worksheets.Item((& $sheetOf $req.cell.Substring(0, $at))) }
    else { try { $ws = $wb.Worksheets.Item((& $sheetOf $req.cell)); $addr = '' } catch { $ws = $wb.Worksheets.Item(1) } }
    $res.sheet = $ws.Name
    $rng = $null
    if ($addr -eq '') { $res.found = $true; $res.whole = 'sheet' }
    else {
      $rng = $ws.Range($addr)
      if ($rng.Cells.CountLarge -eq 1) {
        if ($rng.MergeCells) { $rng = $rng.MergeArea.Cells.Item(1, 1) }
        $res.text = [string]$rng.Text
        $res.formula = [string]$rng.Formula
        $res.found = [bool](& $same $rng)
      } else {
        $res.whole = 'range'
        # The part of it that holds anything, worked out by hand: Excel's own
        # Intersect cannot be called with two arguments on an instance this
        # script started (measured: "Cannot find an overload"), only on one
        # that was already running
        $u = $ws.UsedRange; $used = $null
        $r1 = [Math]::Max($rng.Row, $u.Row); $r2 = [Math]::Min($rng.Row + $rng.Rows.Count - 1, $u.Row + $u.Rows.Count - 1)
        $c1 = [Math]::Max($rng.Column, $u.Column); $c2 = [Math]::Min($rng.Column + $rng.Columns.Count - 1, $u.Column + $u.Columns.Count - 1)
        if ($r1 -le $r2 -and $c1 -le $c2) { $used = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r2, $c2)) }
        $i = 0
        if ($used) { foreach ($c in $used.Cells) { if (++$i -gt 5000) { break }; if (& $same $c) { $res.found = $true; $res.text = [string]$c.Text; $res.at = $c.Address($false, $false); break } } }
      }
    }
    if ($show) {
      if (-not $app.Visible) { $app.Visible = $true }
      foreach ($w in $wb.Windows) { if (-not $w.Visible) { $w.Visible = $true } }
      $wb.Activate(); $ws.Activate()
      # A row, column or block is selected whole, and the view goes to the
      # cell in it that holds the value: a column's top can be far from it
      # Scrolled only as far as it takes to bring the cell into view: scrolled
      # to the corner, a cell loses the row labels and headings that say what it is
      if ($rng -and $res.at) { $app.Goto($ws.Range($res.at), $false); [void]$rng.Select(); [void]$ws.Range($res.at).Activate() }
      elseif ($rng) { $app.Goto($rng, $false) }
    }
  } catch { $res.error = 'no such sheet or cell' }
  if ($show) { $res.raised = Raise $app.Hwnd 'EXCEL' }
  elseif ($opened) { $wb.Close($false); if ($a.own) { $app.Quit() } }
  $res
}

# A text file's text, whatever it is written in: by its mark when it has one
# (UTF-8, UTF-16), else as UTF-8 when it IS that, else in the machine's own
# code page - which is what a file with a Norwegian letter in it is when it
# is not UTF-8 (read as UTF-8 regardless, those letters come out as nothing
# and no quote with one is ever found).
function ReadText([string]$path) {
  $bytes = [IO.File]::ReadAllBytes($path)
  $marked = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -or
    ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)))
  if ($marked) { return [IO.File]::ReadAllText($path) }
  try { (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes) }
  catch { [Text.Encoding]::Default.GetString($bytes) }
}

# A plain text file (Markdown, a note, a log): the words are looked up in the
# file itself and the place is a LINE. Shown in Notepad, which takes no
# "go to" on its command line: its text is reached through UI Automation,
# the way a screen reader reaches it, and the words selected there.
function Show-Text {
  $text = ReadText $req.path
  $pattern = Pattern $req.quote
  $mt = [regex]::Match($text, $pattern)
  if (-not $mt.Success -and @('.md', '.markdown', '.qmd', '.rmd') -contains [IO.Path]::GetExtension($req.path).ToLower()) {
    # Markdown is quoted as it READS: without the marks for bold, italics and
    # code, and a link by its words. Tried only when the words as written are
    # not there. Nothing is taken out but inside a line, so the line holds.
    $plain = [regex]::Replace($text, '!?\[([^\]\r\n]*)\]\([^)\r\n]*\)', '$1')
    $plain = [regex]::Replace($plain, '\*+(?=\S)|(?<=\S)\*+|`+|(?<![\w\\])_+(?=\S)|(?<=\S)_+(?!\w)', '')
    $again = [regex]::Match($plain, $pattern)
    if ($again.Success) { $mt = $again; $text = $plain }
  }
  $res = @{ found = $mt.Success }
  if ($mt.Success) { $res.page = ($text.Substring(0, $mt.Index) -split "`n").Count }
  if (-not $show) { return $res }
  # A code file goes to VS Code when it is installed: it takes a line and a
  # column on its command line, which Notepad does not. The cursor lands at
  # the first word; VS Code's command line cannot select a range.
  $prose = @('.md', '.markdown', '.txt', '.text', '.log')
  $code = Get-Command code -ErrorAction SilentlyContinue
  if ($code -and ($prose -notcontains [IO.Path]::GetExtension($req.path).ToLower())) {
    $line = 1; $col = 1
    if ($mt.Success) { $line = $res.page; $col = $mt.Index - ($text.LastIndexOf("`n", [Math]::Max($mt.Index - 1, 0)) + 1) + 1 }
    # The program itself, not the "code" command beside it: the command is a
    # script that stays until VS Code has answered, and whoever reads this
    # script's output waits with it (measured: 42 s, and over two minutes
    # through a pipe; the program itself returns in 0.2 s)
    $exe = Join-Path (Split-Path (Split-Path $code.Source)) 'Code.exe'
    if (-not (Test-Path $exe)) { $exe = $code.Source }
    Start-Process -FilePath $exe -ArgumentList '-g', ('"' + $req.path + ':' + $line + ':' + $col + '"')
    $res.editor = 'code'
    for ($i = 0; $i -lt 25; $i++) { Start-Sleep -Milliseconds 200; if (Get-Process Code -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }) { break } }
    $res.raised = Raise $null 'Code'
    return $res
  }
  $name = [IO.Path]::GetFileName($req.path)
  $find = { Get-Process notepad -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "$name*" } | Select-Object -First 1 }
  $p = & $find
  $res.wasOpen = $p -ne $null
  if (-not $p) {
    Start-Process notepad.exe -ArgumentList ('"' + $req.path + '"')
    for ($i = 0; $i -lt 40 -and -not $p; $i++) { Start-Sleep -Milliseconds 200; $p = & $find }
  }
  if (-not $p) { $res.error = 'Notepad did not open the file'; return $res }
  $h = $p.MainWindowHandle
  $res.raised = Raise $h 'notepad'
  if ($mt.Success) {
    try {
      Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
      $root = [Windows.Automation.AutomationElement]::FromHandle($h)
      $has = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::IsTextPatternAvailableProperty, $true)
      $el = $null
      for ($i = 0; $i -lt 20 -and -not $el; $i++) { $el = $root.FindFirst([Windows.Automation.TreeScope]::Descendants, $has); if (-not $el) { Start-Sleep -Milliseconds 150 } }
      $doc = $el.GetCurrentPattern([Windows.Automation.TextPattern]::Pattern).DocumentRange
      $seen = [regex]::Match($doc.GetText(-1), $pattern)
      if ($seen.Success) {
        $start = [Windows.Automation.Text.TextPatternRangeEndpoint]::Start
        $end = [Windows.Automation.Text.TextPatternRangeEndpoint]::End
        $char = [Windows.Automation.Text.TextUnit]::Character
        $r = $doc.Clone()
        $r.MoveEndpointByRange($end, $r, $start)
        [void]$r.MoveEndpointByUnit($end, $char, $seen.Index + $seen.Length)
        [void]$r.MoveEndpointByUnit($start, $char, $seen.Index)
        # Notepad shows Markdown formatted, and a step of one "character" there
        # is not always one character of its text (a list bullet counts and is
        # no text: measured three off after three bullets). So each end is
        # checked against the text in front of it and moved until it agrees.
        foreach ($edge in @($start, $end)) {
          $want = $seen.Index; if ($edge -eq $end) { $want += $seen.Length }
          for ($i = 0; $i -lt 8; $i++) {
            $before = $doc.Clone(); $before.MoveEndpointByRange($end, $r, $edge)
            $off = $want - $before.GetText(-1).Length
            if ($off -eq 0) { break }
            [void]$r.MoveEndpointByUnit($edge, $char, $off)
          }
        }
        $r.Select(); $r.ScrollIntoView($true)
        $res.selected = $r.GetText(80)
      } else { $res.selected = $null }
    } catch { $res.selectError = $_.Exception.Message.Trim() }
  }
  $res
}

# An HTML file, or a web page. The file's words are looked up in its own text
# (tags dropped, the common entities read); a web page is never fetched, so
# its words are never checked. Either is shown in the default browser, which
# scrolls to the words and marks them itself (a text fragment, "#:~:text=").
# The shell drops everything after "#" from a file address, so the browser is
# started by its own program with the address as its argument.
function Show-Html {
  $isWeb = $req.kind -eq 'web'
  $res = @{ found = $false }
  if (-not $isWeb) {
    # The page's text as it reads: no comment, script, style sheet or head; a
    # tag that sets text apart (a paragraph, a line break, a cell, a list
    # item) is a blank, and one inside a sentence or a word (bold, a link) is
    # nothing at all - as a blank it parted a word in two and put a space in
    # front of a full stop. Then every entity, named or numbered, is read.
    $raw = ReadText $req.path
    $text = $raw -replace '(?s)<!--.*?-->', ' ' -replace '(?is)<(script|style|head|template)\b.*?</\1\s*>', ' '
    $inline = 'a|abbr|b|bdi|bdo|cite|code|del|dfn|em|font|i|ins|kbd|mark|q|s|samp|small|span|strike|strong|sub|sup|tt|u|var|wbr'
    $text = $text -replace "(?is)</?(?:$inline)(?:\s[^>]*)?>", '' -replace '(?s)<[^>]+>', ' '
    $text = [Net.WebUtility]::HtmlDecode($text)
    $res.found = [regex]::IsMatch($text, (Pattern $req.quote))
  }
  if (-not $show) { return $res }
  $words = @($req.quote.Trim() -split '\s+')
  $enc = { param($t) [uri]::EscapeDataString($t).Replace('-', '%2D') }
  if ($words.Count -le 12) { $frag = & $enc ($words -join ' ') }
  else { $frag = (& $enc ($words[0..4] -join ' ')) + ',' + (& $enc ($words[($words.Count - 5)..($words.Count - 1)] -join ' ')) }
  if ($isWeb) { $url = ($req.path -split '#')[0] + '#:~:text=' + $frag }
  else { $url = ([uri]$req.path).AbsoluteUri + '#:~:text=' + $frag }
  $exe = $null
  try {
    $id = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice').ProgId
    $cmd = (Get-ItemProperty ("Registry::HKEY_CLASSES_ROOT\$id\shell\open\command")).'(default)'
    if ($cmd -match '^"([^"]+)"') { $exe = $Matches[1] } elseif ($cmd -match '^(\S+)') { $exe = $Matches[1] }
  } catch { }
  if ($exe -and (Test-Path $exe)) { Start-Process $exe -ArgumentList ('"' + $url + '"'); $res.browser = [IO.Path]::GetFileNameWithoutExtension($exe) }
  elseif ($isWeb) { Start-Process $url; $res.browser = 'default' }
  else { Start-Process $req.path; $res.browser = 'default, without the words' }
  $res
}

# A window of a program that showed something on its own word (PDF Scholar,
# through its own server): brought to the front the way the others are
function Show-Raise { @{ found = $true; raised = Raise $null $req.process } }

try {
  $res = switch ($req.kind) { 'word' { Show-Word } 'powerpoint' { Show-PowerPoint } 'excel' { Show-Excel } 'text' { Show-Text } 'html' { Show-Html } 'web' { Show-Html } 'raise' { Show-Raise } default { @{ error = 'unknown kind' } } }
} catch { $res = @{ error = $_.Exception.Message.Trim() } }
$res | ConvertTo-Json -Compress
