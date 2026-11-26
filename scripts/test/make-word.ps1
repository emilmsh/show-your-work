# Builds the Word document the look-up is tested on: a report as reports are -
# a table of contents, cross-references, footnotes, a table, a link, a tracked
# deletion, a comment, hidden text - with one marked sentence after each of
# them. Each of those puts characters in the document that its text does not
# show, which is what a look-up by position gets wrong.
#
#   make-word.ps1 <out.docx>
#
# Writes <out.docx> and, beside it, <out.docx>.json: every marked sentence
# with the page WORD ITSELF finds it on (its own Find, not our look-up).
param([Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $marks = New-Object Collections.Generic.List[string]
  $filler = 'The agency pays the benefit monthly, and each payment is settled against the register at the end of the quarter. '

  function Para([string]$text) { $sel.TypeText($text); $sel.TypeParagraph() }
  function Mark([string]$what) {
    $text = "Marked sentence $($marks.Count + 1), after $what, says that the rate was $(11 + $marks.Count * 7).$($marks.Count) per cent."
    $marks.Add($text); Para $text
  }
  function Body([int]$paragraphs) { for ($i = 0; $i -lt $paragraphs; $i++) { Para ($filler * 6) } }
  function Heading([string]$text) { $sel.Style = $doc.Styles.Item(-2); Para $text; $sel.Style = $doc.Styles.Item(-1) }

  Heading 'Contents'
  $toc = $sel.Range
  $sel.TypeParagraph()
  [void]$sel.InsertBreak(7)

  Heading 'Background'
  Mark 'nothing at all'
  Body 6

  Heading 'Footnotes'
  for ($i = 0; $i -lt 12; $i++) {
    $sel.TypeText('A claim with a note')
    [void]$doc.Footnotes.Add($sel.Range, [Type]::Missing, "Note $i says where the claim is from, at some length, so that it takes its room in the text.")
    $sel.TypeText('. ' + $filler)
  }
  $sel.TypeParagraph()
  Mark 'twelve footnotes'
  Body 5

  Heading 'Cross-references'
  for ($i = 0; $i -lt 15; $i++) {
    $sel.TypeText('See page ')
    [void]$sel.Fields.Add($sel.Range, -1, 'PAGE', $false)
    $sel.TypeText(' of ')
    [void]$sel.Fields.Add($sel.Range, -1, 'NUMPAGES', $false)
    $sel.TypeText(', written ')
    [void]$sel.Fields.Add($sel.Range, -1, 'CREATEDATE \@ "d MMMM yyyy"', $false)
    $sel.TypeText('. ' + $filler)
  }
  $sel.TypeParagraph()
  Mark 'forty-five fields'
  Body 5

  Heading 'A table'
  $table = $doc.Tables.Add($sel.Range, 6, 4)
  for ($r = 1; $r -le 6; $r++) { for ($c = 1; $c -le 4; $c++) { $table.Cell($r, $c).Range.Text = "row $r, column $c" } }
  $table.Cell(3, 2).Range.Text = 'Marked cell: 4 812 recipients in the third row.'
  $marks.Add('Marked cell: 4 812 recipients in the third row.')
  [void]$sel.EndKey(6)
  $sel.TypeParagraph()
  Mark 'a table'
  Body 5

  Heading 'Links'
  for ($i = 0; $i -lt 10; $i++) {
    $sel.TypeText('The source is ')
    [void]$doc.Hyperlinks.Add($sel.Range, "https://example.org/a/long/address/that/is/not/shown/$i", [Type]::Missing, [Type]::Missing, "report $i")
    [void]$sel.EndKey(6)
    $sel.TypeText('. ' + $filler)
  }
  $sel.TypeParagraph()
  Mark 'ten links'
  Body 5

  Heading 'Changes'
  Para ('This sentence is to be deleted with the change kept on record. ' * 8)
  $gone = $doc.Paragraphs.Item($doc.Paragraphs.Count - 1).Range
  Para 'A sentence somebody has a comment on.'
  [void]$doc.Comments.Add($doc.Paragraphs.Item($doc.Paragraphs.Count - 1).Range, 'A comment in the margin, which is no part of the text.')
  Para ('This sentence is hidden text and is not printed. ' * 8)
  $doc.Paragraphs.Item($doc.Paragraphs.Count - 1).Range.Font.Hidden = $true
  $sel.Font.Hidden = $false
  Mark 'a tracked deletion, a comment and hidden text'
  Body 4

  Heading 'The end'
  Mark 'everything'
  $sel.TypeText('A claim whose figure is in its note')
  $note = $doc.Footnotes.Add($sel.Range, [Type]::Missing, 'Marked footnote: the figure is 73.9 per cent, from the register.')
  $sel.TypeText('.'); $sel.TypeParagraph()
  Para 'The same words twice. The same words twice.'
  Para "A soft hyphen in co$([char]0xAD)operation, a no-break space in 12$([char]0xA0)000, a curly $([char]0x201C)quote$([char]0x201D) and a line$([char]0x0B)break."
  $marks.Add('A soft hyphen in cooperation')
  $marks.Add("a no-break space in 12 000, a curly $([char]0x201C)quote$([char]0x201D) and a line break.")

  $doc.TrackRevisions = $true
  $gone.Delete() | Out-Null
  $doc.TrackRevisions = $false

  [void]$doc.TablesOfContents.Add($toc, $true, 1, 1)
  [void]$doc.Fields.Update()
  $doc.Repaginate()

  # Where Word itself finds each sentence: its own Find, forty characters at a
  # time (a Find takes 255 at most), from the top
  $want = foreach ($text in $marks) {
    $r = $doc.Content
    $r.Find.ClearFormatting()
    $head = $text.Substring(0, [Math]::Min(40, $text.Length))
    $ok = $r.Find.Execute($head, $true, $false, $false, $false, $false, $true, 0)
    @{ quote = $text; page = if ($ok) { [int]$r.Information(3) } else { 0 } }
  }

  $want = @($want) + @{ quote = 'Marked footnote: the figure is 73.9 per cent, from the register.'; page = [int]$note.Range.Information(3) }

  $doc.SaveAs2($Out, 12)
  $pages = $doc.ComputeStatistics(2)
  $doc.Close(0)
  @{ pages = $pages; marks = @($want) } | ConvertTo-Json -Depth 4 | Set-Content -Path "$Out.json" -Encoding UTF8
  "$pages pages, $(@($want).Count) marked passages"
} finally {
  $word.Quit([ref]0)
}
