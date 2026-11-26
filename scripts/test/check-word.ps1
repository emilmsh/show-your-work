# Runs the look-up (scripts/show.ps1) on every marked passage of a document
# made by make-word.ps1, and holds its answer against Word's own: the page,
# and that the words at the place it found ARE the words asked for.
#
#   check-word.ps1 <doc.docx>          exit code 0 when every passage holds
param([Parameter(Mandatory = $true)][string]$Doc)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$show = Join-Path (Split-Path $PSScriptRoot) 'show.ps1'
$want = Get-Content "$Doc.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$failed = 0
foreach ($mark in $want.marks) {
  $env:SYW_REQUEST = @{ kind = 'word'; path = $Doc; quote = $mark.quote; cell = ''; show = $false } | ConvertTo-Json -Compress
  $answer = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $show | Select-Object -Last 1) | ConvertFrom-Json
  $ok = $answer.found -eq $true -and $answer.exact -eq $true -and ($mark.page -eq 0 -or $answer.page -eq $mark.page)
  if (-not $ok) { $failed++ }
  $verdict = 'FAIL'; if ($ok) { $verdict = 'ok  ' }
  $short = $mark.quote.Substring(0, [Math]::Min(46, $mark.quote.Length))
  "$verdict  page $($answer.page) (Word: $($mark.page))  exact=$($answer.exact)  $short"
  if (-not $ok -and $answer.at) { "      at the place: $($answer.at)" }
  if ($answer.error) { "      error: $($answer.error)" }
}
"$failed of $($want.marks.Count) failed"
exit $failed
