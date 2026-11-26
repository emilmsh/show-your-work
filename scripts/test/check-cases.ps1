# Runs the look-up (scripts/show.ps1) on every case of a folder made by
# make_hard.py: words that are there must be found at the place they are, and
# words that are not must be refused. Nothing is shown on screen.
#
#   check-cases.ps1 <folder> [-Kind excel]      exit code = cases that failed
param([Parameter(Mandatory = $true)][string]$Folder, [string]$Kind = '')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$show = Join-Path (Split-Path $PSScriptRoot) 'show.ps1'
$cases = Get-Content (Join-Path $Folder 'cases.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$failed = 0; $ran = 0
foreach ($c in $cases) {
  if ($Kind -ne '' -and $c.kind -ne $Kind) { continue }
  $ran++
  $env:SYW_REQUEST = @{ kind = $c.kind; path = (Join-Path $Folder $c.file); quote = $c.quote; cell = $c.cell; show = $false } | ConvertTo-Json -Compress
  $answer = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $show | Select-Object -Last 1) | ConvertFrom-Json
  $found = $answer.found -eq $true
  $ok = $found -eq $c.found -and ($c.page -eq 0 -or $answer.page -eq $c.page)
  if ($ok) { continue }
  $failed++
  "FAIL  $($c.file) $($c.cell)  [$($c.why)]"
  "      quote: $($c.quote)"
  "      found=$found (want $($c.found))  place=$($answer.page) (want $($c.page))  cell shows: $($answer.text)  $($answer.error)"
}
"$failed of $ran failed"
exit $failed
