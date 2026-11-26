# Runs the look-up (scripts/show.ps1) on the demo set, one kind of file after
# the other: words that are there must be found at the place they are, and
# words that are NOT there must be refused. Nothing is shown on screen.
#
#   check-kinds.ps1                    exit code 0 when every case holds
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$root = Split-Path (Split-Path $PSScriptRoot)
$show = Join-Path $root 'scripts\show.ps1'
$demo = Join-Path $root 'demo'

# kind, file, quote, cell, found?, page (0 = not asked)
$cases = @(
  , @('powerpoint', 'slides.pptx', 'Attendance rose by 4.2 percentage points.', '', $true, 3)
  , @('powerpoint', 'slides.pptx', 'Difference-in-differences with matched control schools.', '', $true, 2)
  , @('powerpoint', 'slides.pptx', 'Recommendation', '', $true, 4)
  , @('powerpoint', 'slides.pptx', 'the cost is 4,800 kroner per pupil per year', '', $true, 4)
  , @('powerpoint', 'slides.pptx', 'Attendance rose by 5.2 percentage points.', '', $false, 0)
  , @('excel', 'model.xlsx', '1987', 'Assumptions!B2', $true, 0)
  , @('excel', 'model.xlsx', '1 987', 'Assumptions!B2', $true, 0)
  , @('excel', 'model.xlsx', '9 537 600', 'Results!B2', $true, 0)
  , @('excel', 'model.xlsx', '25,26', 'Results!B3', $true, 0)
  , @('excel', 'model.xlsx', '4.2', 'Results!B:B', $true, 0)
  , @('excel', 'model.xlsx', 'Results', 'Results!', $true, 0)
  , @('excel', 'model.xlsx', '1988', 'Assumptions!B2', $false, 0)
  , @('excel', 'model.xlsx', '4.3', 'Results!B:B', $false, 0)
  , @('excel', 'model.xlsx', '1987', 'Nowhere!B2', $false, 0)
  , @('text', 'notes.md', 'We have attendance records for all 160 schools, and the match between pilot and control schools holds up.', '', $true, 7)
  , @('text', 'notes.md', 'Attendance rose most among pupils from low-income families.', '', $true, 12)
  , @('text', 'notes.md', 'We have attendance records for all 170 schools', '', $false, 0)
  , @('text', 'analysis.R', 'filter(n_distinct(year) == 7)', '', $true, 10)
  , @('text', 'analysis.R', 'attendance <- read.csv("attendance.csv")', '', $true, 5)
  , @('text', 'analysis.R', 'filter(n_distinct(year) == 8)', '', $false, 0)
)
$failed = 0
foreach ($c in $cases) {
  $env:SYW_REQUEST = @{ kind = $c[0]; path = (Join-Path $demo $c[1]); quote = $c[2]; cell = $c[3]; show = $false } | ConvertTo-Json -Compress
  $answer = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $show | Select-Object -Last 1) | ConvertFrom-Json
  $found = $answer.found -eq $true
  $ok = $found -eq $c[4] -and ($c[5] -eq 0 -or $answer.page -eq $c[5])
  if (-not $ok) { $failed++ }
  $verdict = 'FAIL'; if ($ok) { $verdict = 'ok  ' }
  $short = $c[2].Substring(0, [Math]::Min(44, $c[2].Length))
  "$verdict  $($c[0]) $($c[3])  found=$found (want $($c[4]))  place=$($answer.page)$($answer.text) (want $($c[5]))  $short"
  if ($answer.error) { "      error: $($answer.error)" }
}
"$failed of $($cases.Count) failed"
exit $failed
