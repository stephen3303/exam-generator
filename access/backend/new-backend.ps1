# Makes a new, empty shared data file (ExamGen_BE.accdb) with modBuildBackEnd.bas:
# every table, relationship and index, and the seed lookup lists.
#
#   backend\new-backend.ps1                  writes build\ExamGen_BE.accdb (for testing)
#   backend\new-backend.ps1 -Path <file>     writes it there (a full path to a new .accdb)
#
# It never overwrites: if the file is there already it stops and changes nothing.
# Then link a front end to it with  .\build.ps1 -BackEnd <file>.
# Requires Microsoft Access (2016+ / Microsoft 365).
param([string]$Path = '')
$ErrorActionPreference = 'Stop'
$backend = Split-Path -Parent $MyInvocation.MyCommand.Path
$here = Split-Path -Parent $backend
if (-not $Path) {
    New-Item -ItemType Directory -Force (Join-Path $here 'build') | Out-Null
    $Path = Join-Path $here 'build\ExamGen_BE.accdb'
}
$Path = [IO.Path]::GetFullPath($Path)
if (Test-Path $Path) { throw "There is already a file at $Path; nothing was changed." }

# The builder runs inside a scratch database that is deleted afterwards.
$scratch = Join-Path $env:TEMP 'ExamGen_backend_builder.accdb'
Remove-Item $scratch -ErrorAction SilentlyContinue

. (Join-Path $here 'tools\AccessWatch.ps1')
$acc = New-Object -ComObject Access.Application
$watch = Start-AccessWatch $acc -TimeoutSec 300
$failed = $false
try {
    $acc.Visible = $false
    $acc.AutomationSecurity = 1
    $acc.NewCurrentDatabase($scratch)
    # LoadFromText expects Windows line endings.
    $tmp = Join-Path $env:TEMP 'examgen_modBuildBackEnd.bas'
    $text = [IO.File]::ReadAllText((Join-Path $backend 'modBuildBackEnd.bas')) -replace "`r?`n", "`r`n"
    [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
    $acc.LoadFromText(5, 'modBuildBackEnd', $tmp)
    Remove-Item $tmp

    $bad = Test-AccessCompiles $acc $watch
    if ($bad) { Write-Host "Compile error: $bad"; throw 'compile failed' }

    $a1 = [string]$Path
    $result = $acc.Run('BuildBackEndAt', [ref]$a1)
    Write-Host "BuildBackEndAt: $result"
    if ($result -ne 'OK') { $failed = $true }
}
catch { $failed = $true; if (-not $watch.Killed) { throw } }
finally {
    $dialogs = Stop-AccessWatch $watch $acc
    Remove-Item $scratch -ErrorAction SilentlyContinue
}
if ($dialogs) { Write-Host "Stopped: Access showed the dialog(s) above."; exit 1 }
if ($failed) { exit 1 }
Write-Host "Created: $Path"
