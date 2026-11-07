# Fills a NEW, empty Exam Generator data file with made-up demo data, for testing the
# pages and the Word documents: people, topics, templates, references, questions in every
# state, and five exams, two of them approved with their Word documents made.
# Every name, reference and question is invented. See access\SEED-DEMO.md.
#
#   tools\seed-demo.ps1                              seeds build\ExamGen_BE.accdb
#   tools\seed-demo.ps1 -BackEnd <file>              seeds that data file
#   -OwnerUser <name>     your sign-in name in the Personnel file, when your Windows
#                         account isn't saved there (found by Windows account otherwise)
#   -OwnerFirst, -OwnerLast, -OwnerRank
#                         your name, used only when there is no Personnel file yet and
#                         this makes one (in access\Personnel, with you as its admin)
#   -Personnel <file>     the Personnel file to use. Default: Personnel\Personnel_be.accdb
#                         beside the data file when there is one (the web's Database
#                         Files folder), else the usual place (a Personnel folder
#                         beside the data file's folder, made there when missing)
#   -DemoPassword <text>  every made-up person's app password (default ExamDemo-2026)
#   -NoDocs               skip Word (the documents are then made by Download a copy)
#   -NoLineNumbers        load the code without line numbers
#
# It refuses a data file that already has questions or exams. It drives the app's own
# code, signed in as each made-up person in turn, so the app's rules apply to every step.
# The code runs in a throwaway database in TEMP: the front end is not changed, and works
# with the seeded data file as it is (build.ps1 -BackEnd links it).
# Requires Microsoft Access and Word (2016+ / Microsoft 365).
param(
    [string]$BackEnd = '',
    [string]$OwnerUser = '',
    [string]$OwnerFirst = 'Site',
    [string]$OwnerLast = 'Owner',
    [string]$OwnerRank = '',
    [string]$Personnel = '',
    [string]$DemoPassword = 'ExamDemo-2026',
    [switch]$NoDocs,
    [switch]$NoLineNumbers
)
$ErrorActionPreference = 'Stop'
$tools = Split-Path -Parent $MyInvocation.MyCommand.Path
$here = Split-Path -Parent $tools
if (-not $BackEnd) { $BackEnd = Join-Path $here 'build\ExamGen_BE.accdb' }
if (-not (Test-Path $BackEnd)) { throw "No data file at $BackEnd. Make one first with backend\new-backend.ps1." }
$BackEnd = (Resolve-Path $BackEnd).Path
$lock = [IO.Path]::ChangeExtension($BackEnd, '.laccdb')
if (Test-Path $lock) { throw "The data file is open ($lock exists). Close every Access window and try again." }
if (-not $Personnel) {
    $beside = Join-Path (Split-Path -Parent $BackEnd) 'Personnel\Personnel_be.accdb'
    if (Test-Path $beside) { $Personnel = $beside }
}
if ($Personnel) {
    if (-not (Test-Path $Personnel)) { throw "No Personnel file at $Personnel." }
    $Personnel = (Resolve-Path $Personnel).Path
    $plock = [IO.Path]::ChangeExtension($Personnel, '.laccdb')
    if (Test-Path $plock) { throw "The Personnel file is open ($plock exists). Stop the web server and close every Access window, then try again." }
}
if ($DemoPassword.Length -lt 8) { throw 'Use at least 8 characters for -DemoPassword.' }

# The seed works on a copy beside the data file and only takes its place once it
# finished, so a failed run leaves the data file as it was (and still empty).
$copy = Join-Path (Split-Path -Parent $BackEnd) ([IO.Path]::GetFileNameWithoutExtension($BackEnd) + '.seeding.accdb')

# The runner: a throwaway database with the kit's and the app's modules (no forms, so
# no start-up runs) and this seed module.
$work = Join-Path $env:TEMP ('ExamGen_seed_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null
$runner = Join-Path $work 'seed-runner.accdb'
. (Join-Path $tools 'VbaLineNumbers.ps1')
. (Join-Path $tools 'AccessWatch.ps1')

$acc = $null
$watch = $null
$dialogs = $null
$failed = $true
try {
    Remove-Item $copy -ErrorAction SilentlyContinue
    Copy-Item $BackEnd $copy
    $acc = New-Object -ComObject Access.Application
    $watch = Start-AccessWatch $acc -TimeoutSec 1200
    $acc.Visible = $false
    $acc.AutomationSecurity = 1   # let the seed's VBA run
    $acc.NewCurrentDatabase($runner)

    $files = @(Get-ChildItem (Join-Path $here 'kit\modules') -Filter *.bas) +
             @(Get-Item (Join-Path $here 'kit\adapters\modTextMeasure.bas')) +
             @(Get-ChildItem (Join-Path $here 'src\modules') -Filter *.bas) +
             @(Get-Item (Join-Path $here 'src\themes\light\modTheme.bas')) +
             @(Get-Item (Join-Path $tools 'modSeedDemo.bas'))
    foreach ($f in $files) {
        # LoadFromText expects Windows line endings and is given ASCII, as in build.ps1.
        $tmp = Join-Path $work ('seed_' + $f.Name)
        $text = [IO.File]::ReadAllText($f.FullName) -replace "`r?`n", "`r`n"
        if ($text -match '[^\x00-\x7F]') { throw "$($f.Name) has a character outside ASCII; write it with ChrW() instead." }
        if (-not $NoLineNumbers) { $text = Add-VbaLineNumbers $text }
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        $acc.LoadFromText(5, $f.BaseName, $tmp)
        Remove-Item $tmp
    }

    $bad = Test-AccessCompiles $acc $watch
    if ($bad) { throw "Compile error: $bad" }

    Write-Host "Seeding $BackEnd (this takes a few minutes; Word runs hidden for the documents)..."
    $a1 = [string]$copy; $a2 = [string]$OwnerUser; $a3 = [string]$OwnerFirst; $a4 = [string]$OwnerLast
    $a5 = [string]$OwnerRank; $a6 = [string]$DemoPassword; $a7 = [bool](-not $NoDocs); $a8 = [string]$Personnel
    if ($Personnel) { Write-Host "Personnel file: $Personnel" }
    $result = [string]$acc.Run('SeedDemoRun', [ref]$a1, [ref]$a2, [ref]$a3, [ref]$a4, [ref]$a5, [ref]$a6, [ref]$a7, [ref]$a8)
    $result -split "`r?`n" | Where-Object { $_ } | ForEach-Object { Write-Host $_ }
    $failed = -not $result.StartsWith('OK')
}
catch {
    # A dialog the watch closed is printed by Stop-AccessWatch below; print anything else.
    if (-not ($watch -and $watch.Killed)) { Write-Host "Stopped: $($_.Exception.Message)" }
}
finally {
    if ($watch) { $dialogs = Stop-AccessWatch $watch $acc }
    elseif ($acc) { try { $acc.Quit() } catch { } }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    if ($failed -or $dialogs) {
        Remove-Item $copy -ErrorAction SilentlyContinue
        Remove-Item ([IO.Path]::ChangeExtension($copy, '.laccdb')) -ErrorAction SilentlyContinue
    }
}
if ($dialogs -or $failed) {
    if ($dialogs) { Write-Host 'Stopped: Access showed the dialog(s) above.' }
    Write-Host "Seed did not finish cleanly, so $BackEnd was left as it was. Send the lines above."
    exit 1
}
Move-Item -Force $copy $BackEnd
Write-Host "Seeded: $BackEnd"
Write-Host "Made-up people sign in with the password $DemoPassword. Open build\ExamGenerator.accdb (built with -BackEnd on this file)."
