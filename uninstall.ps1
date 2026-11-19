<#
 Avinstallera Puppa på Windows (fungerar även i pwsh på macOS/Linux):
   irm https://raw.githubusercontent.com/pipebreak33/puppa/main/uninstall.ps1 | iex
 Från en lokal kopia:  ./uninstall.ps1

 Tar bort programmet (~/.puppa/app), kommandot puppa och PATH-posten.
 Behåller dina tenants, inställningar, certifikat och egna recept (~/.puppa) och rapporterna i ~/Documents/Puppa.
 $env:PUPPA_PURGE = '1' tar bort även .puppa med dina data. $env:PUPPA_YES = '1' hoppar över frågan.
#>
$ErrorActionPreference = 'Stop'
$Clock = [Diagnostics.Stopwatch]::StartNew()
$OnWindows = ($PSVersionTable.PSEdition -ne 'Core') -or $IsWindows
$Dir = Join-Path $HOME '.puppa'
$Bin = Join-Path $Dir 'bin'
$Shim = if ($OnWindows) { Join-Path $Bin 'puppa.cmd' } else { Join-Path $HOME '.local/bin/puppa' }

# Language: TRUPPA_LANG, else ~/.truppa/config.json (shared by the whole troupe), else English
$TL = $env:TRUPPA_LANG
if (-not $TL) { try { $TL = (Get-Content -Raw -LiteralPath (Join-Path $HOME '.truppa/config.json') -ErrorAction Stop | ConvertFrom-Json).lang } catch { } }
function T([string]$en, [string]$sv) { if ("$TL" -like 'sv*') { $sv } else { $en } }
function Say($m) { Write-Host "  P› $m" -ForegroundColor Cyan }
function Ask($q) { $s = Read-Host "  $q $(T '[y/N]' '[j/N]')"; return ($s -match '^(j|ja|y|yes)$') }

$targets = @()
if ($env:PUPPA_PURGE -eq '1') {
    if (Test-Path $Dir) { $targets += $Dir }
} else {
    foreach ($p in @('app', 'bin')) { $x = Join-Path $Dir $p; if (Test-Path $x) { $targets += $x } }
}
# startkommandot på macOS/Linux: bara om det är vårt (pekar på ~/.puppa)
if (-not $OnWindows -and (Test-Path $Shim) -and ((Get-Content $Shim -Raw) -like "*/.puppa/*")) { $targets += $Shim }
$user = if ($OnWindows) { [Environment]::GetEnvironmentVariable('Path', 'User') } else { '' }
if (-not $user) { $user = '' }
$inPath = ($user -split ';') -contains $Bin

if (-not $targets -and -not $inPath) { Say (T 'Puppa is not installed here. Nothing to remove.' 'Puppa är inte installerad här. Inget att ta bort.'); return }
Write-Host (T "`n  This is removed:" "`n  Det här tas bort:")
$targets | ForEach-Object { Write-Host "    $_" }
if ($inPath) { Write-Host "    $Bin i din PATH" }
if ($env:PUPPA_PURGE -ne '1') { Write-Host (T '  Kept: your tenants, settings, certificates and own recipes (~/.puppa) and the reports in ~/Documents/Puppa.' '  Behålls: dina tenants, inställningar, certifikat och egna recept (~/.puppa) och rapporterna i ~/Documents/Puppa.') }
Write-Host ''
if ($env:PUPPA_YES -ne '1') {
    if (-not (Ask (T 'Uninstall Puppa?' 'Avinstallera Puppa?'))) { Say (T 'Cancelled. Nothing has been removed.' 'Avbrutet. Ingenting har tagits bort.'); return }
}

foreach ($t in $targets) { Remove-Item $t -Recurse -Force }
if ($inPath) {
    [Environment]::SetEnvironmentVariable('Path', (($user -split ';' | Where-Object { $_ -and $_ -ne $Bin }) -join ';'), 'User')
}
if ((Test-Path $Dir) -and -not (Get-ChildItem $Dir -Force)) { Remove-Item $Dir -Force }

$Clock.Stop()
Write-Host (T "`n  ✔ Puppa is uninstalled ($([int]$Clock.Elapsed.TotalSeconds) s).`n" "`n  ✔ Puppa är avinstallerad ($([int]$Clock.Elapsed.TotalSeconds) s).`n") -ForegroundColor Green
