<#PUPPA
namn:        Vem har tagit bort saker på en sajt?
kategori:    SharePoint
typ:         läs
kräver:      exo
beskrivning: Borttagna filer och mappar ur granskningsloggen – vem (UPN), när, var och om det går att återställa. Valfri Excel-rapport.
param:       Sajt? | Sajt – skriv en del av namnet, eller klistra in en länk (även till bibliotek/mapp) (tomt = alla sajter) | | @site
param:       Dagar | Hur många dagar bakåt? (högst 180) | 30
param:       Namn? | Del av fil- eller mappnamnet (valfritt) |
param:       Vem? | Bara borttaget av (UPN, valfritt) | | @user
param:       Excel | Gör en Excel-rapport? | Ja | Ja,Nej
param:       Filnamn | Rapportens filnamn (utan .xlsx) | {kund}-borttaget-{datum}
#>
# Granskningsloggen nås via Exchange Online (Search-UnifiedAuditLog) – kräver rollen Audit Logs / View-Only Audit Logs.
$antal = [Math]::Min([Math]::Max([int]$Dagar, 1), 180)
$sajtUrl = ''
$sajtNamn = ''
$inom = ''
if ($Sajt) {
    $u = [uri]::UnescapeDataString(($Sajt.Trim().Trim('"') -replace '\?.*$', '' -replace '/Forms/[^/]+\.aspx$', '' -replace '/_layouts/.*$', ''))
    if ($u -match '^(https://[^/]+\.sharepoint\.com(/(sites|teams|personal)/[^/]+)?)(/.*)?$') {
        $sajtUrl = $Matches[1].TrimEnd('/') + '/'
        $inom = ($Matches[4] -replace '^/', '').TrimEnd('/')
    } else {
        # Inte en länk: matcha sajtens adress på namnet ("Signium AB" -> *Signium*AB*)
        $sajtNamn = '*' + (($Sajt.Trim() -replace '[^\w]+', '*').Trim('*')) + '*'
    }
}

# Vad händelserna betyder och om det går att få tillbaka
$handelse = @{
    FileRecycled                        = 'Till papperskorgen', 'Ja – sajtens papperskorg (93 dagar)'
    FolderRecycled                      = 'Till papperskorgen', 'Ja – sajtens papperskorg (93 dagar)'
    FileDeleted                         = 'Borttagen', 'Ja – sajtens papperskorg (93 dagar)'
    FolderDeleted                       = 'Borttagen', 'Ja – sajtens papperskorg (93 dagar)'
    FileDeletedFirstStageRecycleBin     = 'Tömd ur papperskorgen', 'Ja – andra stegets papperskorg (webbplatsadmin)'
    FolderDeletedFirstStageRecycleBin   = 'Tömd ur papperskorgen', 'Ja – andra stegets papperskorg (webbplatsadmin)'
    FileDeletedSecondStageRecycleBin    = 'Borta för gott', 'Nej – bara från säkerhetskopia'
    FolderDeletedSecondStageRecycleBin  = 'Borta för gott', 'Nej – bara från säkerhetskopia'
}

$slut = (Get-Date).ToUniversalTime(); $start = $slut.AddDays(-$antal)
$sok = @{ StartDate = $start; EndDate = $slut; RecordType = 'SharePointFileOperation'; Operations = [string[]]$handelse.Keys
          SessionId = [guid]::NewGuid().ToString(); SessionCommand = 'ReturnLargeSet'; ResultSize = 5000 }
if ($Vem) { $sok.UserIds = $Vem }
$rader = @(); $varv = 0
do {
    $sida = @(Search-UnifiedAuditLog @sok)
    $rader += $sida; $varv++
} while ($sida.Count -eq 5000 -and $varv -lt 10)

$sett = @{}
$borta = foreach ($r in $rader) {
    $d = $r.AuditData | ConvertFrom-Json
    if ($sajtUrl -and -not "$($d.SiteUrl)".StartsWith($sajtUrl, [StringComparison]::OrdinalIgnoreCase)) { continue }
    if ($sajtNamn -and -not (("$($d.SiteUrl)" -replace '^https://[^/]+', '') -like $sajtNamn)) { continue }
    $rel = ("$($d.SourceRelativeUrl)/$($d.SourceFileName)").TrimStart('/')
    if ($inom -and -not ($rel -like "$inom*")) { continue }
    if ($Namn -and -not ($d.SourceFileName -like "*$Namn*")) { continue }
    $nyckel = "$($d.Operation)|$($d.ObjectId)|$($d.UserId)|$($d.CreationTime)"
    if ($sett[$nyckel]) { continue }; $sett[$nyckel] = 1
    $bas = "$($d.SiteUrl)".TrimEnd('/')
    $h = $handelse["$($d.Operation)"]
    [pscustomobject]@{
        Tid                = ([datetime]$d.CreationTime).ToLocalTime()
        'Borttaget av'     = $d.UserId
        Händelse           = $h[0]
        Typ                = if ($d.Operation -like 'Folder*') { 'Mapp' } else { 'Fil' }
        Namn               = $d.SourceFileName
        Plats              = ("$bas/$($d.SourceRelativeUrl)".TrimEnd('/')) -replace ' ', '%20'
        Länk               = $(if ($d.ObjectId -match '^https://') { $d.ObjectId } else { "$bas/$rel" }) -replace ' ', '%20'
        Sajt               = $bas
        'Kan återställas?' = $h[1]
        'IP-adress'        = $d.ClientIP
    }
}
$borta = @($borta | Sort-Object Tid -Descending)

if (-not $borta) {
    Write-Host "Inget borttaget hittades$(if ($sajtUrl) { " på $sajtUrl" })$(if ($Vem) { " av $Vem" }) de senaste $antal dagarna."
    Write-Host 'Förslag: öka antalet dagar eller töm namnfiltret. Kolla också papperskorgen (receptet "Papperskorgen").'
    return
}

# Sammanställningar
$perPerson = $borta | Group-Object 'Borttaget av' | Sort-Object Count -Descending | ForEach-Object {
    [pscustomobject]@{
        'Borttaget av' = $_.Name; Totalt = $_.Count
        Filer = @($_.Group | Where-Object Typ -eq 'Fil').Count; Mappar = @($_.Group | Where-Object Typ -eq 'Mapp').Count
        'Borta för gott' = @($_.Group | Where-Object Händelse -eq 'Borta för gott').Count
        Första = ($_.Group.Tid | Measure-Object -Minimum).Minimum; Senaste = ($_.Group.Tid | Measure-Object -Maximum).Maximum
        Sajter = (($_.Group.Sajt | Sort-Object -Unique) -replace '^https://[^/]+', '') -join ', '
    }
}
$perDag = $borta | Group-Object { $_.Tid.ToString('yyyy-MM-dd') } | Sort-Object Name -Descending | ForEach-Object {
    [pscustomobject]@{ Datum = $_.Name; Antal = $_.Count
                       Personer = (($_.Group.'Borttaget av' | Group-Object | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ') }
}

Write-Host "$($borta.Count) borttagningar av $(@($perPerson).Count) personer de senaste $antal dagarna$(if ($sajtUrl) { " på $sajtUrl" }):"
foreach ($p in $perPerson | Select-Object -First 5) { Write-Host ("  {0,4} st  {1}" -f $p.Totalt, $p.'Borttaget av') }

if ($Excel -eq 'Ja') {
    # ImportExcel varnar om autosize på macOS/Linux. Modulfunktioner ser bara den globala inställningen.
    $tidigare = $global:WarningPreference; $global:WarningPreference = 'SilentlyContinue'
    try {
    Import-Module ImportExcel -ErrorAction Stop
    $dir = Join-Path $HOME 'Documents/Puppa'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $fil = ($Filnamn -replace '[\\/:*?"<>|]', '_').Trim(); if (-not $fil) { $fil = 'borttaget' }
    $path = Join-Path $dir "$fil.xlsx"
    if (Test-Path $path) { Remove-Item $path -Force }

    $om = @(
        [pscustomobject]@{ Fält = 'Rapport'; Värde = 'Borttagna filer och mappar i SharePoint' }
        [pscustomobject]@{ Fält = 'Sajt'; Värde = $(if ($sajtUrl) { $sajtUrl + $inom } elseif ($sajtNamn) { "sajter som matchar '$Sajt'" } else { 'Alla sajter' }) }
        [pscustomobject]@{ Fält = 'Period'; Värde = '{0:yyyy-MM-dd} – {1:yyyy-MM-dd} ({2} dagar)' -f $start.ToLocalTime(), $slut.ToLocalTime(), $antal }
        [pscustomobject]@{ Fält = 'Filter'; Värde = (@($(if ($Namn) { "namn innehåller '$Namn'" }), $(if ($Vem) { "borttaget av $Vem" })) | Where-Object { $_ }) -join ' · ' }
        [pscustomobject]@{ Fält = 'Borttagningar'; Värde = $borta.Count }
        [pscustomobject]@{ Fält = 'Personer'; Värde = @($perPerson).Count }
        [pscustomobject]@{ Fält = 'Borta för gott'; Värde = @($borta | Where-Object Händelse -eq 'Borta för gott').Count }
        [pscustomobject]@{ Fält = 'Skapad'; Värde = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
        [pscustomobject]@{ Fält = 'Källa'; Värde = 'Microsoft 365 granskningslogg (Search-UnifiedAuditLog)' }
    )
    $om | Export-Excel -Path $path -WorksheetName 'Sammanfattning' -Title 'Vem har tagit bort vad?' -TitleBold -TitleSize 16 `
        -TableName Om -TableStyle Light1
    $perPerson | Export-Excel -Path $path -WorksheetName 'Sammanfattning' -StartRow ($om.Count + 5) -TableName PerPerson -TableStyle Medium2
    $borta | Export-Excel -Path $path -WorksheetName 'Borttaget' -TableName Borttaget -TableStyle Medium2 -FreezeTopRow -BoldTopRow
    $perDag | Export-Excel -Path $path -WorksheetName 'Per dag' -TableName PerDag -TableStyle Medium2 -FreezeTopRow -BoldTopRow

    # Formatering: kolumnbredder, datumformat och rött för det som är borta för gott
    $pkg = Open-ExcelPackage -Path $path
    $ws = $pkg.Workbook.Worksheets['Sammanfattning']
    $ws.Cells[($om.Count + 4), 1].Value = 'Per person'
    $ws.Cells[($om.Count + 4), 1].Style.Font.Bold = $true; $ws.Cells[($om.Count + 4), 1].Style.Font.Size = 13
    foreach ($c in 1..8) { $ws.Column($c).Width = @(30, 48, 10, 10, 15, 18, 18, 50)[$c - 1] }
    foreach ($c in 6, 7) { $ws.Column($c).Style.Numberformat.Format = 'yyyy-mm-dd hh:mm' }
    $ws = $pkg.Workbook.Worksheets['Borttaget']
    foreach ($c in 1..10) { $ws.Column($c).Width = @(18, 34, 22, 7, 34, 60, 70, 45, 44, 15)[$c - 1] }
    $ws.Column(1).Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
    $n = $borta.Count + 1
    Add-ConditionalFormatting -Worksheet $ws -Range "A2:J$n" -RuleType Expression -ConditionValue '$C2="Borta för gott"' `
        -ForegroundColor ([System.Drawing.Color]::DarkRed) -BackgroundColor ([System.Drawing.Color]::MistyRose)
    $ws = $pkg.Workbook.Worksheets['Per dag']
    foreach ($c in 1..3) { $ws.Column($c).Width = @(14, 10, 90)[$c - 1] }
    $pkg.Workbook.Worksheets['Sammanfattning'].Select()
    Close-ExcelPackage $pkg
    } finally { $global:WarningPreference = $tidigare }
    Write-Host "Excel-rapport: $path"
    Invoke-Item -LiteralPath $path
}

$borta | Select-Object Tid, 'Borttaget av', Händelse, Typ, Namn, Plats, 'Kan återställas?'
