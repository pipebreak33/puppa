<#PUPPA
namn:        Vad har flyttats på en sajt?
kategori:    SharePoint
typ:         läs
kräver:      exo
beskrivning: Flyttade filer och mappar ur granskningsloggen – vem (UPN), när, från var och vart (med länkar) + förslag på var du ska leta.
param:       Sajt? | Sajt – skriv en del av namnet, eller klistra in en länk (även till bibliotek/mapp) (tomt = visa var det flyttats mest) | | @site
param:       Dagar | Hur många dagar bakåt? (högst 180) | 30
param:       Namn? | Del av fil- eller mappnamnet (valfritt) |
param:       Vem? | Bara flyttar gjorda av (UPN, valfritt) | | @user
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

$slut = (Get-Date).ToUniversalTime(); $start = $slut.AddDays(-$antal)
$session = [guid]::NewGuid().ToString()
$sok = @{ StartDate = $start; EndDate = $slut; RecordType = 'SharePointFileOperation'; Operations = 'FileMoved', 'FolderMoved'
         SessionId = $session; SessionCommand = 'ReturnLargeSet'; ResultSize = 5000 }
if ($Vem) { $sok.UserIds = $Vem }
$rader = @(); $varv = 0
do {
    $sida = @(Search-UnifiedAuditLog @sok)
    $rader += $sida; $varv++
} while ($sida.Count -eq 5000 -and $varv -lt 10)

$flytt = foreach ($r in $rader) {
    $d = $r.AuditData | ConvertFrom-Json
    if ($sajtUrl -and -not "$($d.SiteUrl)".StartsWith($sajtUrl, [StringComparison]::OrdinalIgnoreCase)) { continue }
    if ($sajtNamn -and -not (("$($d.SiteUrl)" -replace '^https://[^/]+', '') -like $sajtNamn)) { continue }
    $fran = ("$($d.SourceRelativeUrl)/$($d.SourceFileName)").TrimStart('/')
    $till = ("$($d.DestinationRelativeUrl)/$($d.DestinationFileName)").TrimStart('/')
    if ($inom -and -not ($fran -like "$inom*" -or $till -like "$inom*")) { continue }
    if ($Namn -and -not ($d.SourceFileName -like "*$Namn*" -or $d.DestinationFileName -like "*$Namn*")) { continue }
    $bas = "$($d.SiteUrl)".TrimEnd('/')
    [pscustomobject]@{
        Tid       = ([datetime]$d.CreationTime).ToLocalTime()
        'Flyttad av' = $d.UserId
        Typ       = if ($d.Operation -eq 'FolderMoved') { 'Mapp' } else { 'Fil' }
        Namn      = if ($d.DestinationFileName) { $d.DestinationFileName } else { $d.SourceFileName }
        Från      = "$bas/$fran" -replace ' ', '%20'
        Till      = $(if ($d.ObjectId -match '^https://') { $d.ObjectId } else { "$bas/$till" }) -replace ' ', '%20'
        Sajt      = $bas
        TillMapp  = "$bas/$($d.DestinationRelativeUrl)".TrimEnd('/') -replace ' ', '%20'
    }
}
$flytt = @($flytt | Sort-Object Tid -Descending)

if (-not $flytt) {
    Write-Host (T "No moves found$(if ($sajtUrl) { " on $sajtUrl" })$(if ($Vem) { " made by $Vem" }) in the last $antal days." "Inga flyttar hittades$(if ($sajtUrl) { " på $sajtUrl" })$(if ($Vem) { " gjorda av $Vem" }) de senaste $antal dagarna.")
    Write-Host (T 'Tip: increase the number of days, clear the name filter, or check the recycle bin (the "Recycle bin" recipe).' 'Förslag: öka antalet dagar, töm namnfiltret, eller kolla papperskorgen (receptet "Papperskorgen").')
    return
}

if (-not $sajtUrl -and -not $sajtNamn) {
    # Ingen sajt angiven: föreslå var det har flyttats mest
    Write-Host (T "This is where the most was moved in the last $antal days. Run again with one of the links:" "Här har det flyttats mest de senaste $antal dagarna – kör igen med en av länkarna:")
    $flytt | Group-Object Sajt | Sort-Object Count -Descending | Select-Object -First 25 | ForEach-Object {
        [pscustomobject]@{ Sajt = $_.Name; Flyttar = $_.Count; Senast = ($_.Group | Select-Object -First 1).Tid
                           'Flyttad av' = (($_.Group.'Flyttad av' | Group-Object | Sort-Object Count -Descending | Select-Object -First 3).Name) -join ', ' }
    }
    return
}

$flytt | Select-Object Tid, 'Flyttad av', Typ, Namn, Från, Till

# Förslag på ställen att leta: mapparna som saker har flyttats till, flest först
Write-Host ''
Write-Host (T 'Places to look (where things were moved, most first):' 'Förslag på ställen att leta (dit saker har flyttats, flest först):')
$flytt | Group-Object TillMapp | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
    Write-Host ("  {0,4}  {1}" -f $_.Count, $_.Name)
}
$andra = @($flytt | Where-Object { -not $_.Till.StartsWith($sajtUrl, [StringComparison]::OrdinalIgnoreCase) })
if ($andra) { Write-Host (T "  $($andra.Count) were moved to another site, see the To column." "  $($andra.Count) flyttades till en annan sajt – se kolumnen Till.") }
Write-Host (T '  Still missing something? Check the recycle bin (the "Recycle bin" recipe) and the version history in the folder.' '  Saknas något ändå? Kolla papperskorgen (receptet "Papperskorgen") och versionshistoriken i mappen.')
