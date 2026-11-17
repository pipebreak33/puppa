<#PUPPA
namn:        Vem ändrade vad i Entra ID?
kategori:    Entra ID
typ:         läs
kräver:      graph:AuditLog.Read.All,Directory.Read.All
beskrivning: Granskningsloggen: skapade/ändrade/borttagna användare, grupper, appar, roller, lösenordsåterställningar m.m.
param:       Dagar | Hur många dagar bakåt? (högst 30) | 7
param:       Vem? | Bara ändringar gjorda av (UPN, valfritt) | | @user
param:       Mal? | Bara ändringar på ett objekt vars namn innehåller (valfritt) |
param:       Aktivitet? | Bara aktiviteter som innehåller, t.ex. "Reset password", "Add member" (valfritt) |
#>
$fran = (Get-Date).ToUniversalTime().AddDays(-[Math]::Min([int]$Dagar, 30)).ToString('yyyy-MM-ddTHH:mm:ssZ')
$f = "activityDateTime ge $fran"
if ($Vem) { $f += " and initiatedBy/user/userPrincipalName eq '$($Vem -replace "'", "''")'" }
Get-MgAuditLogDirectoryAudit -Filter $f -All | ForEach-Object {
    $av = if ($_.InitiatedBy.User.UserPrincipalName) { $_.InitiatedBy.User.UserPrincipalName } else { "$($_.InitiatedBy.App.DisplayName) (app)" }
    $objekt = ($_.TargetResources | ForEach-Object { if ($_.UserPrincipalName) { $_.UserPrincipalName } else { $_.DisplayName } } | Where-Object { $_ }) -join ', '
    if ($Mal -and $objekt -notlike "*$Mal*") { return }
    if ($Aktivitet -and $_.ActivityDisplayName -notlike "*$Aktivitet*") { return }
    [pscustomobject]@{
        Tid = $_.ActivityDateTime.ToLocalTime(); Vem = $av; Aktivitet = $_.ActivityDisplayName; Mål = $objekt
        Resultat = $_.Result; Kategori = $_.Category
    }
} | Sort-Object Tid -Descending
