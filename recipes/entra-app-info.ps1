<#PUPPA
namn:        Information om en appregistrering
kategori:    Entra ID
typ:         läs
kräver:      graph:Application.Read.All
beskrivning: ClientId, ägare, redirect-URI:er, API-rättigheter (med namn), hemligheter/certifikat och om företagsappen finns.
param:       App | Appens namn eller ClientId | | @app
#>
$a = Get-MgApplication -Filter "appId eq '$App'" -ErrorAction Stop | Select-Object -First 1
if (-not $a) { throw "Hittade ingen app med ClientId $App" }
$sp = Get-MgServicePrincipal -Filter "appId eq '$App'" -ErrorAction SilentlyContinue | Select-Object -First 1
$agare = (Get-MgApplicationOwner -ApplicationId $a.Id -All | ForEach-Object { $_.AdditionalProperties.userPrincipalName ?? $_.AdditionalProperties.displayName }) -join ', '
$rattigheter = foreach ($r in $a.RequiredResourceAccess) {
    $res = Get-MgServicePrincipal -Filter "appId eq '$($r.ResourceAppId)'" -Property displayName,appRoles,oauth2PermissionScopes -ErrorAction SilentlyContinue | Select-Object -First 1
    foreach ($x in $r.ResourceAccess) {
        $namn = if ($x.Type -eq 'Role') { ($res.AppRoles | Where-Object Id -eq $x.Id).Value } else { ($res.Oauth2PermissionScopes | Where-Object Id -eq $x.Id).Value }
        "$($res.DisplayName): $(if ($namn) { $namn } else { $x.Id }) ($(if ($x.Type -eq 'Role') { 'app' } else { 'delegerad' }))"
    }
}
$uris = @($a.PublicClient.RedirectUris) + @($a.Web.RedirectUris) + @($a.Spa.RedirectUris) | Where-Object { $_ }
[pscustomobject]@{
    Namn = $a.DisplayName; ClientId = $a.AppId; ObjektId = $a.Id; Skapad = $a.CreatedDateTime
    Målgrupp = $a.SignInAudience; Ägare = $agare
    'Företagsapp' = if ($sp) { "finns ($(if ($sp.AccountEnabled) { 'aktiv' } else { 'avstängd' }))" } else { 'saknas – ingen har samtyckt' }
    'Redirect-URI:er' = $uris -join ', '
    Rättigheter = $rattigheter -join "`n"
    Hemligheter = ($a.PasswordCredentials | ForEach-Object { "$($_.DisplayName) går ut $($_.EndDateTime.ToString('yyyy-MM-dd'))" }) -join ', '
    Certifikat = ($a.KeyCredentials | ForEach-Object { "$($_.DisplayName) går ut $($_.EndDateTime.ToString('yyyy-MM-dd'))" }) -join ', '
}
