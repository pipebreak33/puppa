<#PUPPA
namn:        Hitta PnP-appen
kategori:    Entra ID
typ:         läs
kräver:      graph:Application.Read.All
beskrivning: Appregistreringar som ser ut som PnP PowerShell-appar: PnP i namnet, eller publik klient (localhost) med SharePoint-rättigheter.
#>
$sp = '00000003-0000-0ff1-ce00-000000000000'
$pnp = Get-MgApplication -All -Property 'id,appId,displayName,createdDateTime,requiredResourceAccess,publicClient,keyCredentials' |
    Where-Object { $_.DisplayName -match 'PnP' -or (($_.RequiredResourceAccess.ResourceAppId -contains $sp) -and
                   (@($_.PublicClient.RedirectUris) -match 'localhost')) } | Sort-Object CreatedDateTime -Descending
if (-not $pnp) { Write-Host 'Hittade ingen PnP-app. Välj "Ny PnP-appregistrering" i Puppa för att skapa en.'; return }
$pnp | ForEach-Object {
    [pscustomobject]@{
        Namn = $_.DisplayName; ClientId = $_.AppId; Skapad = $_.CreatedDateTime
        Inloggning = if ($_.KeyCredentials.Count) { 'certifikat (app-only)' } elseif (@($_.PublicClient.RedirectUris) -match 'localhost') { 'interaktiv' } else { '?' }
        SharePoint = [bool]($_.RequiredResourceAccess.ResourceAppId -contains $sp)
    }
}
