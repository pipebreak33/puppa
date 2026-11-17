<#PUPPA
namn:        Företagsappar och samtycken
kategori:    Entra ID
typ:         läs
kräver:      graph:Application.Read.All,Directory.Read.All
beskrivning: Appar som fått delegerade rättigheter i tenanten – vilka scopes, för alla (admin) eller enskilda användare.
param:       Filter? | Del av appens namn (valfritt) |
#>
$sp = @{}
Get-MgOauth2PermissionGrant -All | Group-Object ClientId | ForEach-Object {
    $id = $_.Name
    if (-not $sp[$id]) { $sp[$id] = Get-MgServicePrincipal -ServicePrincipalId $id -Property displayName,appId,publisherName,appOwnerOrganizationId -ErrorAction SilentlyContinue }
    $s = $sp[$id]
    if ($Filter -and $s.DisplayName -notlike "*$Filter*") { return }
    [pscustomobject]@{
        App = $s.DisplayName; ClientId = $s.AppId; Utgivare = $s.PublisherName
        'Samtycke för alla' = [bool]($_.Group | Where-Object ConsentType -eq 'AllPrincipals')
        Användare = @($_.Group | Where-Object ConsentType -eq 'Principal').Count
        Scopes = (($_.Group.Scope -join ' ').Split(' ', [StringSplitOptions]::RemoveEmptyEntries) | Sort-Object -Unique) -join ' '
    }
} | Sort-Object App
