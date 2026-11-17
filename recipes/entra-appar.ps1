<#PUPPA
namn:        Lista appregistreringar
kategori:    Entra ID
typ:         läs
kräver:      graph:Application.Read.All
beskrivning: Alla appregistreringar i Entra ID med ClientId, när de skapades och när närmaste hemlighet/certifikat går ut.
param:       Filter? | Del av appens namn |
#>
$p = @{ All = $true; Property = 'id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials' }
if ($Filter) { $p.Search = ('"displayName:{0}"' -f ($Filter -replace '"', '')); $p.ConsistencyLevel = 'eventual' }
Get-MgApplication @p | Sort-Object DisplayName | ForEach-Object {
    $slut = @($_.PasswordCredentials.EndDateTime) + @($_.KeyCredentials.EndDateTime) | Where-Object { $_ } | Sort-Object | Select-Object -First 1
    [pscustomobject]@{
        Namn = $_.DisplayName; ClientId = $_.AppId; Skapad = $_.CreatedDateTime
        Hemligheter = $_.PasswordCredentials.Count; Certifikat = $_.KeyCredentials.Count
        'Går ut först' = $slut; Målgrupp = $_.SignInAudience
    }
}
