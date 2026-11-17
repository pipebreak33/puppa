<#PUPPA
namn:        Medlemmar i en grupp
kategori:    Grupper
typ:         läs
kräver:      graph:Group.Read.All,User.Read.All
beskrivning: Medlemmar och ägare i en grupp (namn eller e-post).
param:       Grupp | Gruppens namn eller e-post | | @group
#>
$q = $Grupp -replace "'", "''"
$g = @(Get-MgGroup -Filter "displayName eq '$q' or mail eq '$q'" -All)
if (-not $g) { $g = @(Get-MgGroup -Search "displayName:$Grupp" -ConsistencyLevel eventual -All) }
if ($g.Count -ne 1) { throw "Hittade $($g.Count) grupper för '$Grupp': $(($g.DisplayName | Select-Object -First 5) -join ', ')" }
$g = $g[0]

Write-Host "$($g.DisplayName)"
Get-MgGroupOwner -GroupId $g.Id -All | ForEach-Object { [pscustomobject]@{ Roll = 'Ägare'; Namn = $_.AdditionalProperties.displayName; Konto = $_.AdditionalProperties.userPrincipalName } }
Get-MgGroupMember -GroupId $g.Id -All | ForEach-Object {
    [pscustomobject]@{ Roll = 'Medlem'; Namn = $_.AdditionalProperties.displayName; Konto = $_.AdditionalProperties.userPrincipalName ?? $_.AdditionalProperties.mail }
} | Sort-Object Namn
