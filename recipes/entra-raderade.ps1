<#PUPPA
namn:        Raderade användare och grupper
kategori:    Entra ID
typ:         läs
kräver:      graph:User.Read.All,Group.Read.All,Directory.Read.All
beskrivning: Borttagna användare och Microsoft 365-grupper som går att återställa (30 dagar), med när de togs bort.
#>
$nu = Get-Date
$u = Get-MgDirectoryDeletedItemAsUser -All -Property id,displayName,userPrincipalName,deletedDateTime | ForEach-Object {
    [pscustomobject]@{ Typ = 'Användare'; Namn = $_.DisplayName; Adress = $_.UserPrincipalName; Raderad = $_.DeletedDateTime
                       'Dagar kvar' = 30 - [int]($nu - $_.DeletedDateTime).TotalDays; Id = $_.Id } }
$g = Get-MgDirectoryDeletedItemAsGroup -All -Property id,displayName,mail,deletedDateTime | ForEach-Object {
    [pscustomobject]@{ Typ = 'Grupp'; Namn = $_.DisplayName; Adress = $_.Mail; Raderad = $_.DeletedDateTime
                       'Dagar kvar' = 30 - [int]($nu - $_.DeletedDateTime).TotalDays; Id = $_.Id } }
@($u) + @($g) | Where-Object { $_ } | Sort-Object Raderad -Descending
