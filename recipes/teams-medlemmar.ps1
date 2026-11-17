<#PUPPA
namn:        Medlemmar i ett team
kategori:    Teams
typ:         läs
kräver:      teams
beskrivning: Ägare, medlemmar och gäster i ett team.
param:       Team | Teamets namn |
#>
$t = @(Get-Team -DisplayName $Team)
if ($t.Count -ne 1) { throw "Hittade $($t.Count) team som heter '$Team'." }
$t = $t[0]

Get-TeamUser -GroupId $t.GroupId | Sort-Object Role, Name | Select-Object Name, User, Role
