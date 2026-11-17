<#PUPPA
namn:        Kanaler i ett team
kategori:    Teams
typ:         läs
kräver:      teams
beskrivning: Standard-, privata och delade kanaler.
param:       Team | Teamets namn | | @team
#>
$t = @(Get-Team -DisplayName $Team)
if ($t.Count -ne 1) { throw "Hittade $($t.Count) team som heter '$Team'." }
$t = $t[0]

Get-TeamChannel -GroupId $t.GroupId | Select-Object DisplayName, MembershipType, Description
