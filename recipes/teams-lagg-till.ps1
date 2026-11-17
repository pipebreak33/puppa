<#PUPPA
namn:        Lägg till i team
kategori:    Teams
typ:         ändrar
kräver:      teams
beskrivning: Lägger till en eller flera användare i ett team, som medlem eller ägare.
param:       Team | Teamets namn |
param:       Anvandare | Användare (UPN, flera med komma) |
param:       Roll | Roll | Member | Member,Owner
#>
$t = @(Get-Team -DisplayName $Team)
if ($t.Count -ne 1) { throw "Hittade $($t.Count) team som heter '$Team'." }
$t = $t[0]

foreach ($a in @($Anvandare -split '[,;\s]+' | Where-Object { $_ })) {
    Add-TeamUser -GroupId $t.GroupId -User $a -Role $Roll -ErrorAction Stop
    Write-Host "  + $a ($Roll)"
}
