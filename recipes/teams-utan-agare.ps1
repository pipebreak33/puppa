<#PUPPA
namn:        Team med bara en ägare
kategori:    Teams
typ:         läs
kräver:      teams
beskrivning: Team som har en eller inga ägare – sårbart när någon slutar.
#>
Get-Team | ForEach-Object {
    $o = @(Get-TeamUser -GroupId $_.GroupId -Role Owner)
    if ($o.Count -le 1) { [pscustomobject]@{ Team = $_.DisplayName; Ägare = ($o.User -join ', '); Antal = $o.Count } }
}
