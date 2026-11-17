<#PUPPA
namn:        Lista team
kategori:    Teams
typ:         läs
kräver:      teams
beskrivning: Alla team med synlighet och om de är arkiverade.
param:       Filter? | Del av teamnamnet |
#>
Get-Team | Where-Object { -not $Filter -or $_.DisplayName -like "*$Filter*" } | Sort-Object DisplayName |
    Select-Object DisplayName, Visibility, Archived, MailNickName
