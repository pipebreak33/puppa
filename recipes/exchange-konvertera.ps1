<#PUPPA
namn:        Konvertera till delad postlåda
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Gör om en användarpostlåda till delad (t.ex. vid offboarding, frigör sedan licensen).
param:       Postlada | Postlådan (adress) |
#>
Set-Mailbox -Identity $Postlada -Type Shared -ErrorAction Stop
Get-EXOMailbox -Identity $Postlada | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
Write-Host 'Kom ihåg: ta bort licensen när konverteringen syns (postlådor under 50 GB behöver ingen).'
