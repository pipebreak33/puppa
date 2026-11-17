<#PUPPA
namn:        Sätt vidarebefordran
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Vidarebefordrar e-post från en postlåda till en annan adress (eller stänger av).
param:       Postlada | Postlådan (adress) | | @mailbox
param:       Till? | Vidarebefordra till (tomt = stäng av) |
param:       Kopia | Behåll kopia i postlådan? | Ja | Ja,Nej
#>
if ($Till) { Set-Mailbox -Identity $Postlada -ForwardingSmtpAddress "smtp:$Till" -DeliverToMailboxAndForward ($Kopia -eq 'Ja') -ErrorAction Stop }
else { Set-Mailbox -Identity $Postlada -ForwardingSmtpAddress $null -ForwardingAddress $null -ErrorAction Stop }
Get-EXOMailbox -Identity $Postlada -Properties ForwardingSmtpAddress, DeliverToMailboxAndForward | Select-Object DisplayName, ForwardingSmtpAddress, DeliverToMailboxAndForward
