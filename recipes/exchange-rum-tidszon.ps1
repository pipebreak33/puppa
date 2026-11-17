<#PUPPA
namn:        Rum: tidszon och språk
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Sätter tidszon och språk på en rumsresurs så att bokningar visas rätt.
param:       Rum | Rummets adress | | @mailbox
param:       Tidszon | Tidszon | W. Europe Standard Time
param:       Sprak | Språk | sv-SE | sv-SE,en-US,nb-NO,da-DK,fi-FI
#>
Set-MailboxRegionalConfiguration -Identity $Rum -TimeZone $Tidszon -Language $Sprak -LocalizeDefaultFolderName -ErrorAction Stop
Set-MailboxCalendarConfiguration -Identity $Rum -WorkingHoursTimeZone $Tidszon -ErrorAction Stop
Get-MailboxRegionalConfiguration -Identity $Rum | Select-Object Identity, TimeZone, Language
