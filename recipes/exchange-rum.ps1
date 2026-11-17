<#PUPPA
namn:        Rum och resurser
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Alla rums- och utrustningspostlådor.
#>
Get-EXOMailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited | Sort-Object DisplayName |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
