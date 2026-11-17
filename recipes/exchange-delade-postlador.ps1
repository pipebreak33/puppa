<#PUPPA
namn:        Delade postlådor och medlemmar
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Alla delade postlådor med vilka som har fullbehörighet.
param:       Filter? | Del av namn eller adress |
#>
Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Where-Object { -not $Filter -or $_.DisplayName -like "*$Filter*" -or $_.PrimarySmtpAddress -like "*$Filter*" } | Sort-Object DisplayName |
    ForEach-Object {
        $med = Get-EXOMailboxPermission -Identity $_.ExternalDirectoryObjectId | Where-Object { $_.User -notlike 'NT AUTHORITY*' -and -not $_.IsInherited } | ForEach-Object User
        [pscustomobject]@{ Namn = $_.DisplayName; Adress = $_.PrimarySmtpAddress; Medlemmar = ($med -join ', ') }
    }
