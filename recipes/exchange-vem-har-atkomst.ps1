<#PUPPA
namn:        Vem har åtkomst till en postlåda?
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Fullbehörighet, Skicka som och Skicka för på en postlåda.
param:       Postlada | Postlådan (adress) | | @mailbox
#>
Get-EXOMailboxPermission -Identity $Postlada | Where-Object { $_.User -notlike 'NT AUTHORITY*' -and -not $_.IsInherited } |
    ForEach-Object { [pscustomobject]@{ Vem = $_.User; Åtkomst = ($_.AccessRights -join ', ') } }
Get-EXORecipientPermission -Identity $Postlada | Where-Object { $_.Trustee -notlike 'NT AUTHORITY*' } |
    ForEach-Object { [pscustomobject]@{ Vem = $_.Trustee; Åtkomst = 'SendAs' } }
(Get-EXOMailbox -Identity $Postlada -Properties GrantSendOnBehalfTo).GrantSendOnBehalfTo |
    ForEach-Object { [pscustomobject]@{ Vem = $_; Åtkomst = 'SendOnBehalf' } }
