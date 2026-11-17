<#PUPPA
namn:        Ta bort åtkomst till postlåda
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Tar bort fullbehörighet, Skicka som och Skicka för för en användare.
param:       Postlada | Postlådan (adress) | | @mailbox
param:       Anvandare | Vems åtkomst ska bort? (UPN) | | @recipient
#>
Remove-MailboxPermission -Identity $Postlada -User $Anvandare -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction SilentlyContinue
Remove-RecipientPermission -Identity $Postlada -Trustee $Anvandare -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
Set-Mailbox -Identity $Postlada -GrantSendOnBehalfTo @{ Remove = $Anvandare } -ErrorAction SilentlyContinue
Write-Host "Klart. Kvarvarande åtkomst:"
Get-EXOMailboxPermission -Identity $Postlada | Where-Object { $_.User -notlike 'NT AUTHORITY*' -and -not $_.IsInherited } | Select-Object User, AccessRights
