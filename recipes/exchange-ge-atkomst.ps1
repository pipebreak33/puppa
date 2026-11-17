<#PUPPA
namn:        Ge åtkomst till (delad) postlåda
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Fullbehörighet och/eller Skicka som för en eller flera användare.
param:       Postlada | Postlådan (adress) |
param:       Anvandare | Vem ska få åtkomst? (UPN, flera med komma) |
param:       Atkomst | Vad ska de få? | Fullbehörighet + Skicka som | Fullbehörighet + Skicka som,Bara fullbehörighet,Bara Skicka som,Skicka för (send on behalf)
param:       Automapping | Visa automatiskt i Outlook? | Ja | Ja,Nej
#>
foreach ($a in @($Anvandare -split '[,;\s]+' | Where-Object { $_ })) {
    if ($Atkomst -like '*fullbehörighet*' -or $Atkomst -like 'Fullbehörighet*') {
        Add-MailboxPermission -Identity $Postlada -User $a -AccessRights FullAccess -InheritanceType All -AutoMapping ($Automapping -eq 'Ja') -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Host "  + $a fullbehörighet"
    }
    if ($Atkomst -like '*Skicka som*') {
        Add-RecipientPermission -Identity $Postlada -Trustee $a -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Host "  + $a skicka som"
    }
    if ($Atkomst -like 'Skicka för*') {
        Set-Mailbox -Identity $Postlada -GrantSendOnBehalfTo @{ Add = $a } -ErrorAction Stop
        Write-Host "  + $a skicka för"
    }
}
