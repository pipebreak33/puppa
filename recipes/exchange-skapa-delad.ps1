<#PUPPA
namn:        Skapa delad postlåda
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Ny delad postlåda, valfritt med medlemmar som får fullbehörighet och Skicka som.
param:       Adress | E-postadress (t.ex. kundservice@contoso.se) |
param:       Namn | Visningsnamn |
param:       Medlemmar? | Medlemmar (UPN, flera med komma) | | @recipient+
param:       SkickatSparas | Spara skickat som även i den delade postlådan? | Ja | Ja,Nej
#>
$alias = ($Adress -split '@')[0]
$mb = New-Mailbox -Shared -Name $Namn -DisplayName $Namn -Alias $alias -PrimarySmtpAddress $Adress -ErrorAction Stop
if ($SkickatSparas -eq 'Ja') { Set-Mailbox -Identity $mb.Identity -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true }
foreach ($m in @($Medlemmar -split '[,;\s]+' | Where-Object { $_ })) {
    Add-MailboxPermission -Identity $Adress -User $m -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false | Out-Null
    Add-RecipientPermission -Identity $Adress -Trustee $m -AccessRights SendAs -Confirm:$false | Out-Null
    Write-Host (T "  + $m (full access + send as)" "  + $m (fullbehörighet + skicka som)")
}
Get-EXOMailbox -Identity $Adress | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
