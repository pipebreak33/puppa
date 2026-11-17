<#PUPPA
namn:        Vilka delade postlådor kommer en användare åt?
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Går igenom alla delade postlådor (tar en stund i stora tenanter).
param:       Anvandare | Användare (UPN) | | @recipient
#>
Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
    $m = $_
    $p = Get-EXOMailboxPermission -Identity $m.ExternalDirectoryObjectId -User $Anvandare -ErrorAction SilentlyContinue
    if ($p) { [pscustomobject]@{ Postlåda = $m.DisplayName; Adress = $m.PrimarySmtpAddress; Åtkomst = ($p.AccessRights -join ', ') } }
}
