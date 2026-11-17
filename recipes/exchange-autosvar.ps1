<#PUPPA
namn:        Sätt autosvar
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Slår på (eller av) automatiskt svar för en postlåda, samma text internt och externt.
param:       Postlada | Postlådan (adress) | | @mailbox
param:       Lage | Läge | På | På,Av
param:       Text? | Svarstext (krävs för På) |
#>
if ($Lage -eq 'Av') { Set-MailboxAutoReplyConfiguration -Identity $Postlada -AutoReplyState Disabled -ErrorAction Stop }
else {
    if (-not $Text) { throw 'Ange en svarstext.' }
    Set-MailboxAutoReplyConfiguration -Identity $Postlada -AutoReplyState Enabled -InternalMessage $Text -ExternalMessage $Text -ExternalAudience All -ErrorAction Stop
}
Get-MailboxAutoReplyConfiguration -Identity $Postlada | Select-Object Identity, AutoReplyState
