<#PUPPA
namn:        Vidarebefordringar
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Postlådor som vidarebefordrar e-post (internt eller externt).
#>
Get-EXOMailbox -ResultSize Unlimited -Properties ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward |
    Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress } |
    Select-Object DisplayName, PrimarySmtpAddress, @{ n = 'Till'; e = { if ($_.ForwardingSmtpAddress) { $_.ForwardingSmtpAddress } else { $_.ForwardingAddress } } },
        @{ n = 'Behåll kopia'; e = { $_.DeliverToMailboxAndForward } }
