<#PUPPA
namn:        Medlemmar i distributionslista
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Vilka som är med i en distributionslista.
param:       Lista | Listans namn eller adress |
#>
Get-DistributionGroupMember -Identity $Lista -ResultSize Unlimited | Sort-Object DisplayName |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientType
