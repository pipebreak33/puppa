<#PUPPA
namn:        Distributionslistor
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Alla distributionslistor och e-postaktiverade säkerhetsgrupper.
param:       Filter? | Del av namn eller adress |
#>
Get-DistributionGroup -ResultSize Unlimited | Where-Object { -not $Filter -or $_.DisplayName -like "*$Filter*" -or $_.PrimarySmtpAddress -like "*$Filter*" } |
    Sort-Object DisplayName | Select-Object DisplayName, PrimarySmtpAddress, GroupType, @{ n = 'Externa får skicka'; e = { -not $_.RequireSenderAuthenticationEnabled } }
