<#PUPPA
namn:        Gästkonton
kategori:    Användare
typ:         läs
kräver:      graph:User.Read.All
beskrivning: Alla gäster med inbjudningsstatus och när de skapades.
#>
Get-MgUser -Filter "userType eq 'Guest'" -All -Property displayName,mail,createdDateTime,externalUserState,accountEnabled |
    Sort-Object CreatedDateTime -Descending |
    Select-Object DisplayName, Mail, @{ n = 'Status'; e = { $_.ExternalUserState } }, AccountEnabled, CreatedDateTime
