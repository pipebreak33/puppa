<#PUPPA
namn:        Spärra inloggning
kategori:    Användare
typ:         ändrar
kräver:      graph:User.ReadWrite.All
beskrivning: Stänger av kontot och loggar ut alla sessioner (första steget vid offboarding eller intrång).
param:       Anvandare | Användare (UPN) |
#>
Update-MgUser -UserId $Anvandare -AccountEnabled:$false -ErrorAction Stop
Revoke-MgUserSignInSession -UserId $Anvandare -ErrorAction Stop | Out-Null
Get-MgUser -UserId $Anvandare -Property displayName,userPrincipalName,accountEnabled | Select-Object DisplayName, UserPrincipalName, AccountEnabled
