<#PUPPA
namn:        Inaktiva användare
kategori:    Användare
typ:         läs
kräver:      graph:User.Read.All,AuditLog.Read.All
beskrivning: Aktiva konton som inte loggat in på X dagar (kräver Entra ID P1).
param:       Dagar | Ingen inloggning på minst | 90
#>
$grans = (Get-Date).AddDays(-[int]$Dagar)
Get-MgUser -All -Filter 'accountEnabled eq true' -Property displayName,userPrincipalName,userType,signInActivity,assignedLicenses |
    Where-Object { -not $_.SignInActivity.LastSignInDateTime -or $_.SignInActivity.LastSignInDateTime -lt $grans } |
    Sort-Object { $_.SignInActivity.LastSignInDateTime } |
    Select-Object DisplayName, UserPrincipalName, UserType, @{ n = 'Senast inloggad'; e = { $_.SignInActivity.LastSignInDateTime } },
        @{ n = 'Licenser'; e = { @($_.AssignedLicenses).Count } }
