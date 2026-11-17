<#PUPPA
namn:        Användare på en sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp:Url
beskrivning: Alla användare som finns i sajtens användarlista (även gäster).
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
param:       Filter? | Del av namn eller e-post |
#>
Get-PnPUser | Where-Object { $_.PrincipalType -eq 'User' } |
    Where-Object { -not $Filter -or $_.Title -like "*$Filter*" -or $_.Email -like "*$Filter*" } | Sort-Object Title |
    Select-Object Title, Email, @{ n = 'Gäst'; e = { $_.LoginName -like '*#ext#*' } }, IsSiteAdmin
