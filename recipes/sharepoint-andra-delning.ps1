<#PUPPA
namn:        Ändra extern delning på en sajt
kategori:    SharePoint
typ:         ändrar
kräver:      pnp-admin
beskrivning: Sätter delningsnivån för en sajt (kan inte vara öppnare än tenantens).
param:       Url | Sajtens URL | {rootUrl}/sites/
param:       Niva | Ny nivå | ExistingExternalUserSharingOnly | Disabled,ExistingExternalUserSharingOnly,ExternalUserSharingOnly,ExternalUserAndGuestSharing
#>
Set-PnPTenantSite -Identity $Url -SharingCapability $Niva -ErrorAction Stop
Get-PnPTenantSite -Identity $Url | Select-Object Url, @{ n = 'Delning'; e = { "$($_.SharingCapability)" } }
