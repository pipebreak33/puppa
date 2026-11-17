<#PUPPA
namn:        Tenantinställningar (SharePoint)
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: De viktigaste SharePoint-inställningarna för tenanten: delning, länkar, OneDrive.
#>
$t = Get-PnPTenant
[pscustomobject]@{
    'Extern delning'          = "$($t.SharingCapability)"
    'Standardlänk'            = "$($t.DefaultSharingLinkType)"
    'Standardbehörighet länk' = "$($t.DefaultLinkPermission)"
    'Länk för alla går ut (d)'= $t.RequireAnonymousLinksExpireInDays
    'Gäster måste matcha'     = $t.RequireAcceptingAccountMatchInvitedAccount
    'OneDrive kvot (GB)'      = [math]::Round($t.OneDriveStorageQuota / 1024)
    'OneDrive kvar efter avslut (d)' = $t.OrphanedPersonalSitesRetentionPeriod
    'Legacy-autentisering'    = $t.LegacyAuthProtocolsEnabled
}
