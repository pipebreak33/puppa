<#PUPPA
namn:        Information om en sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp-admin
beskrivning: Ägare, mall, lagring, delning, låsstatus och kopplad grupp för en sajt.
param:       Url | Sajtens URL | {rootUrl}/sites/
#>
$s = Get-PnPTenantSite -Identity $Url -Detailed -ErrorAction Stop
[pscustomobject]@{
    Titel = $s.Title; Url = $s.Url; Mall = $s.Template; Ägare = $s.Owner
    'Lagring GB' = [math]::Round($s.StorageUsageCurrent / 1024, 2); 'Kvot GB' = [math]::Round($s.StorageQuota / 1024)
    Delning = "$($s.SharingCapability)"; Låst = "$($s.LockState)"; 'M365-grupp' = $s.GroupId
    'Senast ändrad' = $s.LastContentModifiedDate
}
