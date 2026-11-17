<#PUPPA
namn:        Information om en användare
kategori:    Användare
typ:         läs
kräver:      graph:User.Read.All
beskrivning: Konto, licenser, chef, senaste lösenordsbyte och om kontot är aktivt.
param:       Anvandare | Användare (UPN) | | @user
#>
$u = Get-MgUser -UserId $Anvandare -Property id,displayName,userPrincipalName,accountEnabled,jobTitle,department,createdDateTime,lastPasswordChangeDateTime,usageLocation,assignedLicenses,userType -ErrorAction Stop
$sku = @{}; Get-MgSubscribedSku -All | ForEach-Object { $sku["$($_.SkuId)"] = $_.SkuPartNumber }
$chef = try { (Get-MgUserManager -UserId $u.Id -ErrorAction Stop).AdditionalProperties.displayName } catch { '' }
[pscustomobject]@{
    Namn = $u.DisplayName; UPN = $u.UserPrincipalName; Aktiv = $u.AccountEnabled; Typ = $u.UserType
    Titel = $u.JobTitle; Avdelning = $u.Department; Chef = $chef; Land = $u.UsageLocation
    Skapad = $u.CreatedDateTime; 'Lösenord bytt' = $u.LastPasswordChangeDateTime
    Licenser = ($u.AssignedLicenses | ForEach-Object { $sku["$($_.SkuId)"] }) -join ', '
}
