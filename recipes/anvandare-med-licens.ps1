<#PUPPA
namn:        Vilka har en viss licens?
kategori:    Användare
typ:         läs
kräver:      graph:User.Read.All,Organization.Read.All
beskrivning: Användare med en viss licens (SKU-namn, t.ex. O365_BUSINESS_PREMIUM eller SPE_E3).
param:       Sku | SKU-namn (se Licensöversikt) | | @sku
#>
$s = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $Sku
if (-not $s) { throw "Hittar ingen SKU som heter $Sku." }
Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($s.SkuId))" -All -Property displayName,userPrincipalName,accountEnabled |
    Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, AccountEnabled
