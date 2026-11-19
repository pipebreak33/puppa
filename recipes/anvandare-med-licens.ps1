<#PUPPA
namn:        Vilka har en viss licens?
kategori:    Användare
typ:         läs
kräver:      graph:User.Read.All,Organization.Read.All
beskrivning: Användare med en, flera eller alla av tenantens licenser. Välj i listan; ctrl+a kryssar alla.
param:       Sku | Licenser | | @sku+
#>
$alla = @(Get-MgSubscribedSku -All)
$val = @($Sku -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$skus = if (-not $val -or $val -contains '*') { $alla | Where-Object { $_.ConsumedUnits -gt 0 } }
        else { $alla | Where-Object { $_.SkuPartNumber -in $val } }
$saknas = @($val | Where-Object { $_ -ne '*' -and $_ -notin $alla.SkuPartNumber })
if ($saknas) { throw (T "No SKU called $($saknas -join ', ')." "Hittar ingen SKU som heter $($saknas -join ', ').") }
foreach ($s in $skus | Sort-Object { SkuName $_.SkuPartNumber }) {
    $namn = SkuName $s.SkuPartNumber
    Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($s.SkuId))" -All -ConsistencyLevel eventual -CountVariable n `
        -Property displayName,userPrincipalName,accountEnabled |
        Sort-Object DisplayName | ForEach-Object {
            [pscustomobject]@{ Licens = $namn; Namn = $_.DisplayName; UPN = $_.UserPrincipalName
                               Aktiv = $(if ($_.AccountEnabled) { 'ja' } else { 'nej' }) } }
}
