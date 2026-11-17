<#PUPPA
namn:        En användares enheter
kategori:    Entra ID
typ:         läs
kräver:      graph:Device.Read.All,User.Read.All
beskrivning: Enheter som användaren äger eller har registrerat, med OS, anslutning, kompatibilitet och senaste inloggning.
param:       Anvandare | Användare (UPN) | | @user
#>
$ids = @(Get-MgUserOwnedDevice -UserId $Anvandare -All) + @(Get-MgUserRegisteredDevice -UserId $Anvandare -All) | Select-Object -ExpandProperty Id -Unique
foreach ($id in $ids) {
    $d = Get-MgDevice -DeviceId $id -ErrorAction SilentlyContinue
    if (-not $d) { continue }
    [pscustomobject]@{
        Namn = $d.DisplayName; OS = "$($d.OperatingSystem) $($d.OperatingSystemVersion)".Trim()
        Anslutning = switch ($d.TrustType) { 'AzureAd' { 'Entra-ansluten' } 'ServerAd' { 'Hybrid' } 'Workplace' { 'Registrerad' } default { $d.TrustType } }
        Kompatibel = $d.IsCompliant; Hanterad = $d.IsManaged; 'Senast inloggad' = $d.ApproximateLastSignInDateTime
    }
}
