<#PUPPA
namn:        Enheter i Entra ID
kategori:    Entra ID
typ:         läs
kräver:      graph:Device.Read.All
beskrivning: Datorer och mobiler: OS, hur de är anslutna (Entra/hybrid/registrerad), kompatibla, hanterade och senaste inloggning.
param:       Filter? | Del av enhetens namn (valfritt) |
param:       Inaktiva? | Bara enheter som inte loggat in på X dagar (valfritt) |
#>
$p = @{ All = $true; Property = 'id,displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,accountEnabled,approximateLastSignInDateTime,registrationDateTime' }
if ($Filter) { $p.Search = ('"displayName:{0}"' -f ($Filter -replace '"', '')); $p.ConsistencyLevel = 'eventual' }
$grans = if ($Inaktiva) { (Get-Date).AddDays(-[int]$Inaktiva) }
Get-MgDevice @p | Where-Object { -not $grans -or -not $_.ApproximateLastSignInDateTime -or $_.ApproximateLastSignInDateTime -lt $grans } |
    Sort-Object DisplayName | ForEach-Object {
        [pscustomobject]@{
            Namn = $_.DisplayName; OS = "$($_.OperatingSystem) $($_.OperatingSystemVersion)".Trim()
            Anslutning = switch ($_.TrustType) { 'AzureAd' { 'Entra-ansluten' } 'ServerAd' { 'Hybrid' } 'Workplace' { 'Registrerad' } default { $_ } }
            Kompatibel = $_.IsCompliant; Hanterad = $_.IsManaged; Aktiv = $_.AccountEnabled
            'Senast inloggad' = $_.ApproximateLastSignInDateTime; Registrerad = $_.RegistrationDateTime
        }
    }
