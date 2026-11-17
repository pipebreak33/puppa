<#PUPPA
namn:        Hemligheter och certifikat som går ut
kategori:    Entra ID
typ:         läs
kräver:      graph:Application.Read.All
beskrivning: Klienthemligheter och certifikat på appregistreringar som går ut inom X dagar (eller redan har gått ut).
param:       Dagar | Inom hur många dagar? | 30
#>
$grans = (Get-Date).AddDays([int]$Dagar)
Get-MgApplication -All -Property 'appId,displayName,passwordCredentials,keyCredentials' | ForEach-Object {
    $a = $_
    foreach ($c in @($a.PasswordCredentials | ForEach-Object { $_ | Add-Member -NotePropertyName Typ -NotePropertyValue 'Hemlighet' -PassThru }) +
                   @($a.KeyCredentials | ForEach-Object { $_ | Add-Member -NotePropertyName Typ -NotePropertyValue 'Certifikat' -PassThru })) {
        if ($c -and $c.EndDateTime -le $grans) {
            [pscustomobject]@{
                App = $a.DisplayName; ClientId = $a.AppId; Typ = $c.Typ; Namn = $c.DisplayName
                'Går ut' = $c.EndDateTime; Dagar = [int]($c.EndDateTime - (Get-Date)).TotalDays
            }
        }
    }
} | Sort-Object 'Går ut'
