<#PUPPA
namn:        Mina M365-moduler
kategori:    Lokalt
typ:         läs
kräver:      ingen
beskrivning: Vilka versioner av PnP, Exchange, Teams, Graph m.fl. som finns på den här datorn.
#>
$namn = 'PnP.PowerShell', 'ExchangeOnlineManagement', 'MicrosoftTeams', 'Microsoft.Graph.Authentication',
        'Microsoft.Graph', 'Az.Accounts', 'ImportExcel', 'Microsoft.PowerShell.SecretManagement', 'Microsoft.PowerShell.SecretStore'
Get-Module -ListAvailable -Name $namn | Group-Object Name | ForEach-Object {
    $v = $_.Group | Sort-Object Version -Descending
    [pscustomobject]@{ Modul = $_.Name; Version = "$($v[0].Version)"; 'Äldre' = (@($v | Select-Object -Skip 1 | ForEach-Object { "$($_.Version)" }) -join ', ') }
} | Sort-Object Modul
