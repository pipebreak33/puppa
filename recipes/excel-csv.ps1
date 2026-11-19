<#PUPPA
namn:        Exportera senaste resultat till CSV
kategori:    Excel
typ:         lokal
kräver:      ingen
beskrivning: Sparar förra resultatet som CSV med semikolon och UTF-8 (öppnas rätt i svensk Excel).
param:       Filnamn | Filnamn (utan .csv) | {kund}-{datum}
#>
$senaste = Join-Path $env:PUPPA_DIR 'senaste.xml'
# PnP-recept körs i en egen PowerShell-process – filen har alltid det senaste resultatet
if (Test-Path $senaste) { $global:PuppaLast = Import-Clixml $senaste }
$rows = @($global:PuppaLast | Where-Object { $_ -isnot [string] })
if (-not $rows.Count) { throw (T 'Nothing to export yet. Run a read recipe first.' 'Inget resultat att exportera än. Kör ett läsrecept först.') }
$dir = Join-Path $HOME 'Documents/Puppa'; New-Item -ItemType Directory -Force -Path $dir | Out-Null
$path = Join-Path $dir ((($Filnamn -replace '[\\/:*?"<>|]', '_').Trim()) + '.csv')
$rows | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName | Export-Csv -Path $path -Delimiter ';' -Encoding utf8BOM -NoTypeInformation
Write-Host (T "Saved $($rows.Count) rows to $path" "Sparat $($rows.Count) rader i $path")
