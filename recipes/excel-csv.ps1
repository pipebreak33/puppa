<#PUPPA
namn:        Exportera senaste resultat till CSV
kategori:    Excel
typ:         lokal
kräver:      ingen
beskrivning: Sparar förra resultatet som CSV med semikolon och UTF-8 (öppnas rätt i svensk Excel).
param:       Filnamn | Filnamn (utan .csv) | {profil}-{datum}
#>
$senaste = Join-Path $env:PUPPA_DIR 'senaste.xml'
if ((-not $global:PuppaLast -or -not @($global:PuppaLast).Count) -and (Test-Path $senaste)) { $global:PuppaLast = Import-Clixml $senaste }
$rows = @($global:PuppaLast | Where-Object { $_ -isnot [string] })
if (-not $rows.Count) { throw 'Inget resultat att exportera än. Kör ett läsrecept först.' }
$dir = Join-Path $HOME 'Documents/Puppa'; New-Item -ItemType Directory -Force -Path $dir | Out-Null
$path = Join-Path $dir ((($Filnamn -replace '[\\/:*?"<>|]', '_').Trim()) + '.csv')
$rows | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName | Export-Csv -Path $path -Delimiter ';' -Encoding utf8BOM -NoTypeInformation
Write-Host "Sparat $($rows.Count) rader i $path"
