<#PUPPA
namn:        Exportera senaste resultat till Excel
kategori:    Excel
typ:         lokal
kräver:      ingen
beskrivning: Sparar det förra receptets resultat som en formaterad .xlsx (ImportExcel).
param:       Filnamn | Filnamn (utan .xlsx) | {kund}-{datum}
param:       Oppna | Öppna filen efteråt? | Ja | Ja,Nej
#>
$senaste = Join-Path $env:PUPPA_DIR 'senaste.xml'
if ((-not $global:PuppaLast -or -not @($global:PuppaLast).Count) -and (Test-Path $senaste)) { $global:PuppaLast = Import-Clixml $senaste }
if (-not $global:PuppaLast -or -not @($global:PuppaLast).Count) { throw 'Inget resultat att exportera än. Kör ett läsrecept först.' }
$dir = Join-Path $HOME 'Documents/Puppa'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$name = ($Filnamn -replace '[\\/:*?"<>|]', '_').Trim()
if (-not $name) { $name = 'puppa-export' }
$path = Join-Path $dir "$name.xlsx"
$rows = @($global:PuppaLast | Where-Object { $_ -isnot [string] })
if (-not $rows.Count) { throw 'Förra resultatet var bara text – inget att lägga i en tabell.' }
$rows = $rows | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName
$rows | Export-Excel -Path $path -WorksheetName 'Puppa' -FreezeTopRow -BoldTopRow -TableStyle Medium2 -ClearSheet
Write-Host "Sparat $($rows.Count) rader i $path"
if ($Oppna -eq 'Ja') { Invoke-Item -LiteralPath $path }
