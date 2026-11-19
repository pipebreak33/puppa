<#PUPPA
namn:        Exportera senaste resultat till Excel
kategori:    Excel
typ:         lokal
kräver:      ingen
beskrivning: Excel-rapport av förra receptets resultat: sammanfattning med sammanställningar (per person, typ, dag …) + formaterad tabell.
param:       Filnamn | Filnamn (utan .xlsx) | {kund}-{datum}
param:       Oppna | Öppna filen efteråt? | Ja | Ja,Nej
param:       Rubrik? | Rapportens rubrik (valfritt) |
param:       Detaljer? | Detaljer, t.ex. vilka val receptet kördes med (valfritt) |
#>
$senaste = Join-Path $env:PUPPA_DIR 'senaste.xml'
# PnP-recept körs i en egen PowerShell-process – filen har alltid det senaste resultatet
if (Test-Path $senaste) { $global:PuppaLast = Import-Clixml $senaste }
if (-not $global:PuppaLast -or -not @($global:PuppaLast).Count) { throw (T 'Nothing to export yet. Run a read recipe first.' 'Inget resultat att exportera än. Kör ett läsrecept först.') }
$rows = @($global:PuppaLast | Where-Object { $_ -isnot [string] })
if (-not $rows.Count) { throw (T 'The last result was only text, nothing to put in a table.' 'Förra resultatet var bara text – inget att lägga i en tabell.') }
$rows = @($rows | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName)
$cols = @($rows[0].PSObject.Properties.Name)

# ImportExcel varnar om autosize på macOS/Linux. Modulfunktioner ser bara den globala inställningen.
$tidigare = $global:WarningPreference; $global:WarningPreference = 'SilentlyContinue'
try {
Import-Module ImportExcel -ErrorAction Stop
$dir = Join-Path $HOME 'Documents/Puppa'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$name = ($Filnamn -replace '[\\/:*?"<>|]', '_').Trim()
if (-not $name) { $name = 'puppa-export' }
$path = Join-Path $dir "$name.xlsx"
if (Test-Path $path) { Remove-Item $path -Force }

# Kolumner: datum, och sådana som går att sammanställa (få olika värden, t.ex. vem/typ/sajt)
function Varde($v) { if ($null -eq $v) { '' } elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { ($v | ForEach-Object { "$_" }) -join ', ' } else { "$v" } }
$datum = @($cols | Where-Object { $c = $_; $f = $rows | Where-Object { $null -ne $_.$c } | Select-Object -First 1; $f -and $f.$c -is [datetime] })
$grupp = foreach ($c in $cols) {
    if ($c -in $datum) { continue }
    $vals = @($rows | ForEach-Object { Varde $_.$c } | Where-Object { $_ -ne '' })
    if ($vals.Count -lt 2) { continue }
    $olika = @($vals | Sort-Object -Unique).Count
    $langd = ($vals | Measure-Object -Property Length -Maximum).Maximum
    # sammanställ bara kolumner som upprepar sig: 2–40 olika värden och korta värden (inte URL:er/id:n)
    if ($olika -ge 2 -and $olika -le 40 -and $olika -lt $vals.Count -and $langd -le 80 -and $vals[0] -notmatch '^https?://|^[0-9a-f-]{36}$') { $c }
}
$grupp = @($grupp | Select-Object -First 6)

# Texts in the report (English or Swedish)
$kF = T 'Field' 'Fält'; $kV = T 'Value' 'Värde'; $kN = T 'Count' 'Antal'; $kA = T 'Share' 'Andel'; $kD = T 'Date' 'Datum'
$wsSum = T 'Summary' 'Sammanfattning'; $wsDay = T 'Per day' 'Per dag'; $wsRes = T 'Result' 'Resultat'
# 1. Sammanfattning
$info = [System.Collections.Generic.List[object]]::new()
if ($Rubrik) { $info.Add([pscustomobject]@{ $kF = (T 'Report' 'Rapport'); $kV = $Rubrik }) }
if ($Detaljer) { foreach ($d in $Detaljer -split ' · ') { $k, $v = $d -split ': ', 2; $info.Add([pscustomobject]@{ $kF = $(if ($v) { $k } else { T 'Choice' 'Val' }); $kV = $(if ($v) { $v } else { $k }) }) } }
$info.Add([pscustomobject]@{ $kF = (T 'Rows' 'Rader'); $kV = $rows.Count })
if ($datum) {
    $t = @($rows | ForEach-Object { $_.($datum[0]) } | Where-Object { $_ -is [datetime] } | Sort-Object)
    if ($t) { $info.Add([pscustomobject]@{ $kF = "Period ($($datum[0]))"; $kV = ('{0:yyyy-MM-dd HH:mm} – {1:yyyy-MM-dd HH:mm}' -f $t[0], $t[-1]) }) }
}
$info.Add([pscustomobject]@{ $kF = (T 'Created' 'Skapad'); $kV = (Get-Date).ToString('yyyy-MM-dd HH:mm') })
$info.Add([pscustomobject]@{ $kF = (T 'Created with' 'Skapad med'); $kV = 'Puppa' })

$titel = if ($Rubrik) { $Rubrik } else { T 'Puppa report' 'Puppa-rapport' }
$info | Export-Excel -Path $path -WorksheetName $wsSum -Title $titel -TitleBold -TitleSize 16 -TableName Info -TableStyle Light1
$rad = 4                                   # sammanställningarna står bredvid, i kolumn E–G
$block = @()
foreach ($c in $grupp) {
    $sum = @($rows | Group-Object { Varde $_.$c } | Sort-Object Count -Descending | Select-Object -First 50 | ForEach-Object {
        [pscustomobject]@{ $c = $(if ($_.Name) { $_.Name } else { T '(empty)' '(tomt)' }); $kN = $_.Count; $kA = [math]::Round($_.Count / $rows.Count, 3) } })
    $tn = "Per_$($block.Count + 1)"
    $sum | Export-Excel -Path $path -WorksheetName $wsSum -StartRow $rad -StartColumn 5 -TableName $tn -TableStyle Medium2
    $block += [pscustomobject]@{ Rad = $rad - 1; Rubrik = "Per $($c.ToLower())"; Forsta = $rad + 1; Sista = $rad + $sum.Count }
    $rad += $sum.Count + 4
}
if ($datum -and $rows.Count -gt 1) {
    $dag = @($rows | Where-Object { $_.($datum[0]) -is [datetime] } | Group-Object { $_.($datum[0]).ToString('yyyy-MM-dd') } |
             Sort-Object Name -Descending | ForEach-Object { [pscustomobject]@{ $kD = $_.Name; $kN = $_.Count } })
    if ($dag.Count -gt 1) {
        $dag | Export-Excel -Path $path -WorksheetName $wsDay -TableName PerDag -TableStyle Medium2 -FreezeTopRow -BoldTopRow
    }
}

# 2. Resultatet
$rows | Export-Excel -Path $path -WorksheetName $wsRes -TableName Resultat -TableStyle Medium2 -FreezeTopRow -BoldTopRow

# 3. Formatering: rubriker, bredder, datum, procent
$pkg = Open-ExcelPackage -Path $path
$ws = $pkg.Workbook.Worksheets[$wsSum]
foreach ($b in $block) { $ws.Cells[$b.Rad, 5].Value = $b.Rubrik; $ws.Cells[$b.Rad, 5].Style.Font.Bold = $true; $ws.Cells[$b.Rad, 5].Style.Font.Size = 13
                         if ($b.Sista -ge $b.Forsta) { $ws.Cells["G$($b.Forsta):G$($b.Sista)"].Style.Numberformat.Format = '0.0%' } }
$ws.Column(1).Width = 22; $ws.Column(2).Width = 60; $ws.Column(3).Width = 3; $ws.Column(4).Width = 3
$ws.Column(5).Width = 36; $ws.Column(6).Width = 9; $ws.Column(7).Width = 9
$ws = $pkg.Workbook.Worksheets[$wsRes]
for ($i = 0; $i -lt $cols.Count; $i++) {
    $c = $cols[$i]
    $max = ($rows | Select-Object -First 500 | ForEach-Object { (Varde $_.$c).Length } | Measure-Object -Maximum).Maximum
    $ws.Column($i + 1).Width = [math]::Min(70, [math]::Max(10, [math]::Max($c.Length + 4, $max + 2)))
    if ($c -in $datum) { $ws.Column($i + 1).Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'; $ws.Column($i + 1).Width = 18 }
}
if ($pkg.Workbook.Worksheets[$wsDay]) { $pkg.Workbook.Worksheets[$wsDay].Column(1).Width = 14 }
$pkg.Workbook.Worksheets[$wsSum].Select()
Close-ExcelPackage $pkg
} finally { $global:WarningPreference = $tidigare }

Write-Host (T "Saved $($rows.Count) rows to $path$(if ($grupp) { " · summarised per $(($grupp | ForEach-Object { $_.ToLower() }) -join ', ')" })" "Sparat $($rows.Count) rader i $path$(if ($grupp) { " · sammanställt per $(($grupp | ForEach-Object { $_.ToLower() }) -join ', ')" })")
if ($Oppna -eq 'Ja') { Invoke-Item -LiteralPath $path }
