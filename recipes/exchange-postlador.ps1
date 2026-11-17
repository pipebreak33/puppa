<#PUPPA
namn:        Lista postlådor
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Postlådor med typ och (valfritt) storlek. Storlek tar längre tid.
param:       Filter? | Del av namn eller adress |
param:       Storlek | Ta med storlek? | Nej | Nej,Ja
#>
$mb = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails |
    Where-Object { -not $Filter -or $_.DisplayName -like "*$Filter*" -or $_.PrimarySmtpAddress -like "*$Filter*" } |
    Sort-Object DisplayName
foreach ($m in $mb) {
    $o = [ordered]@{ Namn = $m.DisplayName; Adress = $m.PrimarySmtpAddress; Typ = "$($m.RecipientTypeDetails)" }
    if ($Storlek -eq 'Ja') {
        $s = Get-EXOMailboxStatistics -Identity $m.ExternalDirectoryObjectId -ErrorAction SilentlyContinue
        $o.Storlek = if ($s) { ($s.TotalItemSize.ToString() -split ' \(')[0] } else { '' }
    }
    [pscustomobject]$o
}
