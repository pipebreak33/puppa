<#PUPPA
namn:        Meddelandespårning
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Har mejlet kommit fram? Söker på avsändare och/eller mottagare (max 10 dagar bakåt).
param:       Avsandare? | Avsändare |
param:       Mottagare? | Mottagare |
param:       Dagar | Hur många dagar bakåt? | 2
#>
$p = @{ StartDate = (Get-Date).AddDays(-[int]$Dagar); EndDate = Get-Date }
if ($Avsandare) { $p.SenderAddress = $Avsandare }
if ($Mottagare) { $p.RecipientAddress = $Mottagare }
$cmd = if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) { 'Get-MessageTraceV2' } else { 'Get-MessageTrace' }
& $cmd @p | Sort-Object Received -Descending |
    Select-Object @{ n = 'Tid'; e = { $_.Received.ToLocalTime().ToString('MM-dd HH:mm') } }, SenderAddress, RecipientAddress, Subject, Status
