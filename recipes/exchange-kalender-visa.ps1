<#PUPPA
namn:        Vem ser en kalender?
kategori:    Exchange
typ:         läs
kräver:      exo
beskrivning: Visar behörigheterna på en postlådas kalender.
param:       Postlada | Vems kalender? (adress) | | @mailbox
#>
$kal = Get-EXOMailboxFolderStatistics -Identity $Postlada -FolderScope Calendar | Where-Object FolderType -eq 'Calendar' | Select-Object -First 1
Get-EXOMailboxFolderPermission -Identity "${Postlada}:\$($kal.Name)" | Select-Object User, @{ n = 'Åtkomst'; e = { $_.AccessRights -join ', ' } }
