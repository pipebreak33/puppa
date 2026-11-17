<#PUPPA
namn:        Ge kalenderbehörighet
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Ger en användare behörighet till en kalender (fungerar oavsett språk på mappen).
param:       Postlada | Vems kalender? (adress) | | @mailbox
param:       Anvandare | Vem ska få behörighet? (UPN) | | @recipient
param:       Niva | Nivå | Reviewer | AvailabilityOnly,LimitedDetails,Reviewer,Author,Editor
#>
$kal = Get-EXOMailboxFolderStatistics -Identity $Postlada -FolderScope Calendar | Where-Object FolderType -eq 'Calendar' | Select-Object -First 1
$mapp = "${Postlada}:\$($kal.Name)"
if (Get-EXOMailboxFolderPermission -Identity $mapp -User $Anvandare -ErrorAction SilentlyContinue) {
    Set-MailboxFolderPermission -Identity $mapp -User $Anvandare -AccessRights $Niva -ErrorAction Stop | Out-Null
} else {
    Add-MailboxFolderPermission -Identity $mapp -User $Anvandare -AccessRights $Niva -ErrorAction Stop | Out-Null
}
Get-EXOMailboxFolderPermission -Identity $mapp | Select-Object User, AccessRights
