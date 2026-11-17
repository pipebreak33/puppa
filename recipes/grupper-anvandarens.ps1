<#PUPPA
namn:        Vilka grupper är en användare med i?
kategori:    Grupper
typ:         läs
kräver:      graph:Group.Read.All,User.Read.All
beskrivning: Alla grupper (och roller) en användare är direkt medlem i.
param:       Anvandare | Användare (UPN) | | @user
#>
Get-MgUserMemberOf -UserId $Anvandare -All | ForEach-Object {
    $t = $_.AdditionalProperties.'@odata.type'
    [pscustomobject]@{ Namn = $_.AdditionalProperties.displayName; Typ = if ($t -like '*directoryRole') { 'Admin-roll' } elseif ($_.AdditionalProperties.groupTypes -contains 'Unified') { 'Microsoft 365' } else { 'Grupp' }; Mail = $_.AdditionalProperties.mail }
} | Sort-Object Typ, Namn
