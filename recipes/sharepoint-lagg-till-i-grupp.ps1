<#PUPPA
namn:        Lägg till i SharePoint-grupp
kategori:    SharePoint
typ:         ändrar
kräver:      pnp:Url
beskrivning: Lägger till användare i en av sajtens SharePoint-grupper (Ägare/Medlemmar/Besökare …).
param:       Url | Sajtens URL | {rootUrl}/sites/ | @site
param:       Grupp | Gruppens namn (t.ex. 'Projekt Medlemmar') |
param:       Anvandare | Användare (UPN, flera med komma) | | @spuser+
#>
foreach ($a in @($Anvandare -split '[,;\s]+' | Where-Object { $_ })) {
    Add-PnPGroupMember -Group $Grupp -LoginName $a -ErrorAction Stop
    Write-Host "Lade till $a i $Grupp"
}
Get-PnPGroupMember -Group $Grupp | Select-Object Title, Email
