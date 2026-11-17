<#PUPPA
namn:        Skapa sajt
kategori:    SharePoint
typ:         ändrar
kräver:      pnp-admin
beskrivning: Ny kommunikationssajt eller teamsajt (Microsoft 365-grupp).
param:       Typ | Vilken sorts sajt? | Kommunikation | Kommunikation,Team
param:       Titel | Sajtens titel |
param:       Alias | Kortnamn i URL:en (/sites/<alias>) |
param:       Agare? | Ägare (UPN) – tomt = du själv | {upn}
#>
if ($Typ -eq 'Team') {
    $p = @{ Type = 'TeamSite'; Title = $Titel; Alias = $Alias }
    if ($Agare) { $p.Owners = $Agare }
} else {
    $p = @{ Type = 'CommunicationSite'; Title = $Titel; Url = "{rootUrl}/sites/$Alias" }
    if ($Agare) { $p.Owner = $Agare }
}
$url = New-PnPSite @p -ErrorAction Stop
[pscustomobject]@{ Titel = $Titel; Typ = $Typ; Url = $url }
