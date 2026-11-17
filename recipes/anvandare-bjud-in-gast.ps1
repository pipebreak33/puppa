<#PUPPA
namn:        Bjud in gäst
kategori:    Användare
typ:         ändrar
kräver:      graph:User.Invite.All
beskrivning: Bjuder in en extern användare som gäst, med eller utan inbjudningsmejl.
param:       Epost | Gästens e-post |
param:       Namn? | Visningsnamn |
param:       Skicka | Skicka inbjudningsmejl? | Ja | Ja,Nej
#>
$p = @{ InvitedUserEmailAddress = $Epost; InviteRedirectUrl = 'https://myapplications.microsoft.com'; SendInvitationMessage = ($Skicka -eq 'Ja') }
if ($Namn) { $p.InvitedUserDisplayName = $Namn }
New-MgInvitation @p -ErrorAction Stop | Select-Object InvitedUserEmailAddress, Status, InviteRedeemUrl
