<#PUPPA
namn:        Inloggningsmetoder (MFA)
kategori:    Användare
typ:         läs
kräver:      graph:UserAuthenticationMethod.Read.All
beskrivning: Vilka MFA-metoder en användare har registrerat.
param:       Anvandare | Användare (UPN) |
#>
Get-MgUserAuthenticationMethod -UserId $Anvandare -All | ForEach-Object {
    $a = $_.AdditionalProperties
    [pscustomobject]@{ Metod = ($a.'@odata.type' -replace '#microsoft.graph.', '' -replace 'AuthenticationMethod', ''); Detalj = $a.displayName ?? $a.phoneNumber ?? $a.emailAddress ?? '' }
}
