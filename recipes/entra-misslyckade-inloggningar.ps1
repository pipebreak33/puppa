<#PUPPA
namn:        Misslyckade inloggningar i tenanten
kategori:    Entra ID
typ:         läs
kräver:      graph:AuditLog.Read.All,Directory.Read.All
beskrivning: Vem misslyckas med att logga in, hur ofta, varför och varifrån – flest först (kräver Entra ID P1).
param:       Timmar | Hur många timmar bakåt? | 24
#>
$fran = (Get-Date).ToUniversalTime().AddHours(-[int]$Timmar).ToString('yyyy-MM-ddTHH:mm:ssZ')
Get-MgAuditLogSignIn -Filter "createdDateTime ge $fran and status/errorCode ne 0" -All |
    Group-Object UserPrincipalName | Sort-Object Count -Descending | ForEach-Object {
        $g = $_.Group
        [pscustomobject]@{
            Användare = $_.Name; Försök = $_.Count
            Orsaker = (($g | Group-Object { "$($_.Status.ErrorCode) $($_.Status.FailureReason)" } | Sort-Object Count -Descending |
                        Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ' | ')
            'IP-adresser' = (($g.IPAddress | Sort-Object -Unique | Select-Object -First 5) -join ', ')
            Länder = (($g.Location.CountryOrRegion | Where-Object { $_ } | Sort-Object -Unique) -join ', ')
            Senast = ($g.CreatedDateTime | Measure-Object -Maximum).Maximum.ToLocalTime()
        }
    }
