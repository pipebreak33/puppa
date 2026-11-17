<#PUPPA
namn:        Papperskorgen på en sajt
kategori:    SharePoint
typ:         läs
kräver:      pnp:Url
beskrivning: Visar raderade filer (första och andra steget), senaste först.
param:       Url | Sajtens URL | {rootUrl}/sites/
param:       Filter? | Del av filnamn |
param:       Antal | Max antal rader | 100
#>
Get-PnPRecycleBinItem -RowLimit 5000 | Where-Object { -not $Filter -or $_.LeafName -like "*$Filter*" } |
    Sort-Object DeletedDate -Descending | Select-Object -First ([int]$Antal) |
    Select-Object LeafName, DirName, @{ n = 'Raderad'; e = { $_.DeletedDate.ToString('yyyy-MM-dd HH:mm') } }, DeletedByEmail, ItemState
