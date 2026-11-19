<#PUPPA
namn:        Lägg till i grupp
kategori:    Grupper
typ:         ändrar
kräver:      graph:GroupMember.ReadWrite.All,User.Read.All
beskrivning: Lägger till användare i en Microsoft 365- eller säkerhetsgrupp. (Distributionslistor: använd Exchange-receptet.)
param:       Grupp | Gruppens namn eller e-post | | @group
param:       Anvandare | Användare (UPN, flera med komma) | | @user+
param:       Roll | Som | Medlem | Medlem,Ägare
#>
$q = $Grupp -replace "'", "''"
$g = @(Get-MgGroup -Filter "displayName eq '$q' or mail eq '$q'" -All)
if (-not $g) { $g = @(Get-MgGroup -Search "displayName:$Grupp" -ConsistencyLevel eventual -All) }
if ($g.Count -ne 1) { throw (T "Found $($g.Count) groups for '$Grupp': $(($g.DisplayName | Select-Object -First 5) -join ', ')" "Hittade $($g.Count) grupper för '$Grupp': $(($g.DisplayName | Select-Object -First 5) -join ', ')") }
$g = $g[0]

foreach ($a in @($Anvandare -split '[,;\s]+' | Where-Object { $_ })) {
    $u = Get-MgUser -UserId $a -ErrorAction Stop
    if ($Roll -eq 'Ägare') { New-MgGroupOwnerByRef -GroupId $g.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($u.Id)" } -ErrorAction Stop }
    else { New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $u.Id -ErrorAction Stop }
    Write-Host (T "  + $a as $($Roll.ToLower()) in $($g.DisplayName)" "  + $a som $($Roll.ToLower()) i $($g.DisplayName)")
}
