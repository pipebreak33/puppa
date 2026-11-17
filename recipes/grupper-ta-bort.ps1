<#PUPPA
namn:        Ta bort ur grupp
kategori:    Grupper
typ:         ändrar
kräver:      graph:GroupMember.ReadWrite.All,User.Read.All
beskrivning: Tar bort användare ur en Microsoft 365- eller säkerhetsgrupp.
param:       Grupp | Gruppens namn eller e-post | | @group
param:       Anvandare | Användare (UPN, flera med komma) | | @user+
#>
$q = $Grupp -replace "'", "''"
$g = @(Get-MgGroup -Filter "displayName eq '$q' or mail eq '$q'" -All)
if (-not $g) { $g = @(Get-MgGroup -Search "displayName:$Grupp" -ConsistencyLevel eventual -All) }
if ($g.Count -ne 1) { throw "Hittade $($g.Count) grupper för '$Grupp': $(($g.DisplayName | Select-Object -First 5) -join ', ')" }
$g = $g[0]

foreach ($a in @($Anvandare -split '[,;\s]+' | Where-Object { $_ })) {
    $u = Get-MgUser -UserId $a -ErrorAction Stop
    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id -ErrorAction Stop
    Write-Host "  - $a ur $($g.DisplayName)"
}
