<#PUPPA
namn:        Lägg till i distributionslista
kategori:    Exchange
typ:         ändrar
kräver:      exo
beskrivning: Lägger till en eller flera medlemmar i en distributionslista eller e-postaktiverad säkerhetsgrupp.
param:       Lista | Listans namn eller adress |
param:       Medlemmar | Medlemmar (adress, flera med komma) |
#>
foreach ($m in @($Medlemmar -split '[,;\s]+' | Where-Object { $_ })) {
    Add-DistributionGroupMember -Identity $Lista -Member $m -BypassSecurityGroupManagerCheck -ErrorAction Stop
    Write-Host "  + $m"
}
