<#PUPPA
namn:        Vem har adminroller?
kategori:    Entra ID
typ:         läs
kräver:      graph:RoleManagement.Read.Directory,User.Read.All
beskrivning: Alla aktiva katalogroller och deras medlemmar (Global admin m.fl.).
#>
Get-MgDirectoryRole -All | ForEach-Object {
    $r = $_
    Get-MgDirectoryRoleMember -DirectoryRoleId $r.Id -All | ForEach-Object {
        [pscustomobject]@{ Roll = $r.DisplayName; Namn = $_.AdditionalProperties.displayName; Konto = $_.AdditionalProperties.userPrincipalName }
    }
} | Sort-Object Roll, Namn
