<#PUPPA
namn:        Villkorlig åtkomst (policyer)
kategori:    Entra ID
typ:         läs
kräver:      graph:Policy.Read.All
beskrivning: Alla Conditional Access-policyer: läge (på/av/bara rapport), vilka som omfattas och undantas, appar och krav.
#>
$namn = @{}
function Namn($id, $typ) {
    if ($id -in 'All', 'None', 'GuestsOrExternalUsers', 'Office365', 'MicrosoftAdminPortals') { return $id }
    if (-not $namn[$id]) {
        $namn[$id] = try {
            switch ($typ) {
                'user'  { (Get-MgUser -UserId $id -Property userPrincipalName -ErrorAction Stop).UserPrincipalName }
                'group' { (Get-MgGroup -GroupId $id -Property displayName -ErrorAction Stop).DisplayName }
                'role'  { (Get-MgDirectoryRoleTemplate -DirectoryRoleTemplateId $id -ErrorAction Stop).DisplayName }
                'app'   { (Get-MgServicePrincipal -Filter "appId eq '$id'" -Property displayName -ErrorAction Stop | Select-Object -First 1).DisplayName }
            }
        } catch { $id }
        if (-not $namn[$id]) { $namn[$id] = $id }
    }
    $namn[$id]
}
function Lista($ids, $typ) { (@($ids) | Where-Object { $_ } | ForEach-Object { Namn $_ $typ }) -join ', ' }
Get-MgIdentityConditionalAccessPolicy -All | Sort-Object DisplayName | ForEach-Object {
    $u = $_.Conditions.Users
    [pscustomobject]@{
        Policy = $_.DisplayName
        Läge = switch ($_.State) { 'enabled' { 'på' } 'disabled' { 'av' } 'enabledForReportingButNotEnforced' { 'bara rapport' } default { $_ } }
        Omfattar = (@((Lista $u.IncludeUsers 'user'), (Lista $u.IncludeGroups 'group'), (Lista $u.IncludeRoles 'role')) | Where-Object { $_ }) -join ' · '
        Undantar = (@((Lista $u.ExcludeUsers 'user'), (Lista $u.ExcludeGroups 'group'), (Lista $u.ExcludeRoles 'role')) | Where-Object { $_ }) -join ' · '
        Appar = Lista $_.Conditions.Applications.IncludeApplications 'app'
        Krav = (@($_.GrantControls.BuiltInControls) + @($(if ($_.GrantControls.AuthenticationStrength.DisplayName) { $_.GrantControls.AuthenticationStrength.DisplayName })) |
                Where-Object { $_ }) -join " $($_.GrantControls.Operator) "
        Ändrad = $_.ModifiedDateTime
    }
}
