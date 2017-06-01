$allGPOs = get-gpo -all
foreach ($gpo in $allGPOs)
{
    #first check to see if GPO has per-user settings, since this fix really only needs to apply to per user GPOs. Remove this check if you really want to modify all GPOs
    if ($gpo.user.DSVersion -gt 0)
    {
        # first read the GPO permissions to find out if Authn Users and Domain Computers is missing. Note--depending upon the version of Windows/GPMC you are on--Get-GPPermission might be Get-GPPermissionS
        $perm1 = Get-GPPermission -Guid $gpo.id -TargetName "Authenticated Users" -TargetType group -ErrorAction SilentlyContinue
        $perm2 = Get-GPPermission -Guid $gpo.id -TargetName "Domain Computers" -TargetType group -ErrorAction SilentlyContinue
        if ($perm1 -eq $null -and $perm2 -eq $null) # if no authn users or domain computers is found, then add Authn Users read perm
        {
            Set-GPPermission -Guid $gpo.Id -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group
            Write-Host $gpo.DisplayName "has been modified to grant Authenticated Users read access"
        }
    }

}