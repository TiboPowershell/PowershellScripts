$Global:TenantId = "<TenantID>"
$Global:ClientId = "<ClientID = Application ID>"
$Global:clientSecret = "<ClientSecret>"
$ExportToExcelPath = "<Path for Excel file ending in .xlsx>"

$SecuredPasswordPassword = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $SecuredPasswordPassword

#Connect MgGraph
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential

function Get-GroupsDictionary {
    $groupDictionary = @{}
    $groups = Get-MgBetaGroup -All -Property Id, DisplayName
    
    foreach ($group in $groups) {
        $groupDictionary[$group.Id] = $group.DisplayName
    }

    return $groupDictionary
}

## NOTE we do this because if you want to use $Assignments = Get-MgBetaEntitlementManagementAccessPackageAssignment it will only work for access packages that have a user assigned to them
function Get-ResourcesFromAccessPackages{
    param (
        [hashtable]$GroupDictionary
    )

    $accessPackageCatalogs = Get-MgBetaEntitlementManagementAccessPackageCatalog -All
    $exportList = @()
    $totalCatalogs = $accessPackageCatalogs.Count
    foreach ($catalog in $accessPackageCatalogs) {
        Write-Host "[$($accessPackageCatalogs.IndexOf($catalog) + 1)/$($totalCatalogs)][Catalog: $($catalog.DisplayName)]"

        ##get resource from catalog
        $resources = Get-MgBetaEntitlementManagementAccessPackageCatalogAccessPackageResource -AccessPackageCatalogId $catalog.Id -ExpandProperty *
        ## get all access packages within this resource
        $accessPackages = Get-MgBetaEntitlementManagementAccessPackage -CatalogId $catalog.Id -ExpandProperty AccessPackageResourceRoleScopes

        $totalAccessPackagesInCatalog = $accessPackages.count
        foreach($accessPackage in $accessPackages){
            Write-Host "`t[$($accessPackages.IndexOf($accessPackage) + 1)/$($totalAccessPackagesInCatalog)][Access Package: $($accessPackage.DisplayName)]"
            $roleIDs = $accessPackage.AccessPackageResourceRoleScopes.Id | ForEach-Object {($_ -split '_')[0]} 
            foreach($roleID in $roleIDs){
                ##match the roleIDs with $resources.AccessPackageResourceRoles.ID to get the origin ID (we split it this with underscore since this value is prefixed with Member or Owner but we dont need this)
                $matchedRole = (($resources.AccessPackageResourceRoles | Where-Object {$_.id -eq $roleID}).OriginId -split '_')[1]
                ##match this with our GroupDictionary (to make sure we get the correct name, if a groupname is changed and the group is not refreshed from origin in the catalog, it will show the old name)
                $exportList += [PSCustomObject][Ordered]@{
                    Catalog = $catalog.DisplayName
                    CatalogID = $catalog.id
                    AccessPackage = $accessPackage.DisplayName
                    AccessPackageID = $accessPackage.Id
                    GroupDisplayname = $GroupDictionary[$matchedRole]
                    GroupID  = $matchedRole
                }
            }
        }
    }
    return $exportList
}

$groupsDictionary = Get-GroupsDictionary
$getAllAccessPackagesWithResources = Get-ResourcesFromAccessPackages -GroupDictionary $groupsDictionary
##export values to Excel or Csv 
Export-Excel -Path $ExportToExcelPath -InputObject $getAllAccessPackagesWithResources  -WorksheetName "AccessPackageResources" -TableStyle Light1 -TableName "Results"




