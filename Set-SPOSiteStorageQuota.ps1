
<#PSScriptInfo

.VERSION 0.1

.GUID c1f897e0-4ab3-4dfb-9f10-a64300588d09

.AUTHOR June Castillote

.COMPANYNAME

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


#>

<#

.DESCRIPTION
 PowerShell script to update SharePoint Online sites' storage quota limit

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory,
        Position = 0
    )]
    [System.Object[]]
    $Site,

    [Parameter()]
    [bool]
    $TestMode = $true
)
begin {

    # Test if SPO Shell is connected
    try {
        $spoTenant = Get-SpoTenant
    }
    catch {}

    if (!$spoTenant) {
        "The Get-SpoTenant command failed. Please connect to your SharePoint Online organization first using the Connect-SpoService cmdlet." | Out-Default
        Continue
    }


    # Define the required input fields in the $Site object
    $requiredFields = @(
        'Url', 'NewStorageQuota', 'NewStorageQuotaWarningLevel'
    )

    # Inspect the input fields
    ## Initialize the inputPassed flag
    $script:inputPassed = $true

    ## Get the input object properties (NoteProperty)
    $inputProperties = ($Site | Get-Member -MemberType NoteProperty)
    $requiredFields | ForEach-Object {
        ## Check if the required field is present in the input object
        if ($inputProperties.Name -notcontains $_) {
            ## If the required input object field is missing, set the inputPassed flag to $false
            "The required [$_] column is missing." | Out-Default
            $script:inputPassed = $false
        }
    }
    ## If the inputPassed flag is $false, exit the script.
    if (!$script:inputPassed) { continue }

    ## If the inputPassed flag is not $false, resume the script.
    $siteIndex = 0
    $siteCount = $Site.Count
}
process {
    foreach ($currentSite in $Site) {
        try {
            "[$($siteIndex+1) of $($siteCount)]" | Out-Default
            "    Site: $($currentSite.Url)" | Out-Default
            "    Current Usage: $("{0:N0}" -f $currentSite.StorageUsageCurrent)" | Out-Default
            "    New Quota: $("{0:N0}" -f $currentSite.NewStorageQuota)" | Out-Default
            "    New Warning: $("{0:N0}" -f $currentSite.NewStorageQuotaWarningLevel)" | Out-Default
            if ($TestMode -eq $false) {
                if ($currentSite.NewStorageQuota -and $currentSite.NewStorageQuotaWarningLevel) {
                    ## Update only if NewStorageQuota and NewStorageQuotaWarningLevel have values.
                    Set-SPOSite -Identity $currentSite.Url -StorageQuota $currentSite.NewStorageQuota -StorageQuotaWarningLevel $currentSite.NewStorageQuotaWarningLevel -ErrorAction Continue
                    "    Update Status: Successful" | Out-Default
                }
                else {
                    ## Skip the update if NewStorageQuota and NewStorageQuotaWarningLevel don't have values.
                    "    Update Status: Skipped" | Out-Default
                    "    The NewStorageQuota and NewStorageQuotaWarningLevel fields cannot be null or empty. Skipping this site."
                }
            }
            ## If TestMode, do nothing.
            else {
                "    Update Status: Test Mode Only" | Out-Default
            }
        }
        catch {
            ## If the update failed, show error message.
            "    Update Status: Failed" | Out-Default
            "    ERROR: $($_.Exception.Message)" | Out-Default
        }
        $siteIndex++
    }
}
end {

}



