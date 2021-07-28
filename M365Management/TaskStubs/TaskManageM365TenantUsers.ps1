#
# TaskManageM365TenantUsers.ps1
#

<# Header Information **********************************************************
Name: TaskManageM365TenantUsers.ps1
Created By: Mike John
Created Date: 02/20/2021
Summary:
    Task stub to run ManageM365TneantUsers.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 07/06/2021
    Reason Updated: Added $justTesting parameter"
Updated By: Mike John
Updated Date: 02/20/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][bool]$justTesting
      )
begin {}
process {
    Import-Module MSOnline -UseWindowsPowerShell;

    [bool]$downLoadALChapterSchemaFile = $true;

    # get tenant specific variable values
    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    # get administrator credentials
    [system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    [string]$tenantDomain = $tenantObj.DomainName;

    # create the log file name
    $dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = "c:\logs\"+ $tenantAbbreviation +"\ManageM365_" + $tenantAbbreviation + "_Users_" + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

    # SharePoint File download variables
    $downLoadFileName = "ALChapterSchema.xlsx";

    [string]$downloadPath = $tenantObj.PSStartUpDir + "\" + $tenantAbbreviation + "\ExcelDataFiles";

    $logMessage = "Starting TaskManageM365TenantUsers.ps1";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    # download ALChapterSchema.xlsx
    #Configuration Variables
    $SiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam-M365Data";
    $FileRelativeURL = "/sites/TechnologyTeam-M365Data/Shared Documents/M365Data/" + $downLoadFileName;


    if($downLoadALChapterSchemaFile)
    {
        .\M365SharePoint\DownLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                    -SharePointSiteURL $SiteURL `
                                                    -SharePointFileRelativeURL $FileRelativeURL `
                                                    -LocalFileDownloadPath $DownloadPath `
                                                    -FileName $downLoadFileName `
                                                    -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }

    [string]$alChapterSchemaFilePathAndName = $DownloadPath + "\" + $downLoadFileName;
    #set this to true to just write out what would be changed

    .\M365Management\ManageM365TenantUsers.ps1 -tenantCredentials $psAdminCredentials `
                                               -tenantAbbreviation $tenantAbbreviation `
                                               -tenantObj $tenantObj `
                                               -alChapterSchemaFilePathAndName $alChapterSchemaFilePathAndName `
                                               -masterLogFilePathAndName $myMasterLogFilePathAndName `
                                               -justTestingOnly $justTesting;

    # remove $downLoadFileName

    $logMessage = "Finished TaskManageM365Users.ps1";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

}