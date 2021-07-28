#
# TaskMembersInSchemaNotInCHExport.ps1
#
<# Header Information **********************************************************
Name: TaskUpdateTenantALChapterSchema.ps1
Created By: Mike John
Created Date: 06/16/2021
Summary:
    Task stub to run MembersInSchemaNotInCHExport.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 06/16/2021
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("ALGA","ALSA","ALLV")][string]$tenantAbbreviation
)
begin {}
process
{
    #starts here
    [bool]$downLoadChapterHubExportFile = $true;
    [bool]$downLoadALChapterSchemaFile = $true;

    # get tenant specific variable values
    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    #Get Credentials to connect
    [system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    [string]$tenantDomain = $tenantObj.DomainName;

    # create the log file name
    $dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\Update_' + $tenantAbbreviation + '_ALChapterSchema_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

    [string]$downloadPath = $tenantObj.PSStartUpDir + "\" + $tenantAbbreviation + "\ExcelDataFiles";

    [string]$downloadSchemaFileName = "ALChapterSchema.xlsx";

    [string]$chExportDownloadFileName = "Not Found"; # searched for and then returned; This is the default if it not found
    [string]$baseExportFileName = $tenantObj.UserExportFileName;
    [string]$baseExportFileExtension = "csv";

    Get-Process -Name "Excel" -ErrorAction Ignore | Stop-Process -Force

    $logMessage = "Starting TaskMembersInSchemaNotInCHExport.ps1 for Tenant: " + $tenantAbbreviation;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    $logMessage = "Processing Tenant: " + $tenantAbbreviation;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    #region get name of latest copy of Chapter Hub user CSV export file
    #download Latest Chapter hub export file
    # ******************** This has to be reworked *********************

    [string]$SharePointSiteURL = "https://" + $tenantDomain + "-my.sharepoint.com/personal/data_" + $tenantDomain + "_org";

    [string]$SharePointFileRelativeURL = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/";
    [string]$ListName = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/";

    if($downLoadChapterHubExportFile)
    {

        $chExportDownloadFileName = .\Utilities\GetTodaysLatestCHUserExportFile.ps1 -tenantCredentials $psAdminCredentials `
                                                                                    -sharePointSiteURL $SharePointSiteURL `
                                                                                    -sharePointFileRelativeURL $SharePointFileRelativeURL `
                                                                                    -sharePointListName $ListName `
                                                                                    -baseExportFileName $baseExportFileName `
                                                                                    -exportFileExtension $baseExportFileExtension `
                                                                                    -masterLogFile $myMasterLogFilePathAndName;

    }
    else
    {
        if($tenantAbbreviation -eq "ALSA")
        {
            $chExportDownloadFileName = "ALSA User Export 20210106 120800.csv";
        }
    }
    #endregion get name of latest copy of Chapter Hub user CSV export file

    # if not found then LOG and bail out
    if($chExportDownloadFileName -ne "File Not Found" -and $chExportDownloadFileName -ne "Not Found" -and $chExportDownloadFileName -ne "Error")
    {
        #region download latest copy of Chapter Hub user CSV export file
        [string]$SharePointCSVSiteURL = "https://"+ $tenantDomain + "-my.sharepoint.com/personal/data_" + $tenantDomain + "_org";
        #[string]$csvSiteRelativeUrl = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/" + $chExportDownloadFileName;
        #[string]$csvSiteRelativeUrl = "%2FDocuments%2F" + $tenantAbbreviation + "%2FChapter%20Hub%20Exports%2F" + $chExportDownloadFileName;
        [string]$csvSiteRelativeUrl = "/Documents/" + $tenantAbbreviation + " Chapter Hub Exports/" + $chExportDownloadFileName;
        if($downLoadChapterHubExportFile)
        {
            .\M365SharePoint\DownLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                        -SharePointSiteURL $SharePointCSVSiteURL `
                                                        -SharePointFileRelativeURL $csvSiteRelativeUrl `
                                                        -LocalFileDownloadPath $downloadPath `
                                                        -FileName $chExportDownloadFileName `
                                                        -masterLogFilePathAndName $myMasterLogFilePathAndName;
        }

        #endregion download latest copy of Chapter Hub user CSV export file

        # delete current copy of $downloadSchemaFileName

        #region download latest copy of Chapter Schema Excel Spreadsheet from SharePoint
        [string]$downLoadSchemaSiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam-M365Data"
        [string]$downLoadschemaFileRelativeURL = "/sites/TechnologyTeam-M365Data/Shared Documents/M365Data/" + $downloadSchemaFileName;

        if($downLoadALChapterSchemaFile)
        {
        .\M365SharePoint\DownLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                    -SharePointSiteURL $downLoadSchemaSiteURL `
                                                    -SharePointFileRelativeURL $downLoadschemaFileRelativeURL `
                                                    -LocalFileDownloadPath $downloadPath `
                                                    -FileName $downloadSchemaFileName `
                                                    -masterLogFilePathAndName $myMasterLogFilePathAndName;
        }
        #endregion download latest copy of Chapter Schema Excel Spreadsheet from SharePoint

        # Update ALChapterSchema.xlsx from Chapter Hub CSV file
        # Assign values to variables
        [string]$chCSVFilePathAndName = $DownloadPath + "\" + $chExportDownloadFileName;
        [string]$schemaFilePathName = $DownloadPath + "\" + $downloadSchemaFileName;

        .\ExcelManagement\MembersInSchemaNotInCHExport.ps1 -tenantCredentials $psAdminCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -chCSVFilePathAndName $chCSVFilePathAndName `
                                                           -m365SchemaFilePathAndName $schemaFilePathName `
                                                           -masterLogFile $myMasterLogFilePathAndName;
    }
}
