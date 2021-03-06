#
# TaskUpdateTenantALChapterSchemaFrom_ChapterHub.ps1
#
<# Header Information **********************************************************
Name: TaskUpdateTenantALChapterSchema.ps1
Created By: Mike John
Created Date: 01/14/2021
Summary:
    Task stub to run UpdateTenantALChapterSchemaFrom_ChapterHub.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 07/06/2021
    Reason Updated: Added $justTesting parameter"
Updated By: Mike John
Updated Date: 02/10/2021
    Reason Updated: Added tenant input parameter
Updated By: Mike John
Updated Date: 01/14/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][bool]$justTesting
      )
begin {}
process {
    #starts here
    #set these to true for production
    [bool]$downLoadChapterHubExportFile = $true;
    [bool]$downLoadALChapterSchemaFile = $true;
    [bool]$upLoadALChapterSchemaFile = $true;
    if($justTesting)
    {
        $downLoadChapterHubExportFile = $false;
        $downLoadALChapterSchemaFile = $false;
        $upLoadALChapterSchemaFile = $false;
    }

    # get tenant specific variable values
    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    #Get Credentials to connect
    [system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    [string]$tenantDomain = $tenantObj.DomainName;

    # create the log file name
    $dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\Update_' + $tenantAbbreviation + '_ALChapterSchema_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

    [string]$downloadPath = $tenantObj.PSStartUpDir + "\" + $tenantAbbreviation + "\ExcelDataFiles";

    [string]$downloadSchemaFileName = $tenantObj.TenantSchemaFileName + "." + $tenantObj.TenantSchemaFileExtension;
    [string]$uploadSchemaFileName = $downloadSchemaFileName; # after updating, return the same file
    [string]$chExportDownloadFileName = "Not Found"; # searched for and then returned; This is the default if it not found
    [string]$baseExportFileName = $tenantObj.UserExportFileName;
    [string]$baseExportFileExtension = $tenantObj.UserExportFileExtension;

    Get-Process -Name "Excel" -ErrorAction Ignore | Stop-Process -Force

    $logMessage = "Starting TaskUpdateTenantChapterSchema.ps1 for Tenant: " + $tenantAbbreviation;
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
        if($tenantAbbreviation -eq "NAL")
        {
            $chExportDownloadFileName = "NAL User Export 20210728 020432.csv";
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

        .\ExcelManagement\UpdateTenantALChapterSchemaFrom_ChapterHub.ps1 -tenantCredentials $psAdminCredentials `
                                                                         -tenantAbbreviation $tenantAbbreviation `
                                                                         -tenantDomain $tenantDomain `
                                                                         -chCSVFilePathAndName $chCSVFilePathAndName `
                                                                         -m365SchemaFilePathAndName $schemaFilePathName `
                                                                         -masterLogFile $myMasterLogFilePathAndName `
                                                                         -testingOnly $justTesting;
    

        #region upload updated AlChapterSchema.xlsx to Teams Storage Site
        if($upLoadALChapterSchemaFile)
        {
            [string]$upLoadSchemaSiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam-M365Data"
            [string]$upLoadschemaFileRelativeURL = "/Shared Documents/M365Data/";
            [string]$upLoadSourceFilePath  = $downloadPath + "\" + $uploadSchemaFileName;

            .\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                      -SharePointSiteURL $upLoadSchemaSiteURL `
                                                      -SharePointFileRelativeURL $upLoadschemaFileRelativeURL `
                                                      -uploadPathAndFileName $upLoadSourceFilePath `
                                                      -masterLogFilePathAndName $myMasterLogFilePathAndName;
        }
        #endregion upload updated AlChapterSchema.xlsx to Teams Storage Site
    
        $logMessage = "Finished TaskUpdateChapterSchema.ps1 for Tenant: " + $tenantAbbreviation;
        .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    }
    else
    {
        if($tenantAbbreviation -eq "NAL")
        {
            # send email notification
            [string]$myEmailSender           = "ITSupport@algeorgetownarea.org";
            [string[]]$myEmailRecipientArray = @("ITSupport@algeorgetownarea.org");
            [string]$myEmailSubject          = "Daily Chapter Hub User Export File Not Found for Tenant: " + $tenantAbbreviation;
            [string]$myEmailBody             = "Check why Chapter Hub User Export File Not Found. - $(Get-Date -Format g)";
            .\EMailer\SendAnEmail.ps1 -tenantCredentials $psAdminCredentials `
                                    -emailSender $myEmailSender `
                                    -emailRecipientArray $myEmailRecipientArray `
                                    -emailSubject $myEmailSubject `
                                    -emailBody $myEmailBody `
                                    -masterLogFilePathAndName $myMasterLogFilePathAndName;
        }
        $logMessage = "Daily Chapter Hub User Export File Not Found for Tenant: " + $tenantAbbreviation;
        .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    }
}

