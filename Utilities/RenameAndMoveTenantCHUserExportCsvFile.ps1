#
# RenameAndMoveTenantCHUserExportCsvFile.ps1
#
<# Header Information **********************************************************
Name: RenameAndMoveChapterHubCSVFile.ps1
Created By: Mike John
Created Date: 11/07/2020
Summary:
    Renames and Moves $tenantAbbreviation + " Nightly User Export.csv" to .\"Chapter Hub Exports" folder
    Moves from User Data's one drive to another folder
    This file is deposited there by a business automation flow 
     that monitors the "Data@" + $tenantDomain + ".org" mail box for an email with an attachment and 
     sweeps it to "/Documents/Email attachments from Flow/" folder
Update History *****************************************************************
Updated By: Mike John
Updated Date: 02/14/2020
    Reason Updated: Parameterized $tenantAbbreviation and $tenantDomain 
                     to make it work for more than one tenant
Update History *****************************************************************
Updated By: Mike John
Updated Date: 12/27/2020
    Reason Updated: rename script; Added call to InsertOrUpdateStatusRptList.ps1
Updated By: Mike John
UpdatedDate: 11/07/2020
    Reason Updated: original version
#>

<#
 This will rename a SharePoint or OneDrive file and move it to the archive directory
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2)][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3)][string]$masterLogFilePathAndName
)
begin {}
process {
    [string]$ExportFileName = $tenantAbbreviation + " Nightly User Export";
    [string]$ExportFileExtension = "csv";
    [string]$ExportFileNameFullName = $ExportFileName + "." + $ExportFileExtension

    [string]$SiteURL = "https://" + $tenantDomain + "-my.sharepoint.com/personal/data_" + $tenantDomain + "_org";
    [string]$siteRelativeUrl = "/Documents/Email attachments from Flow/" + $ExportFileNameFullName;



    [string]$logMessage = "Starting RenameAndMoveChapterHubCSVFile.ps1";
    .\LogManagement\WriteToLogFile.ps1 -logFile $masterLogFilePathAndName -message $logMessage;

    [string]$ReturnFileName = $ExportFileNameFullName;

    #region updating APStatusTable with Start values
    [string]$ProcessName = "RenameAndMoveTenantCHUserExportCsvFile";
    [string]$ProcessCategory = "CHWorkFlow";
    [DateTime]$ProcessStartDate = Get-date;
    #[DateTime]$ProcessStartDate = (Get-date).ToUniversalTime();
    [DateTime]$ProcessStopDateBegins = (Get-date("1/1/1900 00:01")).ToUniversalTime();
    [DateTime]$SProcessStopDateFinished =( Get-date("1/1/2099 00:01")).ToUniversalTime();
    [string]$ProcessStatusStart = "Started";
    [string]$ProcessStatusFinish = "Successful";
    [string]$ProcessProgressStart = "In-progress";
    [string]$ProcessProgressFinish = "Completed";

    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $myCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $ProcessStopDateBegins `
                                                           -ProcessStatus $ProcessStatusStart `
                                                           -ProcessProgress $ProcessProgressStart `
                                                           -masterLogFilePathAndName $masterLogFilePathAndName;
    #endregion


    Try
    {
        #Connect to PNP Online
        Connect-PnPOnline -Url $SiteURL -Credentials $tenantCredentials

        $fileToGet = $null;
        $fileToGet = Get-PnPFile -Url $siteRelativeUrl -ErrorAction SilentlyContinue
        if($null -ne $fileToGet)
        {
            [datetime]$createTime = $fileToGet.TimeCreated;
            $localTime = $createTime.ToLocalTime();
            $targetFileName = $ExportFileName + $localTime.ToString(" yyyyMMdd HHmmss") + "." + $ExportFileExtension;
            $logMessage = "Renaming File to: " + $targetFileName
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

            # PowerShell rename file in SharePoint on-line (or OneDrive on-line)
            Rename-PnPFile -SiteRelativeUrl $siteRelativeUrl -TargetFileName $targetFileName -Force;

            $newSiteRelativeUrl = "/Documents/Email attachments from Flow/" + $targetFileName;
            $targetUrl = "/personal/data_" + $tenantDomain + "_org/Documents/Chapter Hub Exports/" + $targetFileName;
            $logMessage = "Moving to: " + $targetUrl
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            Move-PnPFile -SiteRelativeUrl "$newSiteRelativeUrl" -TargetUrl $targetUrl -OverwriteIfAlreadyExists -Force
        }
        else
        {
            $logMessage = "No file found to rename and move.";
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            $ReturnFileName = "FileNotFound";
            $ProcessStatusFinish = "Failed";
            $ProcessProgressFinish = "Error";
        }
    }
    catch
    {
        $logMessage = "Error: $($_.Exception.Message)";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $ReturnFileName = "ErrorFound";
        $ProcessStatusFinish = "Failed";
        $ProcessProgressFinish = "Error";
    }
    $logMessage = "Finished RenameAndMoveTenantCHUserExportCsvFile.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    #region updating APStatusTable with finish values
    $SProcessStopDateFinished = Get-Date;
    #$SProcessStopDateFinished = (Get-Date).ToUniversalTime();
    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $myCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $SProcessStopDateFinished `
                                                           -ProcessStatus $ProcessStatusFinish `
                                                           -ProcessProgress $ProcessProgressFinish `
                                                           -masterLogFilePathAndName $masterLogFilePathAndName;
    #endregion
    #Disconnect-PnPOnline -ErrorAction SilentlyContinue;

    return $ReturnFileName;
}