#
# DownloadTodaysCHExportFile.ps1
#
<#
#           Author: Mike John
#     Date Created: 11/02/2020
#
# Last date edited: 11/21/2020
#   Last edited By: Mike John
# Last Edit Reason: Updated log file messages
#
# Last date edited: 11/02/2020
#   Last edited By: Mike John
# Last Edit Reason: Original
#
Preconditions:
    $siteURL must exist
    $siteRelativeURL must exist
    $masterLogFile directory must exist

Download most recent file from today only

Return file name or "Not Found"

#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$SharePointSiteURL
     ,[Parameter(Mandatory=$True,Position=2)][string]$SharePointFileRelativeURL
     ,[Parameter(Mandatory=$True,Position=3)][string]$LocalFileDownloadPath
     ,[Parameter(Mandatory=$True,Position=4)][string]$FileName
     ,[Parameter(Mandatory=$True,Position=5)][string]$masterLogFilePathAndName
)
begin {}
process {
    Try {
        [string]$logMessage = "Starting DownLoadSharePointFIle: " + $FileName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        #Connect to PNP On-line
        Connect-PnPOnline -Url $SharePointSiteURL -Credentials $tenantCredentials

        #PowerShell download file from SharePoint on-line
        Get-PnPFile -Url $SharePointFileRelativeURL -Path $LocalFileDownloadPath -AsFile -Force
    }
    catch 
    {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
        $logMessage = "Error: $($_.Exception.Message)";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    }
    $logMessage = "Finished DownLoadSharePointFIle: " + $FileName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
