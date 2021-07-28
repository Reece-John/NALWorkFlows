#
# DownLoadSharePointFile.ps1
#
<# Header Information **********************************************************
Name: DownLoadSharePointFile.ps1
Created By: Mike John
Created Date: 08/16/2020
Summary:
    Copies a SharePoint file to the local machine
     Remember OneDrive and Teams files are SharePoint files
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 08/16/2020
    Reason Updated: original version
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
process
{
    Try
    {
        [bool]$connectionMade = $false;
        [string]$logMessage = "Starting DownLoadSharePointFIle: " + $FileName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $urlDownload = $SharePointSiteURL + $SharePointFileRelativeURL;
        #Connect to PNP On-line
        Connect-PnPOnline -Url $SharePointSiteURL -Credentials $tenantCredentials;
        $connectionMade = $true;

        #PowerShell download file from SharePoint on-line
        #Get-PnPFile -Url $SharePointFileRelativeURL -Path $LocalFileDownloadPath -FileName $FileName -AsFile -ThrowExceptionIfFileNotFound
        Get-PnPFile -Url $SharePointFileRelativeURL -Path $LocalFileDownloadPath -FileName $FileName -AsFile -Force;
        Disconnect-PnPOnline;
    }
    catch 
    {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
        $logMessage = "Error: $($_.Exception.Message)";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $FileName = "Error ******* File Not Copied **********";
        if($connectionMade)
        {
            Disconnect-PnPOnline;
        }
    }
    $logMessage = "Finished DownLoadSharePointFIle: " + $FileName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
