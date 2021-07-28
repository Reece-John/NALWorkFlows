#
# GetTodaysLatestTenantCHUserExportFile.ps1
#
<# Header Information **********************************************************
Name: GetTodaysLatestTenantCHUserExportFile.ps1
Created By: Mike John
Created Date: 03/20/2021
Summary:
    Copies a latest Chapter Hub user Export file to the local machine
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 03/20/2021
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$sharePointSiteURL
     ,[Parameter(Mandatory=$True,Position=2)][string]$sharePointFileRelativeURL
     ,[Parameter(Mandatory=$True,Position=3)][string]$sharePointListName
     ,[Parameter(Mandatory=$True,Position=4)][string]$baseExportFileName
     ,[Parameter(Mandatory=$True,Position=5)][string]$exportFileExtension
     ,[Parameter(Mandatory=$True,Position=6)][string]$masterLogFilePathAndName
)
begin {}
process {
    #region function definitions
    function Convert-UTCtoLocal
    { param( [parameter(Mandatory=$true)] $UTCTime )
        [bool]$isDayLightSavingTime = (Get-Date).IsDaylightSavingTime();
        $timeZoneInfo = [TimeZoneInfo]::Local;
        $baseTimeZoneOffset = $timeZoneInfo.BaseUtcOffset;
        [bool]$supportsDaylightSavingTime = $timeZoneInfo.SupportsDaylightSavingTime;
        if($supportsDaylightSavingTime -and $isDayLightSavingTime)
        {
            $LocalTime = $UTCTime.AddHours( - ($baseTimeZoneOffset.totalhours + 1));
        }
        else
        {
            $LocalTime = $UTCTime.AddHours( - ($baseTimeZoneOffset.totalhours));
        }
      Return $localTime;
    }
    #endregion

    [string]$logMessage = "Starting GetTodaysLatestCHUserExportFile: " + $FileName;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    #region Set Variables
    [string]$fileNameFound = "File Not Found";
    [DateTime]$rightNow = Get-Date
    [string]$filterString = $baseExportFileName + " " + $rightNow.ToString("yyyyMMdd") + "*." + $baseExportFileExtension
    #endregion

    Try
    {
        #Connect to PNP On-line
        Connect-PnPOnline -Url $SharePointSiteURL -Credentials $tenantCredentials

        #Get All Files from the document library - In batches of 500
        $ListItems = Get-PnPListItem -List $sharePointListName -PageSize 500 | Where-Object {$_["FileLeafRef"] -like $filterString}
        #$ListItems = Get-PnPListItem -List $sharePointListName -PageSize 500;
  
        #Loop through all documents
        $DocumentsData=@()
        ForEach($Item in $ListItems)
        {
            #Collect Documents Data
            $DocumentsData += New-Object PSObject -Property @{
                                                                FileName = $Item.FieldValues['FileLeafRef']
                                                                FileURL = $Item.FieldValues['FileRef']
                                                                CreatedUTC = Convert-UTCtoLocal $Item.FieldValues['Created']
                                                                Created = $Item.FieldValues['Created']
                                                             }
        }
        if($DocumentsData.Count -gt 0)
        {
            if($DocumentsData.Count -eq 1)
            {
                $fileNameFound = $DocumentsData[0].FileName;
            }
            else
            {
               $fileNameFound = $DocumentsData[0].FileName;
               $fileDate = $DocumentsData[0].Created;
               foreach($fil in $DocumentsData)
               {
                   if($fil.Created -gt $fileDate)
                   {
                       $fileNameFound = $fil.FileName;
                       $fileDate = $fil.Created;
                   }
               }
            }
        }
        Disconnect-PnPOnline;
    }
    catch 
    {
        $fileNameFound = "Error";
        $logMessage = "Error: $($_.Exception.Message)";
        .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    }
    $logMessage = "Finished GetTodaysLatestCHUserExportFile: " + $FileName;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    return $fileNameFound;
}

