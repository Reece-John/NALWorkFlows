#
# TestUpdateStatusRptList.ps1
#

    #region function definitions
    function CompareListItems($listItemx, $listItemy, [string]$logFilePathAndName)
    {
        [bool]$passedComparison = $true;
        [string]$logMessage = "";
        if($null -eq $listItemx)
        {
            $passedComparison = $false;
        }
        else
        {
            if($null -eq $listItemy)
            {
                $passedComparison = $false;
            }
            else
            {
                $tx = $listItemx["Title"];
                $ty = $listItemy["Title"];
                if($listItemx["Title"] -ne $listItemy["Title"])
                {
                    $logMessage = "Title !=";
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
                if($listItemx["ProcessCategory"] -ne $listItemy["ProcessCategory"])
                {
                    $logMessage = "ProcessCategory !=";
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
                if($listItemx["StartDate"] -ne $listItemy["StartDate"]) 
                {
                    $logMessage = "StartDate !=";  
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
                if($listItemx["StopDate"] -ne $listItemy["StopDate"]) 
                {
                    $logMessage = "StopDate !=";  
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
                if($listItemx["ProcessStatus"] -ne $listItemy["ProcessStatus"]) 
                {
                    $logMessage = "ProcessStatus !=";  
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
                if($listItemx["ProcessProgress"] -ne $listItemy["ProcessProgress"]) 
                {
                    $logMessage = "ProcessProgress !=";  
                    .\LogManagement\WriteToLogFile.ps1 -logFile $logFilePathAndName -message $logMessage;
                    $passedComparison = $false;
                }
            }
        }
        return $passedComparison;
    }
    #endregion

#starts Here
clear-host;

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
cd $startLoc;

# create the log file name
$dateRightNow = Get-Date;
[string]$myMasterLogFilePathAndName = 'c:\logs\TestUpdateStatusRptList_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

#Get Credentials to connect
$myCredentials = .\Common\ReturnCredentials.ps1

[bool]$isValidCompare = $true;
[string]$listName = "APStatuses";
[string]$ProcessName = "DailyRosterCreateTest";
[string]$ProcessCategory = "CHWorkFlow";
[DateTime]$StartDate = Get-date("1/1/9999 12:01");
[DateTime]$StopDate1 = Get-date("1/1/1900 00:01");
[DateTime]$StopDate2 = Get-date("1/2/9999 12:02");
[string]$ProcessStatus1 = "Failed";
[string]$ProcessStatus2 = "Successful";
[string]$ProcessProgress1 = "In-progress";
[string]$ProcessProgress2 = "Completed";

# get list item delete if there
$listItem0 = .\M365SharePoint\GetStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                            -ProcessName $ProcessName `
                                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;

if($null -ne $listItem0)
{
    $XX = 1;
    .\M365SharePoint\DeleteStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                      -ListItem $listItem0 `
                                                      -masterLogFilePathAndName $myMasterLogFilePathAndName;
}
$listItem1 = .\M365SharePoint\GetStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                            -ProcessName $ProcessName `
                                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;
if($null -eq $listItem1)
{
    # Create new item with no stop date
    $listItem2 = .\M365SharePoint\InsertOrUpdateStatusRptList.ps1 -tenantCredentials $myCredentials `
                                                                       -ProcessName $ProcessName `
                                                                       -ProcessCategory $ProcessCategory `
                                                                       -StartDate $StartDate `
                                                                       -StopDate $StopDate1 `
                                                                       -ProcessStatus $ProcessStatus1 `
                                                                       -ProcessProgress $ProcessProgress1 `
                                                                       -masterLogFilePathAndName $myMasterLogFilePathAndName;

    # get item and compare
    $listItem3 = .\M365SharePoint\GetStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                                -ProcessName $ProcessName `
                                                                -masterLogFilePathAndName $myMasterLogFilePathAndName;
    if($null -ne $listItem3)
    {
        $isValidCompare = CompareListItems $listItem2 $listItem3 $myMasterLogFilePathAndName;
        if($isValidCompare)
        {
            # Update item with stop date
            $listItem4 = .\M365SharePoint\InsertOrUpdateStatusRptList.ps1 -tenantCredentials $myCredentials `
                                                                               -ProcessName $ProcessName `
                                                                               -ProcessCategory $ProcessCategory `
                                                                               -StartDate $StartDate `
                                                                               -StopDate $StopDate2 `
                                                                               -ProcessStatus $ProcessStatus2 `
                                                                               -ProcessProgress $ProcessProgress2 `
                                                                               -masterLogFilePathAndName $myMasterLogFilePathAndName;

            # get item and compare
            $listItem5 = .\M365SharePoint\GetStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                                        -ProcessName $ProcessName `
                                                                        -masterLogFilePathAndName $myMasterLogFilePathAndName;
            $isValidCompare = CompareListItems $listItem4 $listItem5 $myMasterLogFilePathAndName;
            if($isValidCompare)
            {
                $xx = 1;
                # delete item
                .\M365SharePoint\DeleteStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                                  -ListItem $listItem4 `
                                                                  -masterLogFilePathAndName $myMasterLogFilePathAndName;
                # confirm deletion
                $listItem6 = .\M365SharePoint\GetStatusRptListItem.ps1 -tenantCredentials $myCredentials `
                                                                            -ProcessName $ProcessName `
                                                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;
                if($null -ne $listItem6)
                {
                    Write-Host("Failed to delete listItem6");
                }
            }
            else
            {
                Write-Host("Failed CompareListItems: listItem4 & listItem5");
            }
        }
        else
        {
            Write-Host("Failed CompareListItems: listItem4 & listItem5");
        }
    }
    else
    {
        Write-Host("listItem3 is null;");
    }
}
else
{
    Write-Host("Failed to delete listItem0");
}