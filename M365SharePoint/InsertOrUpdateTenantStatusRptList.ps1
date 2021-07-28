#
# InsertOrUpdateTenantStatusRptList.ps1
#
<# Header Information **********************************************************
Created By: Mike John
Created Date: 01/14/2021
Summary:
    Insert or Update Automated Processes Status Report List Item in SharePoint
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 02/20/2021
    Reason Updated: Added tenantDomain parameter; Added validation for 2 parameters
Updated By: Mike John
UpdatedDate: 01/14/2021
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2)][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3)][string]$ProcessName
     ,[Parameter(Mandatory=$True,Position=4)][string]$ProcessCategory
     ,[Parameter(Mandatory=$True,Position=5)][DateTime]$StartDate
     ,[Parameter(Mandatory=$True,Position=6)][DateTime]$StopDate
     ,[Parameter(Mandatory=$True,Position=7)][ValidateSet("Started","Successful","Failed")][string]$ProcessStatus
     ,[Parameter(Mandatory=$True,Position=8)][ValidateSet("In-progress","Completed","Error")][string]$ProcessProgress
     ,[Parameter(Mandatory=$True,Position=9)][string]$masterLogFilePathAndName
)
begin {}
process {
    # start here
    [string]$logMessage = "Starting InsertOrUpdateTenantStatusRptList.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    [string]$SharePointSiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam-M365Data"

    Connect-PnPOnline -Url $SharePointSiteURL -Credential $tenantCredentials;
    [string]$listName = "APStatuses";

    $listItem = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ProcessName</Value></Eq></Where></Query></View>"
    if($null -ne $listItem)
    {
        $id = $listItem["ID"];
        # update list here
        Set-PnPListItem -List $listName -ID $id -Values @{"StartDate"=$StartDate; "StopDate"=$StopDate; "ProcessStatus"=$ProcessStatus; "ProcessProgress"=$ProcessProgress}  | out-null;
    }
    else
    {
        # list item not found; so add it
        [datetime]$signalDate = Get-date("1/1/1900 00:01")
        if($signalDate -eq $StopDate)
        {
            # insert everything except stop date
            Add-PnPListItem -List $listName -Values @{"Title"=$ProcessName; "ProcessCategory"=$ProcessCategory; "StartDate"=$StartDate; "ProcessStatus"=$ProcessStatus; "ProcessProgress"=$ProcessProgress}  | out-null;
        }
        else
        {
            # insert everything
            Add-PnPListItem -List $listName -Values @{"Title"=$ProcessName; "ProcessCategory"=$ProcessCategory; "StartDate"=$StartDate;  "StopDate"=$StopDate; "ProcessStatus"=$ProcessStatus; "ProcessProgress"=$ProcessProgress}  | out-null;
        }
    }
    $logMessage = "Finished InsertOrUpdateTenantStatusRptList.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    Disconnect-PnPOnline -ErrorAction SilentlyContinue;
}
