#
# GetStatusRptListItem.ps1
#
<# Header Information **********************************************************
Created By: Mike John
Created Date: 12/09/2020
Summary:
    Get Automated Processes Status Report List Item from SharePoint
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 12/09/2020
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$ProcessName
     ,[Parameter(Mandatory=$True,Position=2)][string]$masterLogFilePathAndName
)
begin {}
process {
    # start here
    [string]$logMessage = "Starting GetStatusRptListItem.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    [string]$SharePointSiteURL = "https://algeorgetownarea.sharepoint.com/sites/TechnologyTeam"

    Connect-PnPOnline -Url $SharePointSiteURL -Credential $tenantCredentials;
    [string]$listName = "APStatuses";

    $listItems = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ProcessName</Value></Eq></Where></Query></View>"
    if($null -ne $listItems)
    {
        if($listItems.Length -eq 1)
        {
            $listItem = $listItems
        }
        else
        {
            $listItem =$listItems[0];
            $logMessage = "Found wrong number of List Items: " +  $listItems.Length + ".  GetStatusRptListItem.ps1";
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        }
    }
    else
    {
        $listItem = $null;
    }
    $logMessage = "Finished GetStatusRptListItem.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    return $listItem;
}