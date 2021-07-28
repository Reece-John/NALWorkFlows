#
# DeleteStatusRptListItem.ps1
#
<# Header Information **********************************************************
Created By: Mike John
Created Date: 12/09/2020
Summary:
    Delete Automated Processes Status Report List Item from SharePoint
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 12/09/2020
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][Microsoft.SharePoint.Client.ListItem]$ListItem
     ,[Parameter(Mandatory=$True,Position=2)][string]$masterLogFilePathAndName
)
begin {}
process {
    #endregion

    # start here
    [string]$logMessage = "Starting DeleteStatusListItem.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    [string]$SharePointSiteURL = "https://algeorgetownarea.sharepoint.com/sites/TechnologyTeam"

    Connect-PnPOnline -Url $SharePointSiteURL -Credential $tenantCredentials;
    [string]$listName = "APStatuses";

    $listItem = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ProcessName</Value></Eq></Where></Query></View>"
    Remove-PnPListItem -List $listName -Identity $ListItem -Force -Recycle
    $logMessage = "Finishing DeleteStatusListItem.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}