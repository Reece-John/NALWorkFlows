#
# ManageTeamSites.ps1
#
#           Author: Mike John
#     Date Created: 01/02/2020
#
# Last date edited: 01/02/2020
#   Last edited By: Mike John
# Last Edit Reason: Original
#
<#  Team Site Links

https://docs.microsoft.com/en-us/powershell/module/teams/new-teamchannel?view=teams-ps

https://docs.microsoft.com/en-us/powershell/module/teams/?view=teams-ps

https://laurakokkarinen.com/category/microsoft-teams/

https://laurakokkarinen.com/my-most-used-powershell-scripts-for-managing-sharepoint-online/

# from Flow
https://laurakokkarinen.com/provisioning-teams-with-a-site-design-flow-and-microsoft-graph/

#>

<#
Set-TeamPicture
   -GroupId <String>
   -ImagePath <String>
   [<CommonParameters>]
#>

<#
  Add-TeamUser -GroupId 31f1ff6c-d48c-4f8a-b2e1-abca7fd399df -User dmx@example.com

   -Role  ("Member" or "Owner")

   Add-TeamUser
   -GroupId <String>
   -User <String>
   [-Role <String>]
   [<CommonParameters>]
#>

function CreateOrUpdateTeam([PSObject]$tenantDefaultsObj, [PSObject]$objTeam)
{

$displayName                       = $objTeam.DisplayName;
$description                       = $objTeam.Description;
$mailNickName                      = $objTeam.MailNickName;
$visibility                        = $objTeam.Visibility;
$templateName                      = $objTeam.Template;
$owner                             = $objTeam.Owner;
$allowGiphy                        = $objTeam.AllowGiphy;
$giphyContentRating                = $objTeam.GiphyContentRating;
$allowStickersAndMemes             = $objTeam.AllowStickersAndMemes;
$allowCustomMemes                   = $objTeam.AllowCustomMemes;
$allowGuestCreateUpdateChannels    = $objTeam.AllowGuestCreateUpdateChannels;
$allowGuestDeleteChannels          = $objTeam.AllowGuestDeleteChannels;
$allowCreateUpdateChannels         = $objTeam.AllowCreateUpdateChannels;
$allowDeleteChannels               = $objTeam.AllowDeleteChannels;
$allowAddRemoveApps                = $objTeam.AllowAddRemoveApps;
$allowCreateUpdateRemoveTabs       = $objTeam.AllowCreateUpdateRemoveTabs;
$allowCreateUpdateRemoveConnectors = $objTeam.AllowCreateUpdateRemoveConnectors;
$allowUserEditMessages             = $objTeam.AllowUserEditMessages;
$allowUserDeleteMessages           = $objTeam.AllowUserDeleteMessages;
$allowOwnerDeleteMessages          = $objTeam.AllowOwnerDeleteMessages;
$allowTeamMentions                 = $objTeam.AllowTeamMentions;
$allowChannelMentions              = $objTeam.AllowChannelMentions;
$showInTeamsSearchAndSuggestions   = $objTeam.ShowInTeamsSearchAndSuggestions;

$curTeamGroup = New-Team `
                  -DisplayName $displayName `
                  -Description $description `
                  -MailNickName $mailNickName `
                  -Visibility $visibility `
                  -Template $templateName `
                  -Owner $owner `
                  -AllowGiphy $allowGiphy `
                  -GiphyContentRating $giphyContentRating `
                  -AllowStickersAndMemes $allowStickersAndMemes `
                  -AllowCustomMemes $allowCustomMemes `
                  -AllowGuestCreateUpdateChannels $allowGuestCreateUpdateChannels `
                  -AllowGuestDeleteChannels $allowGuestDeleteChannels `
                  -AllowCreateUpdateChannels $allowCreateUpdateChannels `
                  -AllowDeleteChannels $allowDeleteChannels `
                  -AllowAddRemoveApps $allowAddRemoveApps `
                  -AllowCreateUpdateRemoveTabs $allowCreateUpdateRemoveTabs `
                  -AllowCreateUpdateRemoveConnectors $allowCreateUpdateRemoveConnectors `
                  -AllowUserEditMessages $allowUserEditMessages `
                  -AllowUserDeleteMessages $allowUserDeleteMessages `
                  -AllowOwnerDeleteMessages $allowOwnerDeleteMessages `
                  -AllowTeamMentions $allowTeamMentions `
                  -AllowChannelMentions $allowChannelMentions `
                  -ShowInTeamsSearchAndSuggestions $showInTeamsSearchAndSuggestions;
    return $curTeamGroup;
}

function UpdateTeam([PSObject]$tenantDefaultsObj, [PSObject]$newTeamObj)
{

}

function IsTeamDifferent([PSObject]$tenantDefaultsObj, [PSObject]$tobj, [PSObject]$teamObj)
{

}

function ReportDeletes([PSObject]$teamObjs)
{

}

function LoadTeamObjectsFromExcelFile([string]$filePathName)
{
    $teamsExcelData = Import-Excel -Path $filePathName  -WorksheetName "M365Teams" -StartRow 1 -DataOnly;
    return $teamsExcelData;
}

function LoadTeamMemberObjectsFromExcelFile([string]$filePathName)
{
    $teamMembersExcelData = Import-Excel -Path $filePathName  -WorksheetName "M365TeamMembers" -StartRow 1 -DataOnly;
    return $teamMembersExcelData;
}


function LoadTenantDefaultsObjFromExcelFile([string]$filePathName)
{
    $m365DefaultsExcelData = Import-Excel -Path $filePathName  -WorksheetName "M365TenantDefaults" -StartRow 1 -DataOnly;
    return $m365DefaultsExcelData;
}

# starts here
Clear-Host;

$startLoc = [Environment]::GetEnvironmentVariable("PSStartup","Machine");
Set-Location $startLoc;

[string]$filePathName = "C:\PSScripts\ExcelDataFiles\ALGASchema.xlsx";

#load team objects from file
[PSObject]$objTeams = LoadTeamObjectsFromExcelFile $filePathName

#load team objects from file
[PSObject]$objTeamMembers = LoadTeamMemberObjectsFromExcelFile $filePathName
foreach($tmObj in $objTeamMembers)
{
    Write-Host($tmObj)
}

#[PSObject]$tenantDefaultsObj1 = LoadTenantDefaultsObjFromExcelFile $filePathName


#connect to Tenant
[String]$username = "PSAdmin";
[system.Management.Automation.PSCredential]$tmpCreds = .\Common\GetRemoteCredentials.ps1 $username;
[String]$clearTextPassword = $tmpCreds.GetNetworkCredential().password;
$password = $clearTextPassword | ConvertTo-SecureString -asPlainText -Force;
$userName = "mjohn@ALGeorgetownArea.org";
[System.Management.Automation.PSCredential]$cred = New-Object System.Management.Automation.PSCredential($username,$password);

#connect to Microsoft Teams Service
Connect-MicrosoftTeams -Credential $cred;  

foreach($objTeam in $objTeams)
{
    $teamGroup = CreateOrUpdateTeam $tenantDefaultsObj $objTeam;
    $teamMembers = GetTeamMembers $teamGroup;
    foreach($teamMemberObj in $teamMembers)
    {
        if($teamMemberObj.MemberType -eq "Owner")
        {
            Add-TeamUser -GroupId $teamGroup.GroupId -Owner $teamMemberObj.ALGAEmail;
        }
        else
        {
            Add-TeamUser -GroupId $teamGroup.GroupId -User $teamMemberObj.ALGAEmail;
        }
    }
}