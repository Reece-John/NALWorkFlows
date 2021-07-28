#
# MakeSharePointListReadonly.ps1
#
<#

https://www.sharepointdiary.com/2018/06/limited-access-user-permission-lockdown-mode-feature.html

#>
#Parameter
$SiteURL = "https://Crescent.sharepoint.com/sites/PMO"
$ListName= "Projects"
  
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin
  
#Get the Web and List
$Web = Get-PnPWeb
$List = Get-PnPList -Identity $ListName -Includes HasUniqueRoleAssignments, RoleAssignments
 
#Break Permissions of the List
If ($List.HasUniqueRoleAssignments -eq $False)
{
    Set-PnPList -Identity $ListName -BreakRoleInheritance -CopyRoleAssignments
}
  
#Get Read Permission Level
$ReadPermission = Get-PnPRoleDefinition -Identity "Read"
 
#Grant "Read" permissions, if its not granted already
$List.RoleAssignments | ForEach-Object {
    #Get the user or group of the assignment - Handle error for orphans
    $Member = Get-PnPProperty -ClientObject $_ -Property Member -ErrorAction SilentlyContinue
 
    If($Member.IsHiddenInUI -eq $False)
    {
        Get-PnPProperty -ClientObject $_ -Property RoleDefinitionBindings | Out-Null
  
        #Check if the current assignment has any permission other than Read or related
        $PermissionsToReplace = $_.RoleDefinitionBindings | Where {$_.Hidden -eq $False -And $_.Name -Notin ("Read", "Restricted Read", "Restricted Interfaces for Translation")}
         
        #Grant "Read" permissions, if its not granted already
        If($PermissionsToReplace -ne $Null)
        {
            $_.RoleDefinitionBindings.Add($ReadPermission)
            $_.Update()
            Invoke-PnPQuery
            Write-host "Added 'Read' Permissions to '$($Member.Title)'" -ForegroundColor Cyan
        }
    }
}
#Reload List permissions
$List = Get-PnPList -Identity $ListName -Includes RoleAssignments
 
#Remove All permissions other than Read or Similar
$List.RoleAssignments | ForEach-Object {
    #Get the user or group of the assignment - Handle error for orphans
    $Member = Get-PnPProperty -ClientObject $_ -Property Member #-ErrorAction SilentlyContinue | Out-Null   
    If($Member.IsHiddenInUI -eq $False)
    {
        Get-PnPProperty -ClientObject $_ -Property RoleDefinitionBindings | Out-Null
  
        $PermissionsToRemove = $_.RoleDefinitionBindings | Where {$_.Hidden -eq $False -And $_.Name -Notin ("Read", "Restricted Read", "Restricted Interfaces for Translation")}
        If($PermissionsToRemove -ne $null)
        {
            ForEach($RoleDefBinding in $PermissionsToRemove)
            {
                $_.RoleDefinitionBindings.Remove($RoleDefBinding)
                Write-host "Removed '$($RoleDefBinding.Name)' Permissions from '$($Member.Title)'" -ForegroundColor Yellow   
            }
            $_.Update()
            Invoke-PnPQuery
        }
    }
}
Write-host "List is set to Read-Only Successfully!" -f Green


#Read more: https://www.sharepointdiary.com/2020/03/sharepoint-online-make-list-read-only-using-powershell.html#ixzz6ZqFS70s7