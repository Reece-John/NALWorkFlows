#
# CreateTenantRoleBasedEmails.ps1
#
<#
#           Author: Mike John
#     Date Created: 08/17/2020
#***************************************
# Last date edited: 03/05/2021
#   Last edited By: Mike John
# Last Edit Reason: Added Shared mailboxes and Automate mailboxes
#***************************************
# Last date edited: 08/17/2020
#   Last edited By: Mike John
# Last Edit Reason: Original with Individual mailboxes
#>
<#
    Prerequisites:
    Must be in the correct starting directory (This is a hidden requirement)
    Excel Schema file must be in the correct directory passed in as a parameter
    Etc..
#>
<# Role-based email user has the following items set in Microsoft 365 profile
        Example:
        First name: "ALGA" (Chapter acronym)
        Last name: "AccountsReceivable"
    Display name: "AccountsReceivable"
        Job title: "Role-based Email"
#>


[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2 )][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3)][string]$alChapterSchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4)][string]$tenantDefaultsPageName
     ,[Parameter(Mandatory=$True,Position=5)][int]$tenantDefaultsStartRow
     ,[Parameter(Mandatory=$True,Position=6)][string]$rolebasedEmailsPageName
     ,[Parameter(Mandatory=$True,Position=7)][int]$rolebasedEmailsStartRow
     ,[Parameter(Mandatory=$True,Position=8)][string]$masterLogFilePathAndNamePathAndName
     ,[Parameter(Mandatory=$True,Position=9)][bool]$justTestingOnly
)
begin {}
process {
    #region function definitions
    function LoadObjectsFromExcelFile([string]$filePathName, [string]$pageName, [int]$startRow)
    {
        $excelData = Import-Excel -Path $filePathName -WorksheetName $pageName -StartRow $startRow -DataOnly;
        return $excelData;
    }

    function CreateSharedRoleBasedEmail([PSObject]$tenantDefaultsObj, [PSObject]$uObj)
    {
        [string]$sharedMailAddress  = $uObj.RolebasedEmailAddress;
        [string]$firstName          = $tenantDefaultsObj.ChapterAcronym;
        [string]$lastName           = $uObj.LastName;
        [string]$rolebasedEMailName = $firstName + " " + $lastName;
        [string[]]$memberArray = $uObj.GroupMembers.Split(";");
        New-Mailbox -Shared -Name $rolebasedEMailName -PrimarySmtpAddress $sharedMailAddress;
        foreach($sharedMailBoxMember in $memberArray)
        {
            $sharedMailBoxMember.Trim();
            Add-MailboxPermission -Identity $sharedMailBoxName -User $sharedMailBoxMember -AccessRights FullAccess;
            Add-RecipientPermission -Identity $sharedMailBoxName -Trustee $sharedMailBoxMember -AccessRights SendAs -Confirm:$false;
        }
    }

    function UpdateIfSharedRoleBasedEmailProfileDifferent([PSObject]$tenantDefaultsObj, [PSObject]$schemaEmailUserObj, [PSObject]$uObj, [bool]$justTestingOnly)
    {
        <#
        [bool]$ProfileIsDifferent = $false;
        [string]$logMessage = "";

        $rolebasedEMailAddress = $schemaEmailUserObj.RolebasedEmailAddress;
        $rolebasedEMailAddress = $rolebasedEMailAddress.ToLower();

        #put updates here
        [string]$uobjDisplayName = $uObj.DisplayName;
        if($uobjDisplayName -eq $null)
        {
            $uobjDisplayName = "...";
        }
        if($schemaEmailUserObj.DisplayName -ne $uobjDisplayName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " LastName ProfileIsDifferent: " + $schemaEmailUserObj.DisplayName + " - " + $uobjDisplayName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " changing DisplayName to: " + $schemaEmailUserObj.DisplayName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -DisplayName $schemaEmailUserObj.DisplayName
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " would change DisplayName to: " + $schemaEmailUserObj.DisplayName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
        }
            
        [string]$uobjFirstName = $uObj.FirstName;
        if($uobjFirstName -eq $null)
        {
            $uobjFirstName = "...";
        }
        if($tenantDefaultsObj.ChapterAcronym -ne $uobjFirstName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " FirstName ProfileIsDifferent: " + $tenantDefaultsObj.ChapterAcronym + " - " + $uobjFirstName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " changing FirstName to: " + $tenantDefaultsObj.ChapterAcronym;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -FirstName $tenantDefaultsObj.ChapterAcronym
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " would change FirstName to: " + $tenantDefaultsObj.ChapterAcronym;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
        }

        [string]$uobjLastName = $uObj.LastName;
        if($uobjLastName -eq $null)
        {
            $uobjLastName = "...";
        }
        if($schemaEmailUserObj.LastName -ne $uobjLastName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " FirstName ProfileIsDifferent: " + $schemaEmailUserObj.FirstName + " - " + $uobjLastName;
            if(!$justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " changing LastName to: " + $schemaEmailUserObj.LastName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -LastName $schemaEmailUserObj.LastName
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " would change LastName to: " + $schemaEmailUserObj.LastName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
        }

        [string]$uobjTitle = $uObj.Title;
        if($uobjTitle -eq $null)
        {
            $uobjTitle = "...";
        }
        if($tenantDefaultsObj.RoleBasedEmailTitle -ne $uobjTitle)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " Title ProfileIsDifferent: " + $tenantDefaultsObj.RoleBasedEmailTitle + " - " + $uobjTitle;
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " changing Title to: " + $tenantDefaultsObj.RoleBasedEmailTitle;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -Title $tenantDefaultsObj.RoleBasedEmailTitle
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " would change Title to: " + $tenantDefaultsObj.RoleBasedEmailTitle;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
        }
        return $ProfileIsDifferent;
        #>
    }

    function CreateRoleBasedEmail([PSObject]$tenantDefaultsObj, [PSObject]$uObj)
    {
        $userPrincipalName       = $uObj.RolebasedEmailAddress;
        $displayName             = $uObj.DisplayName;
        $firstName               = $tenantDefaultsObj.ChapterAcronym;
        $lastName                = $uObj.LastName;
        $department              = $uObj.Department;
        $title                   = $tenantDefaultsObj.RoleBasedEmailTitle;
        $usageLocation           = $tenantDefaultsObj.UsageLocation;
        $userPassword            = $tenantDefaultsObj.PasswordDefault;
        $forceChangePassword     = $tenantDefaultsObj.ForceChangePassword;
        $licenseAssignment       = $tenantDefaultsObj.LicenseAssignment;

        New-MsolUser `
            -UserPrincipalName $userPrincipalName `
            -BlockCredential $false `
            -DisplayName $displayName `
            -FirstName $firstName `
            -LastName $lastName `
            -Department $department `
            -PasswordNeverExpires $true `
            -StrongPasswordRequired $true `
            -Title $title `
            -UsageLocation $usageLocation `
            -Password $userPassword `
            -ForceChangePassword $forceChangePassword `
            -LicenseAssignment $licenseAssignment;

    }

    function UpdateIfRoleBasedEmailUserProfileDifferent([PSObject]$tenantDefaultsObj, [PSObject]$schemaEmailUserObj, [PSObject]$uObj, [bool]$justTestingOnly)
    {
        [bool]$ProfileIsDifferent = $false;
        [string]$logMessage = "";

        $rolebasedEMailAddress = $schemaEmailUserObj.RolebasedEmailAddress;
        $rolebasedEMailAddress = $rolebasedEMailAddress.ToLower();


        #put updates here
        [string]$uobjDisplayName = $uObj.DisplayName;
        if($uobjDisplayName -eq $null)
        {
            $uobjDisplayName = "...";
        }
        if($schemaEmailUserObj.DisplayName -ne $uobjDisplayName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " LastName ProfileIsDifferent: " + $schemaEmailUserObj.DisplayName + " - " + $uobjDisplayName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if($justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " would change DisplayName to: " + $schemaEmailUserObj.DisplayName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " changing DisplayName to: " + $schemaEmailUserObj.DisplayName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -DisplayName $schemaEmailUserObj.DisplayName
            }
        }
            
        [string]$uobjFirstName = $uObj.FirstName;
        if($uobjFirstName -eq $null)
        {
            $uobjFirstName = "...";
        }
        if($tenantDefaultsObj.ChapterAcronym -ne $uobjFirstName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " FirstName ProfileIsDifferent: " + $tenantDefaultsObj.ChapterAcronym + " - " + $uobjFirstName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if($justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " would change FirstName to: " + $tenantDefaultsObj.ChapterAcronym;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " changing FirstName to: " + $tenantDefaultsObj.ChapterAcronym;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -FirstName $tenantDefaultsObj.ChapterAcronym
            }
        }

        [string]$uobjLastName = $uObj.LastName;
        if($uobjLastName -eq $null)
        {
            $uobjLastName = "...";
        }
        if($schemaEmailUserObj.LastName -ne $uobjLastName)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " FirstName ProfileIsDifferent: " + $schemaEmailUserObj.FirstName + " - " + $uobjLastName;
            if($justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " would change LastName to: " + $schemaEmailUserObj.LastName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " changing LastName to: " + $schemaEmailUserObj.LastName;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -LastName $schemaEmailUserObj.LastName
            }
        }

        [string]$uobjTitle = $uObj.Title;
        if($uobjTitle -eq $null)
        {
            $uobjTitle = "...";
        }
        if($tenantDefaultsObj.RoleBasedEmailTitle -ne $uobjTitle)
        {
            $ProfileIsDifferent = $true;
            $logMessage = $rolebasedEMailAddress + " Title ProfileIsDifferent: " + $tenantDefaultsObj.RoleBasedEmailTitle + " - " + $uobjTitle;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if($justTestingOnly)
            {
                $logMessage = $rolebasedEMailAddress + " would change Title to: " + $tenantDefaultsObj.RoleBasedEmailTitle;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            }
            else
            {
                $logMessage = $rolebasedEMailAddress + " changing Title to: " + $tenantDefaultsObj.RoleBasedEmailTitle;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $rolebasedEMailAddress -Title $tenantDefaultsObj.RoleBasedEmailTitle
            }
        }

        return $ProfileIsDifferent;
    }
    #endregion

    # script starts here
    # used for debugging
    $xx = 1;

    [string]$thisScriptName = $MyInvocation.MyCommand.Name
    [string]$logMessage = "Starting " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    #region load object arrays from Excel sheets
    $algaTenantDefaultsObj = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $tenantDefaultsPageName $tenantDefaultsStartRow;
    $rolebasedEmailObjs = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $rolebasedEmailsPageName $rolebasedEmailsStartRow;
    #endregion load object arrays from Excel sheets

    $clearTextPassword = $tenantCredentials.GetNetworkCredential().password;
    $userName = $tenantCredentials.GetNetworkCredential().UserName;

    #Connect-MsolService -Credential $tenantCredentials;
    Connect-ExchangeOnline -Credential $tenantCredentials -ShowBanner:$false -ShowProgress:$false;


    $x = 1;


    #loop through role-based email objects Creating a user by that name as necessary 
    foreach($rolebasedEMailObj in $rolebasedEmailObjs)
    {
        $rolebasedEMailAddress = $rolebasedEMailObj.RolebasedEmailAddress.ToLower();
        $emailType = $rolebasedEMailObj.TypeEmail;
        if($emailType -eq "Shared")
        {
            $sharedMailBox = $null;
            $sharedMailBox = Get-Mailbox -Identity $rolebasedEMailAddress; # -ErrorAction SilentlyContinue;
            if($null -eq $sharedMailBox)
            {
                if($justTestingOnly)
                {
                    $logMessage = "Would create new shared Role-based email: " + $rolebasedEMailAddress;
                    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                }
                else
                {
                    # Create SharedRole-based Email box
                    $logMessage = "create new shared Role-based email: " + $rolebasedEMailAddress;
                    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    CreateSharedRoleBasedEmail $algaTenantDefaultsObj $rolebasedEMailObj;
                }
            }
            else
            {
                $rolebasedEMailName = $algaTenantDefaultsObj + " " +$rolebasedEMailObj.LastName;
                $logMessage = "Updating shared Role-based Email if Needed: " + $rolebasedEMailAddress;
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                $xx = 5;
                UpdateIfSharedRoleBasedEmailProfileDifferent $algaTenantDefaultsObj $rolebasedEMailObj $roleEmailUserObj $justTestingOnly;
            }
        }
        else
        {
            if($emailType -eq "Individual")
            {
                $roleEmailUserObj = $null;
                #$roleEmailUserObj = Get-MsolUser -UserPrincipalName $rolebasedEMailAddress -ErrorAction SilentlyContinue;
                if($null -eq $roleEmailUserObj)
                {
                    if($justTestingOnly)
                    {
                        $logMessage = "Would create Individual Role-based Email: " + $rolebasedEMailAddress;
                        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    }
                    else
                    {
                        # need to make new Individual Role-based Email user here
                        $logMessage = "create new Individual role-based email: " + $rolebasedEMailAddress;
                        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        CreateRoleBasedEmail $algaTenantDefaultsObj $rolebasedEMailObj;
                    }
                }
                else
                {
                    # need to make new Individual Role-based Email user here
                    $logMessage = "Update new Individual role-based email: " + $rolebasedEMailAddress;
                    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    # update Role-based Email user as needed
                    #UpdateIfRoleBasedEmailUserProfileDifferent $algaTenantDefaultsObj $rolebasedEMailObj $roleEmailUserObj $justTestingOnly;
                }
            }
            else
            {
                if($emailType -eq "Automation")
                {
                    $roleEmailUserObj = $null;
                    #$roleEmailUserObj = Get-MsolUser -UserPrincipalName $rolebasedEMailAddress -ErrorAction SilentlyContinue;
                    if($null -eq $roleEmailUserObj)
                    {
                        if($justTestingOnly)
                        {
                            $logMessage = "Would create Automation Email: " + $rolebasedEMailAddress;
                            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        }
                        else
                        {
                            # need to make Role-based Email user here
                            $logMessage = "Create new Automation email: " + $rolebasedEMailAddress;
                            .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                            CreateRoleBasedEmail $algaTenantDefaultsObj $rolebasedEMailObj;
                        }
                    }
                    else
                    {
                        # need to make Role-based Email user here
                        $logMessage = "Update Automation email if Needed: " + $rolebasedEMailAddress;
                        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        # update Role-based Email user as needed
                        #UpdateIfRoleBasedEmailUserProfileDifferent $algaTenantDefaultsObj $rolebasedEMailObj $roleEmailUserObj $justTestingOnly;
                    }
                }
                else
                {
                    #"Error not a valid email type"
                    $xx = 99;
                }
            }
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false;
    Disconnect-MsolService;
    $logMessage = "Finished " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
end{}
