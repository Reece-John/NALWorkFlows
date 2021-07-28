#
# ManageM365TenantUsers.ps1
#
<#
# Manages the alChapter users in the Microsoft provided tenant
#
# The master data is the spreadsheet in the Technology Committee Teams Site
#   A scheduled PowerShell script runs
#     Copies the spreadsheet from the Teams site to the local machine
#     Spreadsheet is the source data used to maintain users and groups
#     PowerShell object arrays are created from imports from the spreadsheet
#     Users are maintained from the object arrays
#

#  Arrays are iterated through:
#  If the user in the data file is not there it is added
#  If the user in the data file is there, the user record is updated if and only if there are any differences
#  User is added as a guest if there are no licenses assigned
#           Author: Mike John
#     Date Created: 02/20/2021
#
# Last date edited: 02/20/2021
#   Last edited By: Mike John
# Last Edit Reason: Original
#>

<# References: 02/02/2021

    https://docs.microsoft.com/en-us/microsoft-365/enterprise/powershell/manage-office-365-with-office-365-powershell
    https://docs.microsoft.com/en-us/microsoft-365/enterprise/create-user-accounts-with-microsoft-365-powershell?view=o365-worldwide
    https://www.sharepointdiary.com/2018/05/add-members-to-office-365-group-using-powershell.html
    https://docs.microsoft.com/en-us/powershell/module/azuread/new-azureaduser?view=azureadps-2.0
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0 )][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1 )][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2 )][PSObject]$tenantObj
     ,[Parameter(Mandatory=$True,Position=3 )][string]$alChapterSchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4)][string]$masterLogFilePathAndName
     ,[Parameter(Mandatory=$True,Position=5)][bool]$justTestingOnly
)
begin {}
process {
#region function definitions
    function CreateAzureADUserGraph([PSObject]$tenantDefaultsObj, [PSObject]$alChapterUserObj)
    {
        $params = @{
            AccountEnabled = $true
            DisplayName = "Mike Finnegan"
            PasswordProfile = $PasswordProfile
            UserPrincipalName = "FwF@Motor.onmicrosoft.com"
            MailNickName = "Finnegan"
        }
        New-AzureADUser @params;
    }

    function CreateGuest([PSObject]$tenantDefaultsObj, [PSObject]$alChapterUserObj)
    {
        <#
        $personalEmail            = $alChapterUserObj.PersonalEmail;
        $city                    = $alChapterUserObj.City;
        $country                 = $alChapterUserObj.Country;
        $department              = $alChapterUserObj.Department;
        $displayName             = $alChapterUserObj.DisplayName;
        $firstName               = $alChapterUserObj.FirstName;
        $lastName                = $alChapterUserObj.LastName;
        $mobilePhone             = $alChapterUserObj.MobilePhone;
        $office                  = $tenantDefaultsObj.Office;
        $passwordNeverExpires    = $tenantDefaultsObj.PasswordNeverExpires;
        if($alChapterUserObj.HomePhone = "None on File")
        {
            $homePhone = $null
        }
        else
        {
            $homePhone             = $personalEmail;
        }
        $postalCode              = $alChapterUserObj.PostalCode;
        $state                   = $alChapterUserObj.State;
        $streetAddress           = $alChapterUserObj.StreetAddress;
        $strongPasswordRequired  = $tenantDefaultsObj.StrongPasswordRequired;
        $title                   = $alChapterUserObj.Title;
        $usageLocation           = $tenantDefaultsObj.UsageLocation;
        if($alChapterUserObj.PersonalEmail = "*None On File*")
        {
            $PersonalEmail = $null
        }
        else
        {
            $PersonalEmail = $alChapterUserObj.PersonalEmail;
        }
        $userType                = $alChapterUserObj.UserType;
        $userPassword            = $tenantDefaultsObj.PasswordDefault;
        $forceChangePassword     = $tenantDefaultsObj.ForceChangePassword;
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $userPassword;
        $PasswordProfile.ForceChangePasswordNextLogin = $forceChangePassword;
        #>

        <#
        New-MsolUser `
            -UserPrincipalName $chapterEmail `
            -BlockCredential $blockCredential `
            -City $city `
            -Country $country `
            -Department $department `
            -DisplayName $displayName `
            -FirstName $firstName `
            -LastName $lastName `
            -MobilePhone $mobilePhone `
            -Office $office `
            -PasswordNeverExpires $true `
            -PhoneNumber $homePhone `
            -PostalCode $postalCode `
            -State $state `
            -StreetAddress $streetAddress `
            -StrongPasswordRequired $true `
            -Title $title `
            -UsageLocation $usageLocation `
            -AlternateEmailAddresses $PersonalEmail `
            -UserType $userType `
            -Password $userPassword `
            -ForceChangePassword $true;
        #>
        <#


        New-AzureADUser `
            -AccountEnabled $true `
            -City $city `
            -CompanyName $tenantAbbreviation `
            -Country $country `
            -Department $department `
            -DisplayName $displayName `
            -FirstName $firstName `
            -MobilePhone $mobilePhone `
            -OtherMails $personalEmail `
            -PasswordPolicies "DisablePasswordExpiration" `
            -PasswordProfile $PasswordProfile `
           [-PhysicalDeliveryOfficeName <String>]
            -PostalCode $postalCode `
            -ShowInAddressList $true `
            -State $state `
            -StreetAddress $streetAddress `
            -Surname $lastName `
            -TelephoneNumber $homePhone `
            -UsageLocation $usageLocation `
           [-UserPrincipalName <String>]
            -UserType $userType `
        #>
    }

    function CreateUser([PSObject]$tenantDefaultsObj, [PSObject]$alChapterUserObj)
    {
        $chapterEmail            = $alChapterUserObj.ChapterEmail;
        $blockCredential         = $alChapterUserObj.BlockCredential;
        $city                    = $alChapterUserObj.City;
        $country                 = $alChapterUserObj.Country;
        $department              = $alChapterUserObj.Department;
        $displayName             = $alChapterUserObj.DisplayName;
        $firstName               = $alChapterUserObj.FirstName;
        $lastName                = $alChapterUserObj.LastName;
        $mobilePhone             = $alChapterUserObj.MobilePhone;
        $office                  = $tenantDefaultsObj.Office;
        if($alChapterUserObj.HomePhone = "None on File")
        {
            $homePhone = $null
        }
        else
        {
            $homePhone             = $alChapterUserObj.HomePhone;
        }
        $postalCode              = $alChapterUserObj.PostalCode;
        $state                   = $alChapterUserObj.State;
        $streetAddress           = $alChapterUserObj.StreetAddress;
        $title                   = $alChapterUserObj.Title;
        $usageLocation           = $tenantDefaultsObj.UsageLocation;
        if($alChapterUserObj.PersonalEmail = "*NO EMAIL*")
        {
            $PersonalEmail = $null
        }
        else
        {
            $PersonalEmail = $alChapterUserObj.PersonalEmail;
        }
        $userType                = $alChapterUserObj.UserType;
        $userPassword            = $tenantDefaultsObj.PasswordDefault;
        $licenseAssignment       = $tenantDefaultsObj.LicenseAssignment;

        New-MsolUser `
            -UserPrincipalName $chapterEmail `
            -BlockCredential $blockCredential `
            -City $city `
            -Country $country `
            -Department $department `
            -DisplayName $displayName `
            -FirstName $firstName `
            -LastName $lastName `
            -MobilePhone $mobilePhone `
            -Office $office `
            -PasswordNeverExpires $true `
            -PhoneNumber $homePhone `
            -PostalCode $postalCode `
            -State $state `
            -StreetAddress $streetAddress `
            -StrongPasswordRequired $true `
            -Title $title `
            -UsageLocation $usageLocation `
            -AlternateEmailAddresses $PersonalEmail `
            -UserType $userType `
            -Password $userPassword `
            -LicenseAssignment $licenseAssignment `
            -ForceChangePassword $true;
        
        <#
        #now add any licenses
        #split the license string into an array
        [string[]]$usrLicenseAssignmentSkUs = $alChapterUserObj.LicenseAssignment.Split(";");
        foreach($usrLicenseAssignmentSKU in $usrLicenseAssignmentSkUs)
        {
            $ulaSKU = $usrLicenseAssignmentSKU.Trim().ToLower();
            Set-MsolUserLicense -UserPrincipalName $chapterEmail -AddLicenses $ulaSKU;
        }
        #>
    }

    function UpdateUserLicense([PSObject]$alChapterUpdateUserObj)
    {
        [string[]]$usrLicenseAssignmentSkUs = $alChapterUpdateUserObj.LicenseAssignment.Split(";");
        $m365UsrObj = Get-MsolUser -UserPrincipalName $alChapterUpdateUserObj.ChapterEmail;
        # now we have 2 arrays what it should be in one array and plus what it is in another array

        # loop making sure we add what should be there if it is not already there
        foreach($usrLicenseAssignmentSKU in $usrLicenseAssignmentSkUs)
        {
            [bool]$isNotAlreadyThere = $true;
            $ulaSKU = $usrLicenseAssignmentSKU.Trim().ToLower();
            foreach($m365LicenseAssignment in $m365UsrObj.Licenses)
            {
                if($ulaSKU -eq $m365LicenseAssignment.AccountSkuId.Trim().ToLower())
                {
                    $isNotAlreadyThere = $false;
                }
            }
            if($isNotAlreadyThere)
            {
                #not there so add it
                $logMessage  = "**Dummy***Adding license: " + $usrLicenseAssignmentSKU + " - to User: " + $uObj.ChapterEmail;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Write-Host($msgToLog);
                #Set-MsolUserLicense -UserPrincipalName $alChapterUpdateUserObj.ChapterEmail; -AddLicenses $ulaSKU
            }
        }

        # loop taking away what should not be there
        foreach($m365LicenseAssignment in $m365UsrObj.Licenses)
        {
            [bool]$shouldNotBeThere = $true;
            $m365LASKU = $m365LicenseAssignment.AccountSkuId.Trim().ToLower()
            foreach($usrLicenseAssignmentSKU in $usrLicenseAssignmentSkUs)
            {
                if($m365LASKU -eq $usrLicenseAssignmentSKU.Trim().ToLower())
                {
                    $shouldNotBeThere = $false;
                }
            }
            if($shouldNotBeThere)
            {
                #should not be there so remove it
                $logMessage  = "**Dummy***Removing license: " + $m365LASKU + " - from User: " + $uObj.ChapterEmail;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                #Set-MsolUserLicense -UserPrincipalName $alChapterUpdateUserObj.ChapterEmail; -RemoveLicenses  $m365LASKU
            }
        }
    }

    function isUserLicenseAssignmentDifferent([PSObject]$alChapterUpdateUserObj)
    {
        $licensesIsDifferent = $false;
        if($null -eq $alChapterUpdateUserObj.LicenseAssignment)
        {
            [string[]]$usrLicenseAssignmentSkUs = $alChapterUpdateUserObj.LicenseAssignment.Split(";");

            $m365UsrObj = Get-MsolUser -UserPrincipalName $alChapterUpdateUserObj.ChapterEmail;
            $m365LicenseAssignments = $m365UsrObj.Licenses;
            # now we have 2 arrays
            #  - what it should be assigned in the first array (SKUs)
            #  - what is assigned in the second array (License Assignments with a SKU property)

            # loop testing for what should be there
            [bool]$isNotAlreadyThere = $true;
            foreach($usrLicenseAssignmentSKU in $usrLicenseAssignmentSkUs)
            {
                foreach($m365LicenseAssignment in $m365LicenseAssignments)
                {
                    if($usrLicenseAssignmentSKU.Trim().ToLower() -eq $m365LicenseAssignment.AccountSkuId.ToLower())
                    {
                        $isNotAlreadyThere = $false;
                    }
                }
                if($isNotAlreadyThere)
                {
                    $licensesIsDifferent = $true;
                }
            }
        }

        if($null -eq $alChapterUpdateUserObj.LicenseAssignment)
        {
            # if any are there they should be removed
            foreach($m365LicenseAssignment in $m365UsrObj.Licenses)
            {
                $licensesIsDifferent = $true;
                break;
            }
        }
        else
        {
            # loop testing for what should not be there
            [bool]$isNotThere = $true;
            foreach($m365LicenseAssignment in $m365UsrObj.Licenses)
            {
                foreach($usrLicenseAssignmentSKU in $usrLicenseAssignmentSkUs)
                {
                    if($m365LicenseAssignment.AccountSkuId.ToLower() -eq $usrLicenseAssignmentSKU.Trim().ToLower())
                    {
                        $isNotThere = $false;
                    }
                }
                if($isNotThere)
                {
                    $licensesIsDifferent = $true;
                    break;
                }
            }
        }
        return $licensesIsDifferent;
    }

    function UpdateIfUserProfileDifferent([PSObject]$tenantDefaultsObj, [PSObject]$alChapterUObj, [PSObject]$m365UsrObj, [bool]$justTestingOnly)
    {
        [bool]$ProfileIsDifferent = $false;
        #put updates here
        if($alChapterUObj.BlockCredential -ne $m365UsrObj.BlockCredential)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " BlockCredential ProfileIsDifferent: " + $alChapterUObj.BlockCredential + " - " + $m365UsrObj.BlockCredential;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing BlockCredential to: " + $alChapterUObj.BlockCredential;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -BlockCredential $alChapterUObj.BlockCredential
            }
        }
        if($alChapterUObj.City -ne $m365UsrObj.City)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " City ProfileIsDifferent: " + $alChapterUObj.City + " - " + $m365UsrObj.City;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing City to: " + $alChapterUObj.City;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -City $alChapterUObj.City
            }
        }
        if($alChapterUObj.Country -ne $m365UsrObj.Country)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " Country ProfileIsDifferent: " + $alChapterUObj.Country + " - " + $alChapterUObj.Country;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing Country to: " + $alChapterUObj.Country;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -Country $alChapterUObj.Country
            }
        }
        if($alChapterUObj.Department -ne $m365UsrObj.Department)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " Department ProfileIsDifferent: " + $alChapterUObj.Department + " - " + $m365UsrObj.Department;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing Department to: " + $alChapterUObj.Department;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -Department $alChapterUObj.Department
            }
        }
        if($alChapterUObj.DisplayName -ne $m365UsrObj.DisplayName)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " DisplayName ProfileIsDifferent: " + $alChapterUObj.DisplayName + " - " + $m365UsrObj.DisplayName;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing DisplayName to: " + $alChapterUObj.DisplayName;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -DisplayName $alChapterUObj.DisplayName
            }
        }
        if($alChapterUObj.FirstName -ne $m365UsrObj.FirstName)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " FirstName ProfileIsDifferent: " + $alChapterUObj.FirstName + " - " + $m365UsrObj.FirstName;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing FirstName to: " + $alChapterUObj.FirstName;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -FirstName $alChapterUObj.FirstName
            }
        }
        if($alChapterUObj.LastName -ne $m365UsrObj.LastName)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " LastName ProfileIsDifferent: " + $alChapterUObj.LastName + " - " + $m365UsrObj.LastName;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing LastName to: " + $alChapterUObj.LastName;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -LastName $alChapterUObj.LastName
            }
        }
        if($alChapterUObj.MobilePhone -ne $m365UsrObj.MobilePhone)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " MobilePhone ProfileIsDifferent: " + $alChapterUObj.MobilePhone + " - " + $m365UsrObj.MobilePhone;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing MobilePhone to: " + $alChapterUObj.MobilePhone;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -MobilePhone $alChapterUObj.MobilePhone
            }
        }
        if($tenantDefaultsObj.Office -ne $m365UsrObj.Office)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " Office ProfileIsDifferent: " + $tenantDefaultsObj.Office + " - " + $m365UsrObj.Office;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing Office to: " + $tenantDefaultsObj.Office;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -Office $tenantDefaultsObj.Office
            }
        }
        if($tenantDefaultsObj.PasswordNeverExpires -ne $m365UsrObj.PasswordNeverExpires)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " PasswordNeverExpires ProfileIsDifferent: " + $tenantDefaultsObj.PasswordNeverExpires + " - " + $m365UsrObj.PasswordNeverExpires;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing PasswordNeverExpires to: " + $tenantDefaultsObj.PasswordNeverExpires;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -PasswordNeverExpires $tenantDefaultsObj.PasswordNeverExpires
            }
        }
        if($alChapterUObj.HomePhone -ne $m365UsrObj.PhoneNumber)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " PhoneNumber ProfileIsDifferent: " + $alChapterUObj.HomePhone + " - " + $m365UsrObj.PhoneNumber;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing PhoneNumber to: " + $alChapterUObj.HomePhone;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -PhoneNumber $alChapterUObj.HomePhone
            }
        }
        if($alChapterUObj.PostalCode -ne $m365UsrObj.PostalCode)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " PostalCode ProfileIsDifferent: " + $alChapterUObj.PostalCode + " - " + $m365UsrObj.PostalCode;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing PostalCode to: " + $alChapterUObj.PostalCode;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -PostalCode $alChapterUObj.PostalCode
            }
        }
        if($alChapterUObj.State -ne $m365UsrObj.State) 
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " State ProfileIsDifferent: " + $alChapterUObj.State + " - " + $m365UsrObj.State;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing State to: " + $alChapterUObj.State;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -State $alChapterUObj.State
            }
        }
        if($alChapterUObj.StreetAddress -ne $m365UsrObj.StreetAddress)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " StreetAddress ProfileIsDifferent: " + $alChapterUObj.StreetAddress + " - " + $m365UsrObj.StreetAddress;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing StreetAddress to: " + $alChapterUObj.StreetAddress;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -StreetAddress $alChapterUObj.StreetAddress
            }
        }
        if($tenantDefaultsObj.StrongPasswordRequired -ne $m365UsrObj.StrongPasswordRequired)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " StrongPasswordRequired ProfileIsDifferent: " + $tenantDefaultsObj.StrongPasswordRequired + " - " + $m365UsrObj.StrongPasswordRequired;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing StrongPasswordRequired to: " + $tenantDefaultsObj.StrongPasswordRequired;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -StrongPasswordRequired $tenantDefaultsObj.StrongPasswordRequired
            }
        }
        if($alChapterUObj.Title -ne $m365UsrObj.Title)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " Title ProfileIsDifferent: " + $alChapterUObj.Title + " - " + $m365UsrObj.Title;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing Title to: " + $alChapterUObj.Title;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -Title $alChapterUObj.Title
            }
        }
        if($tenantDefaultsObj.UsageLocation -ne $m365UsrObj.UsageLocation)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " UsageLocation ProfileIsDifferent: " + $tenantDefaultsObj.UsageLocation + " - " + $m365UsrObj.UsageLocation;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing UsageLocation to: " + $tenantDefaultsObj.UsageLocation;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -UsageLocation $tenantDefaultsObj.UsageLocation
            }
        }
        if($alChapterUObj.PersonalEmail -ne "*NO EMAIL*")
        {
            if($alChapterUObj.PersonalEmail -ne $m365UsrObj.AlternateEmailAddresses)
            {
                $ProfileIsDifferent = $true;
                $logMessage  = $alChapterUObj.ChapterEmail + " AlternateEmailAddress ProfileIsDifferent: " + $alChapterUObj.PersonalEmail + " - " + $m365UsrObj.AlternateEmailAddresses;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                if(!$justTestingOnly)
                {
                    $logMessage  = $alChapterUObj.ChapterEmail + " changing AlternateEmailAddress to: " + $alChapterUObj.PersonalEmail;
                    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -AlternateEmailAddress $alChapterUObj.PersonalEmail
                }
            }
        }

        if($alChapterUObj.M365UserType -ne $m365UsrObj.UserType)
        {
            $ProfileIsDifferent = $true;
            $logMessage  = $alChapterUObj.ChapterEmail + " UserType ProfileIsDifferent: " + $alChapterUObj.M365UserType + " - " + $m365UsrObj.UserType;
            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
            if(!$justTestingOnly)
            {
                $logMessage  = $alChapterUObj.ChapterEmail + " changing UserType to: " + $alChapterUObj.m365UserType;
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                Set-MsolUser -UserPrincipalName $alChapterUObj.ChapterEmail -UserType $alChapterUObj.M365UserType
            }
        }
        [bool]$IfUpdatingLicenses = $false;
        if($IfUpdatingLicenses)
        {
            [bool]$licenseIsDifferent = isUserLicenseAssignmentDifferent $alChapterUObj;
            if($licenseIsDifferent )
            {
                $ProfileIsDifferent = $true;
                $logMessage  = $alChapterUObj.ChapterEmail + " LicenseAssignment ProfileIsDifferent: ";
                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                if(!$justTestingOnly)
                {
                    $logMessage  = $alChapterUObj.ChapterEmail + " changing LicenseAssignment to: " + $alChapterUObj.LicenseAssignment;
                    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    UpdateUserLicense $alChapterUObj;
                }
            }
        }
        
        return $ProfileIsDifferent;
    }

    function ReportOnesToBeDeleted([PSObject]$rbeObjs, [PSObject]$alChapterUsrObjs)
    {
        # filter out ones we ignore
        $m365Objs = Get-MsolUser  | Where-Object {$_.UserPrincipalName -notlike  '*xx*' `
                                    -and $_.UserPrincipalName -ne "TSConferenceRoom_1@algeorgetownarea.org" `
                                    -and $_.Title -ne "Test Account" `
                                    -and $_.UserPrincipalName -notlike  '*onmicrosoft.com'};
        [string]$logMessage  = "alChapterSchema.xls RoleBasedEmail Count: " + $rbeObjs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage  = "alChapterSchema.xls User Row Count: " + $alChapterUsrObjs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage  = "Microsoft 365 Users including Role-based Emails - Row Count: " + $m365Objs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;

        [string]$logMessage  = "alChapterSchema.xls RoleBasedEmail Count: " + $rbeObjs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage  = "alChapterSchema.xls User Row Count: " + $alChapterUsrObjs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage  = "Microsoft 365 Users including Role-based Emails - Row Count: " + $m365Objs.Length.ToString();
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;

        foreach($m365Obj in $m365Objs)
        {
            $notFound = $true;
            foreach($alChapterUsrObj in $alChapterUsrObjs)
            {
                if($null -ne $alChapterUsrObj.ChapterEmail)
                {
                    [string]$m365UserPrincipalName = $m365Obj.UserPrincipalName.ToLower();
                    [string]$alChapterUsrPrincipleName = $alChapterUsrObj.ChapterEmail.ToLower();
                    if($m365UserPrincipalName -eq $alChapterUsrPrincipleName)
                    {
                        $notFound = $false;
                        break;
                    }
                }
            }
            if($notFound)
            {
                # now make sure it is not a role-based email
                foreach($rbeObj in $rbeObjs)
                {
                    if($m365UserPrincipalName -eq $rbeObj.RolebasedEmailAddress.ToLower())
                    {
                        $notFound = $false;
                        break;
                    }
                }
                if($notFound)
                {
                    #report user in m365 that is not in alChapterSchema users and not a role-based email user
                    $logMessage  = $m365Obj.UserPrincipalName + " - Not Found in alChapterSchema Users or role-based email users.";
                    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;

                }
                else
                {
                    $logMessage  = $m365Obj.UserPrincipalName + " - Found in alChapterSchema Users or role-based email users.";
                    #LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                }
            }
        }
    }

    function ReportOnesComingOffLOA([PSObject]$alChapterUsrObjs)
    {
        [int]$numberComingOffLOA = 0;
        $comingOffLOAArray = @();
        foreach($alChapterUsrObj in $alChapterUsrObjs)
        {
            if($alChapterUsrObj.ALStatus -eq "LOA")
            {
                [datetime]$loaEndDate = [datetime]::FromOADate($alChapterUsrObj.LOAEndDate);
                [datetime]$currentDate = Get-Date;
                $ts = New-TimeSpan -Start $currentDate -End $loaEndDate;
                if($ts.Days -eq 5)
                {
                    $numberComingOffLOA++;
                    $comingOffLOA = New-Object PSObject;
                    $comingOffLOA | Add-Member Noteproperty -Name DisplayName -value $alChapterUsrObj.DisplayName;
                    $comingOffLOA | Add-Member Noteproperty -Name LOAEndDate -value $alChapterUsrObj.LOAEndDate;
                    $comingOffLOAArray += $comingOffLOA;
                }
            }
        }
        <#
        if($numberComingOffLOA -gt 0)
        {
            # Create mail message and send to membership@" + $TenantDomain + ".org"
            $emailBody = "":
            foreach($myObj in $comingOffLOAArray)
            {
                $emailBody += $myObj.Display + " is coming off of LOA on " + $comingOffLOA.LOAEndDate + '/n';
            }
            # now mail it
        }
        #>
    }

    function LoadObjectsFromExcelFile([string]$filePathName, [string]$pageName, [int]$startRow)
    {
        $excelData = Import-Excel -Path $filePathName -WorksheetName $pageName -StartRow $startRow  -DataOnly;
        return $excelData;
    }

    function LogLicensesInformation([string]$myLogFile, [PSObject]$tenantLicenseObjs)
    {
        [string]$logMessage = "";
        $logMessage = "----------------------------------------------------------------------------------------"
        LogToMasterLogFile -logFile $myLogFile -message $logMessage;
        $logMessage = "{0,-42} {1,11} {2,12} {3,12} {4,6}" -f "AccountSkuId", "ActiveUnits", "WarningUnits", "ConsumedUnits", "Unused";
        LogToMasterLogFile -logFile $myLogFile -message $logMessage;
        $logMessage = "{0,-42} {1,11} {2,12} {3,12} {4,6}" -f "------------", "-----------", "------------", "-------------", "------";
        LogToMasterLogFile -logFile $myLogFile -message $logMessage;

        foreach($tlObj in $tenantLicenseObjs)
        {
            $remainingUnits = $tlObj.ActiveUnits - $tlObj.ConsumedUnits;
            $logMessage = "{0,-42} {1,11} {2,12} {3,13} {4,6}" -f $tlObj.AccountSkuId, $tlObj.ActiveUnits, $tlObj.WarningUnits, $tlObj.ConsumedUnits, $remainingUnits;
            LogToMasterLogFile -logFile $myLogFile -message $logMessage;
        }
        $logMessage = "----------------------------------------------------------------------------------------"
        LogToMasterLogFile -logFile $myLogFile -message $logMessage;
    }

    function LogToMasterLogFile($logFile, $message)
    {
        $myTheDate = Get-Date;
        $fDate = $myTheDate.ToString("yyyyMMddTHHmmss");
        $lineToWrite = "[" + $fDate + "]" + $message;
        Add-Content $logFile $lineToWrite;
    }

#endregion

    # script starts here
    [string]$tenantDomain = $tenantObj.DomainName;

    #region define APStatus variable definitions updating APStatusTable with Start values
    #[string]$listName = "APStatuses";
    [string]$ProcessName = "ManageTenantM365Users";
    [string]$ProcessCategory = "CHWorkFlow";
    [DateTime]$ProcessStartDate = Get-date;
    #[DateTime]$ProcessStartDate = (Get-date).ToUniversalTime();
    [DateTime]$ProcessStopDateBegins = (Get-date("1/1/1900 00:01")).ToUniversalTime();
    [DateTime]$SProcessStopDateFinished = (Get-date("1/1/2099 00:01")).ToUniversalTime();
    [string]$ProcessStatusStart = "Started";
    [string]$ProcessStatusFinish = "Successful";
    [string]$ProcessProgressStart = "In-progress";
    [string]$ProcessProgressFinish = "Completed";
    
    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $tenantCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $ProcessStopDateBegins `
                                                           -ProcessStatus $ProcessStatusStart `
                                                           -ProcessProgress $ProcessProgressStart `
                                                           -masterLogFilePathAndName $masterLogFilePathAndName;
   
    #endregion

    $logMessage = "Starting ManageM365TenantUsers";
    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    #region load object arrays from Excel sheets
    [string]$tenantDefaultsPageName = "M365TenantDefaults";
    [int]$tenantDefaultsStartRow = 1;
    [string]$alChapterUsersPageName = "M365Users";
    [int]$alChapterUsersStartRow = 1;



    $alChapterTenantDefaultsObj = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $tenantDefaultsPageName $tenantDefaultsStartRow;
    $userObjs = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $alChapterUsersPageName $alChapterUsersStartRow;
    #endregion load object arrays from Excel sheets

    #connect to SharePoint On-line
    Connect-MsolService -Credential $tenantCredentials;

    Connect-AzAccount -Tenant 

    $tenantSubscriptions = Get-MsolAccountSku;
    LogLicensesInformation $masterLogFilePathAndName $tenantSubscriptions;

    $JustTestingOnesToDelete =$false;
    if(!$JustTestingOnesToDelete)
    {
        #loop through user objects
        foreach($uObj in $userObjs)
        {
            # skip if not active
            if($uObj.M365Status -eq "Active" -and `
            $uObj.ALStatus -ne "R" -and `
            $uObj.ALStatus -ne "D" -and `
            $uObj.ALStatus -ne "CV")
            {
                if($uObj.M365UserType -eq "Member")
                {
                    [bool]$chapterEmailProvided = ($null -ne $uobj.ChapterEmail);
                    if($chapterEmailProvided)
                    {
                        #Test if user already there
                        #
                        # Need to change this to look for Chapter Hub object ID in the exchange custom fields
                        #   Need to write the Chapter Hub object ID in the exchange custom fields 
                        #
                        $m365Usr = $null;
                        $m365Usr = Get-MsolUser -UserPrincipalName $uObj.ChapterEmail -ErrorAction SilentlyContinue;
                        if($null -eq $m365Usr) #If user not there then create user
                        {
                            $logMessage  = "Creating User: " + $uObj.DisplayName + " Chapter Email: " + $uObj.ChapterEmail + " Personal Email: " + $uObj.PersonalEmail;
                            LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                            if(!$justTestingOnly)
                            {
                                CreateUser $alChapterTenantDefaultsObj $uObj
                            }
                        }
                        else
                        {
                            # user exists so test for user profile differences
                            #if different then update
                            [bool]$profileIsDifferent = UpdateIfUserProfileDifferent $alChapterTenantDefaultsObj $uobj $m365Usr $justTestingOnly;
                            if($profileIsDifferent)
                            {
                                #Write-Host("Updated User: " + $uObj.ChapterEmail);
                            }
                        }
                    }
                    else
                    {
                        #
                        # Need to send email
                        #
                        $logMessage  = "Member Missing Chapter EMail: ****************************************";
                        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        $logMessage  = "Member Missing Chapter EMail: " + $uObj.DisplayName + " Personal Email: " + $uObj.PersonalEmail;
                        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        $logMessage  = "Member Missing Chapter EMail: ****************************************";
                        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                    }
                }
                else
                {
                    if($uObj.M365UserType -eq "Guest")
                    {
                        [string]$personalEmail = $uObj.PersonalEmail;
                        if($null -ne $personalEmail -and $personalEmail -ne "None On FIle" )
                        {

                            <#
                                https://docs.microsoft.com/en-us/powershell/azure/install-az-ps?view=azps-5.7.0
                                $guestuser = Get-AzureADUser -Filter "userType eq 'Guest'" -All $true | Where-Object {$_.Mail -eq "$personalEmail"} 
                            #>
                            # search for guest here
                            [string]$guestFilter = "userPrincipalName eq '" + $personalEmail + "'";
                            $guestUser = $null;
                            $guestUser = Get-AzureADUser -Filter $guestFilter;
                            #Test if user already there
                            if($null -eq $guestUser)
                            {
                                $logMessage  = "Creating Guest: " + $uObj.DisplayName + " Personal Email: " + $uObj.PersonalEmail
                                LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                                if(!$justTestingOnly)
                                {
                                    #Create new-guest
                                    #CreateAzureADUserGraph $alChapterTenantDefaultsObj $uObj;
                                }
                            }
                            else
                            {
                                # update Guest if Different
                                [bool]$profileIsDifferent = UpdateIfGuestProfileDifferent $alChapterTenantDefaultsObj $uobj $m365Usr $justTestingOnly;
                                if($profileIsDifferent)
                                {
                                    #Write-Host("Updated User: " + $uObj.ChapterEmail);
                                    if(!$justTestingOnly)
                                    {
                                        #Update guest
                                        #CreateAzureADUserGraph $uObj;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    #end loop
    $tenantSubscriptions = Get-MsolAccountSku;
    LogLicensesInformation $masterLogFilePathAndName $tenantSubscriptions;


    [string]$roleBasedEmailsPageName = "RoleBasedEmailAddresses";
    [int]$roleBasedEmailsStartRow = 1;
    $roleBaseEmailObjs = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $roleBasedEmailsPageName $roleBasedEmailsStartRow;

    # load user objects again
    $userObjs_2 = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName $alChapterUsersPageName $alChapterUsersStartRow;

    ReportOnesToBeDeleted $roleBaseEmailObjs $userObjs_2;

    #ReportOnesComingOffLOA $userObjs_2;


    $logMessage = "Finished Managem365alChapterUsers";
    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;


    #region updating APStatusTable with finish values
    $SProcessStopDateFinished = Get-Date;
    #SProcessStopDateFinished = (Get-Date).ToUniversalTime();

    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $tenantCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $SProcessStopDateFinished `
                                                           -ProcessStatus $ProcessStatusFinish `
                                                           -ProcessProgress $ProcessProgressFinish `
                                                           -masterLogFilePathAndName $masterLogFilePathAndName;
    #endregion
    
}
end{}
