# MembersInSchemaNotInCHExport.ps1
<#
#           Author: Mike John
#     Date Created: 06/08/2021
#
# Last date edited: 06/08/2021
#   Last edited By: Mike John
# Last Edit Reason: Original
#
This changes nothing, it only reports which users are no longer active in Chapter Hub

#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2)][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3)][string]$chCSVFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4)][string]$m365SchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=5)][string]$masterLogFile
)
begin {}
process 
{
    #region function definitions
    function SetALStatus([PSObject]$chObj, [string]$tenantAbbreviation)
    {
        # this is what is returned if assignment fails
        [string]$alStatus    = "UK";
        [string]$chType      = $chObj."Type";
        [string]$chRole      = $chObj."Role";
        [string]$chLOAStatus = $chObj."LOAStatus";
        [string]$chMC        = $chObj."Membership Classification";
        if($chType -eq "Community Volunteer")
        {
                $alStatus = "CV";
        }
        else
        {
            if($chMC -eq "PAL")
            {
                $alStatus = "PAL";
            }
            else
            {
                if($chRole -eq "Nonvoting" -and $chType -eq "Full Member")
                {
                    $alStatus = "NV";
                }
                else
                {
                    if($chType -eq "Full Member" -and $chRole -eq "Voting")
                    {
                        if($chLOAStatus -eq "1")
                        {
                            $alStatus = "LOA";
                        }
                        else
                        {
                            $alStatus = "V";
                        }
                    }
                }
            }
        }
        return $alStatus;
    }

    function IsMemberMissingInCHExport($memberObj, $chMemberObjs)
    {
        [bool]$isMissing = $true;
        foreach($chMemberObj in $chMemberObjs)
        {
            if($chMemberObj.xx = "yy")
            {
                if($memberObj.IndividualID = $chMemberObj."Individual ID")
                {
                    $isMissing = $false;
                    break;
                }
            }
        }

        return $isMissing;
    }


    function LoadCHObj([PSObject]$tenantDefaultObj, [PSObject]$chCsvObj, [string]$alStatus)
    {
        [String]$displayName = $chCsvObj."Informal Name" + " " + $chCsvObj."Last Name";
        
        #region Phone Numbers
        [String]$homePhone;
        [String]$mobilePhone;
        [String]$emergencyContactPhone;
        [String]$emergencyContactAltPhone;
        [String]$emailUsed = "ChapterEmail";
        if($chCsvObj."Preferred Email" -ne "Alternate")
        {
            $emailUsed   = "Personal";
        }
        [String]$personalEmail = $chCsvObj."Personal Email"
        if($null -eq $personalEmail)
        {
            $personalEmail = "None On File";
        }
        [String]$homePhone = FormatPhoneNumber $chCsvObj."Home Phone";
        if($homePhone -eq "(XXX) XXX-XXXX")
        {
            $homePhone = $null;
        }
        [String]$mobilePhone = FormatPhoneNumber $chCsvObj."Mobile Phone";
        if($mobilePhone -eq "(XXX) XXX-XXXX")
        {
            $mobilePhone = $null;
        }
        [String]$emergencyContactPhone   = FormatPhoneNumber $chCsvObj."Emergency Contact Phone";
        if($emergencyContactPhone -eq "(XXX) XXX-XXXX")
        {
            $emergencyContactPhone = $null;
        }
        [String]$emergencyContactAltPhone = FormatPhoneNumber $chCsvObj."Emergency Contact Alt - Phone";
        if($emergencyContactAltPhone -eq "(XXX) XXX-XXXX")
        {
            $emergencyContactAltPhone = $null;
        }
        
        #endregion phone numbers

        [string]$birthdateMonthDay = ConvertToThisYearBirthdate -BirthDateDayMonth $chCsvObj."BirthdateMonthDay";
        [string]$JoinDate = FormatRegularDateTime -DateTimeStrToFormat $chCsvObj."FirstMembershipDate";
        [string]$DateResigned = FormatRegularDateTime -DateTimeStrToFormat $chCsvObj."DateResigned";
        [string]$IndividualLastModifiedDate = FormatRegularDateTime -DateTimeStrToFormat $chCsvObj."IndividualLastModifiedDate";
        [string]$LOAStartDate = FormatRegularDateTime -DateTimeStrToFormat $chCsvObj."LOAStartDate";
        [string]$LOAEndDate = FormatRegularDateTime -DateTimeStrToFormat $chCsvObj."LOAEndDate";

        $chLoadObj = New-Object PSObject;
        $chLoadObj | Add-Member Noteproperty -Name DisplayName                     -value $displayName;
        $chLoadObj | Add-Member Noteproperty -Name IndividualID                    -value $chCsvObj."Individual ID";
        $chLoadObj | Add-Member Noteproperty -Name ChapterEmail                    -value $chCsvObj."Alternate Email";
        $chLoadObj | Add-Member Noteproperty -Name PersonalEmail                   -value $personalEmail;
        $chLoadObj | Add-Member Noteproperty -Name EmailUsed                       -value $emailUsed;
        $chLoadObj | Add-Member Noteproperty -Name ALStatus                        -value $alStatus;
        $chLoadObj | Add-Member Noteproperty -Name FirstName                       -value $chCsvObj."First Name";
        $chLoadObj | Add-Member Noteproperty -Name InformalName                    -value $chCsvObj."Informal Name";
        $chLoadObj | Add-Member Noteproperty -Name LastName                        -value $chCsvObj."Last Name";
        if($alStatus -eq "CV")
        {
            $chLoadObj | Add-Member Noteproperty -Name Title                           -value $null;
            $chLoadObj | Add-Member Noteproperty -Name Department                      -value $null;
            $chLoadObj | Add-Member Noteproperty -Name ReportsTo                       -value $null;
            $chLoadObj | Add-Member Noteproperty -Name Office                          -value $null;
        }
        else
        {
            $chLoadObj | Add-Member Noteproperty -Name Title                           -value $tenantDefaultObj.Title;
            $chLoadObj | Add-Member Noteproperty -Name Department                      -value $tenantDefaultObj.Department;
            $chLoadObj | Add-Member Noteproperty -Name ReportsTo                       -value $null;
            $chLoadObj | Add-Member Noteproperty -Name Office                          -value $tenantDefaultObj.Office;
        }
        $chLoadObj | Add-Member Noteproperty -Name HomePhone                       -value $homePhone;
        $chLoadObj | Add-Member Noteproperty -Name MobilePhone                     -value $mobilePhone;
        $chLoadObj | Add-Member Noteproperty -Name PreferredPhone                  -value $chCsvObj."Preferred Phone";
        $chLoadObj | Add-Member Noteproperty -Name StreetAddress                   -value $chCsvObj."Mailing Street";
        $chLoadObj | Add-Member Noteproperty -Name City                            -value $chCsvObj."Mailing City";
        $chLoadObj | Add-Member Noteproperty -Name State                           -value $chCsvObj."Mailing State/Province";
        $chLoadObj | Add-Member Noteproperty -Name PostalCode                      -value $chCsvObj."Mailing Zip/Postal Code";
        $chLoadObj | Add-Member Noteproperty -Name Country                         -value $tenantDefaultObj.Country;
        $chLoadObj | Add-Member Noteproperty -Name MaritalStatus                   -value $chCsvObj."Marital Status";
        $chLoadObj | Add-Member Noteproperty -Name SpouseFirstName                 -value $chCsvObj."Spouse First Name";
        $chLoadObj | Add-Member Noteproperty -Name SpouseLastName                  -value $chCsvObj."Spouse Last Name";
        $chLoadObj | Add-Member Noteproperty -Name JoinDate                        -value $JoinDate;
        $chLoadObj | Add-Member Noteproperty -Name DateResigned                    -value $DateResigned;
        $chLoadObj | Add-Member Noteproperty -Name BirthdateDayMonth               -value $birthdateMonthDay;
        $chLoadObj | Add-Member Noteproperty -Name CH_UserLastModifiedDate         -value $IndividualLastModifiedDate;
        if($alStatus -eq "CV")
        {
            $chLoadObj | Add-Member Noteproperty -Name m365Status                      -value $null;
        }
        else
        {
            $chLoadObj | Add-Member Noteproperty -Name m365Status                      -value $tenantDefaultObj.m365Status;
        }
        $chLoadObj | Add-Member Noteproperty -Name LOAStartDate                    -value $LOAStartDate;
        $chLoadObj | Add-Member Noteproperty -Name LOAEndDate                      -value $LOAEndDate;
        $chLoadObj | Add-Member Noteproperty -Name LOAStatus                       -value $chCsvObj."LOAStatus";
        $chLoadObj | Add-Member Noteproperty -Name LOADetails                      -value $chCsvObj."LOADetails";
        if($alStatus -eq "CV")
        {
            $chLoadObj | Add-Member Noteproperty -Name MinTrainingLevelNeeded          -value $null;
            $chLoadObj | Add-Member Noteproperty -Name ActualTrainingLevel             -value $null;
            $chLoadObj | Add-Member Noteproperty -Name TrainingStatus                  -value $null;
        }
        else
        {
            $chLoadObj | Add-Member Noteproperty -Name MinTrainingLevelNeeded          -value $tenantDefaultObj.MinTrainingLevelNeeded;
            $chLoadObj | Add-Member Noteproperty -Name ActualTrainingLevel             -value $tenantDefaultObj.ActualTrainingLevel;
            $chLoadObj | Add-Member Noteproperty -Name TrainingStatus                  -value $tenantDefaultObj.TrainingStatus;
        }
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactName            -value $chCsvObj."Emergency Contact Name";
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactRelationship    -value $chCsvObj."Emergency Contact Relationship";
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactPhone           -value $emergencyContactPhone;
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactEmail           -value $chCsvObj."Emergency Contact Email";
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactAltName         -value $chCsvObj."Emergency Contact Alt - Name";
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactAltRelationship -value $chCsvObj."Emergency Contact Alt - Relationship";
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactAltPhone        -value $emergencyContactAltPhone;
        $chLoadObj | Add-Member Noteproperty -Name EmergencyContactAltEmail        -value $chCsvObj."Emergency Contact Alt - Email";
        $chLoadObj | Add-Member Noteproperty -Name CHRecordType                    -value $chCsvObj."Record type";
        $chLoadObj | Add-Member Noteproperty -Name CHRole                          -value $chCsvObj."Role";
        $chLoadObj | Add-Member Noteproperty -Name CHMembershipClassification      -value $chCsvObj."Membership Classification";
        $chLoadObj | Add-Member Noteproperty -Name CHType                          -value $chCsvObj."Type";
        if($alStatus -eq "CV")
        {
            $chLoadObj | Add-Member Noteproperty -Name ForceChangePassword             -value $null;
            $chLoadObj | Add-Member Noteproperty -Name BlockCredential                 -value $null;
            $chLoadObj | Add-Member Noteproperty -Name LicenseAssignment               -value $null;
            return $chLoadObj;
        }
        else
        {
            $chLoadObj | Add-Member Noteproperty -Name ForceChangePassword             -value $tenantDefaultObj.ForceChangePassword;
            $chLoadObj | Add-Member Noteproperty -Name BlockCredential                 -value $tenantDefaultObj.BlockCredential;
            $chLoadObj | Add-Member Noteproperty -Name LicenseAssignment               -value $tenantDefaultObj.LicenseAssignment;
            return $chLoadObj;
        }
    }

    function LoadExportArray($chObjs, $alChapterTenantDefaultsObj)
    {
        [PSObject]$chArray =@();
        foreach($chObj in $chObjs)
        {
            [string]$localALStatus =  SetALStatus $chListObj $tenantAbbreviation;
            if($localALStatus -ne "CV" -and $localALStatus -ne "UK" )
            {
                $chMemberObj = LoadCHObj -tenantDefaultObj $alChapterTenantDefaultsObj -chCsvObj $chObj -$alStatus $alStatus;
                $chArray += $chMemberObj;
            }
        }
        return $chArray
    }
    function LoadObjectsFromExcelFile([string]$filePathName, [string]$pageName, [int]$startRow)
    {
        $excelData = Import-Excel -Path $filePathName -WorksheetName $pageName -StartRow $startRow  -DataOnly;
        return $excelData;
    }
    #endregion function definitions

    #start working here
    $thisScriptName = $MyInvocation.MyCommand.Name
    [string]$logMessage = "Starting " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;

    #region load object arrays from Excel and CSV file
    [string]$m365UsersSheetName         = "M365Users";
    [int]$m365UsersStartRow             = 1;
    [string]$m365DefaultTenantSheetName = "M365TenantDefaults";
    [int]$m365DefaultTenantStartRow     = 1;

    $alChapterTenantDefaultsObj = LoadObjectsFromExcelFile $alChapterSchemaFilePathAndName  $m365DefaultTenantSheetName -startRow $m365DefaultTenantStartRow;
    $memberObjs = LoadObjectsFromExcelFile -filePathName $alChapterSchemaFilePathAndName -pageName $m365UsersSheetName -startRow $m365UsersStartRow;
    

    #[PSObject]$memberObjs = LoadSDataObjs -filePathName $m365SchemaFilePathAndName $m365DefaultTenantSheetName $m365DefaultTenantStartRow;
    [PSObject]$chListObjs = Import-Csv -Path $chCSVFilePathAndName;

    $chMemberObjs = LoadExportArray -chListObjs $chListObjs -siteDefaultObjs $alChapterTenantDefaultsObj;
    #endregion load object arrays

    [int]$missingCount = 0;
    foreach($memberObj in $memberObjs)
    {
        [bool]$notResignedOrDeceased = memberObj
        if($notResignedOrDeceased)
        {
            [bool]$isMissing = IsMemberMissingInCHExport -memberObj $memberObj -chMemberObjs $chMemberObjs
            if($isMissing)
            {
                $missingCount++;
                # WriteToLog
                $logMessage = "Member: " + $memberObj.DisplayName + " is Missingmfrom Chapter Hub export file";
                .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            }
        }
    }
    if($missingCount -gt 0)
    {
        # email notification to review $masterLogFile
    }
}