#
# UpdateTenantALChapterSchemaFrom_ChapterHub.ps1
#
<#
#           Author: Mike John
#     Date Created: 01/14/2021
#
# Last date edited: 01/14/2021
#   Last edited By: Mike John
# Last Edit Reason: Original
#
Preconditions:
    $chCSVFilePathAndName must exist
    $m365SchemaFilePathAndName must exist
    NALSchema File must have a sheet inside it named $m365UsersSheetName
    $masterLogFile directory must exist

Load $chCSVFilePathAndName objects
Open $m365SchemaFilePathAndName
foreach $chObj in chObjs
    search for corresponding record in NALSchema
    if(!there)
        Insert record at end and resort
    else
        UpdateSchemaRecordIfDifferent
    endif
endforeach
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2)][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3)][string]$chCSVFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4)][string]$m365SchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=5)][string]$masterLogFile
     ,[Parameter(Mandatory=$True,Position=6)][bool]$testingOnly
)
begin {}
process {

    #region function definitions
    function Release_Ref ($ref)
    {
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    function Convert-ToLetter([int]$iCol)
    {
        # works up to ZZ
        $ConvertToLetter = $iAlpha = $iRemainder = $null;
        [double]$iAlpha  = ($iCol/27);
        [int]$iAlpha = [system.math]::floor($iAlpha);
        $iRemainder = $iCol - ($iAlpha * 26);
        if ($iRemainder -gt 26)
        {
            $iAlpha = $iAlpha + 1;
            $iRemainder = $iRemainder -26;
        }   
        if ($iAlpha -gt 0 )
        {
            $ConvertToLetter = [char]($iAlpha + 64);
        }
        if ($iRemainder -gt 0)
        {
            $ConvertToLetter = $ConvertToLetter + [Char]($iRemainder + 64);
        }
        return $ConvertToLetter;
    }

    function R1C1Str([int]$rowNumber, [int]$colNumber)
    {
        [string]$r1c1Str = 'R' + $rowNumber + 'C' + $colNumber;
        return $r1c1Str;
    }

    function FormatPhoneNumber([string]$phoneNumberToFormat)
    {
        [string]$formattedPhoneNumber = $null; #default return $null if number to format is $null
        if($null -ne $phoneNumberToFormat)
        {
            $formattedPhoneNumber = "(XXX) XXX-XXXX"; # default return if number is invalid
            [string]$copyPhoneNumberToFormat = $phoneNumberToFormat -replace '[^0-9]';
            if($null -ne $copyPhoneNumberToFormat)
            {
                #Validate length of phone number - 2817501392 - must be 10 characters after strip
                if($copyPhoneNumberToFormat.Length -eq 10)
                {
                    #  convert string into number
                    [Uint64]$pNumber = [Convert]::ToUInt64($copyPhoneNumberToFormat)
                    # now format number
                    $formattedPhoneNumber = $pNumber.ToString("(###) ###-####");
                }
            }
        }
        return $formattedPhoneNumber;
    }

    function FormatRegularDate([string]$DateStrToFormat)
    {
        if($null -eq $DateStrToFormat)
        {
            return $null;
        }
        else
        {
            if($DateStrToFormat -eq "")
            {
                return $null;
            }
            else
            {
                [string]$tmpDate = [datetime]::parse($DateStrToFormat, $null).ToString('MM/dd/yy')
                return $tmpDate
            }
        }
    }

    function FormatRegularDateTime([string]$DateTimeStrToFormat)
    {
        if($null -eq $DateTimeStrToFormat)
        {
            return $null;
        }
        else
        {
            if($DateTimeStrToFormat -eq "")
            {
                return $null;
            }
            else
            {
                [string]$tmpDate = [datetime]::parse($DateTimeStrToFormat, $null).ToString('MM/dd/yyyy')
                return $tmpDate
            }
        }
    }

    function SetALStatus([PSObject]$chObj)
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

    function ConvertToThisYearBirthDate([string]$birthDateDayMonthStr)
    {
        [string]$birthDateDayMonth = "01/01/1900";
        
        if($null -ne $birthDateDayMonthStr)
        {
            [int]$monthIdx = -1;
            if($birthDateDayMonthStr.IndexOf("-") -ge 0)
            {
                $strArray = $birthDateDayMonthStr.Split("-");
                for($i=0;$i-le $MonthAbreviations.length-1;$i++)
                {
                    if($MonthAbreviations[$i] -eq $strArray[1])
                    {
                        $monthIdx = $i + 1;
                        break;
                    }
                }
                if($monthIdx -ge 0)
                {
                    #we now have the day and month
                    $thisYear       = get-date -Format yyyy;
                    $birthDateMonth = $monthIdx;
                    $birthDateDay   = [int]$strArray[0];
                    $tmp = [datetime]"$birthDateMonth/$birthDateDay/$thisYear";
                    $birthDateDayMonth = $tmp.ToString("MM/dd/yyyy");
                }
            }
            else
            {
                $strArray = $birthDateDayMonthStr.Split("/");
                if($strArray.Count -eq 2 -or $strArray.Count -eq 3)
                {
                    try
                    {
                        $thisYear       = get-date -Format yyyy;
                        $birthDateMonth = [int]$strArray[0];
                        $birthDateDay   = [int]$strArray[1];
                        $tmp = [datetime]"$birthDateMonth/$birthDateDay/$thisYear";
                        $birthDateDayMonth = $tmp.ToString("MM/dd/yyyy");
                    }
                    catch
                    {
                        $birthDateDayMonth ="01/01/1900";
                    }
                }
            }
        }
        return $birthDateDayMonth;
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

    function LoadSDataObjs([string]$filePathName, [string]$pageName, [int]$startRow)
    {
        $defaultData = Import-Excel -Path $filePathName -WorksheetName $pageName -StartRow $startRow -DataOnly;
        return $defaultData;
    }

    function FindMatchingRowNumber($m365UserWS, [PSObject]$chObj)
    {
        [string]$IndividualID = $chObj.IndividualID;
        [PSCustomObject]$returnObj = [PSCustomObject][ordered]@{
                                                                CHIndividualID     = $chObj.IndividualID
                                                                CHEmailAddress     = $chObj.ChapterEmail
                                                                CHLastName         = $chObj.LastName
                                                                CHFirstName        = $chObj.FirstName
                                                                CHInformalName     = $chObj.InformalName
                                                                SchemaIndividualID = "N/A"
                                                                SchemaCHEmail      = "N/A"
                                                                SchemaLastName     = "N/A"
                                                                SchemaFirstName    = "N/A"
                                                                SchemaInformalName = "N/A"
                                                                RowNumber          = -1   #default to not found
                                                               }

        #search for matching entry in IndividualID column
        # maybe we need another way to create this range???????????????????
        $mrRange = $m365UserWS.Range("B2").EntireColumn;
        $foundCell = $null;
        $foundCell = $mrRange.Find($IndividualID);
        if($null -ne $foundCell)
        {
            # it was found, so return the values in our return object
            $rowNumber                    = $foundCell.Row;
            $returnObj.RowNumber          = $rowNumber;
            $returnObj.SchemaIndividualID = $m365UserWS.Cells.Item($rowNumber,$IndividualIDCol).Text;
            $returnObj.SchemaCHEmail      = $m365UserWS.Cells.Item($rowNumber,$ChapterEmailCol).Text;
            $returnObj.SchemaLastName     = $m365UserWS.Cells.Item($rowNumber,$LastNameCol).Text;
            $returnObj.SchemaFirstName    = $m365UserWS.Cells.Item($rowNumber,$FirstNameCol).Text;
            $returnObj.SchemaInformalName = $m365UserWS.Cells.Item($rowNumber,$InformalNameCol).Text;
        }
        return $returnObj;
    }

    function InsertRowAtEnd($m365UserWS, [PSObject]$defaultObj, [PSObject]$chObj)
    {
        # these are constants from the Excel library
        $xlDown = -4121;
        #$xlToRight =-4161;

        $downRangeSorting = $m365UserWS.UsedRange();
        $endDownCellRange = $downRangeSorting.End($xlDown);
        $endRow = $endDownCellRange.Row();

        #fill-in a new row getting values from chapter hub object and default object
        $newRow = $endRow + 1;

        $m365UserWS.Cells.Item($newRow, $DisplayNameCol                     ) = $chObj.DisplayName;
        $m365UserWS.Cells.Item($newRow, $IndividualIDCol                    ) = $chObj.IndividualID;
        $m365UserWS.Cells.Item($newRow, $ChapterEmailCol                    ) = $chObj.ChapterEmail;
        $m365UserWS.Cells.Item($newRow, $PersonalEmailCol                   ) = $chObj.PersonalEmail;
        $m365UserWS.Cells.Item($newRow, $EmailUsedCol                       ) = $chObj.EmailUsed;
        $m365UserWS.Cells.Item($newRow, $ALStatusCol                        ) = $chObj.ALStatus;
        $m365UserWS.Cells.Item($newRow, $FirstNameCol                       ) = $chObj.FirstName;
        $m365UserWS.Cells.Item($newRow, $InformalNameCol                    ) = $chObj.InformalName;
        $m365UserWS.Cells.Item($newRow, $LastNameCol                        ) = $chObj.LastName;
        $m365UserWS.Cells.Item($newRow, $TitleCol                           ) = $defaultObj.Title;
        $m365UserWS.Cells.Item($newRow, $DepartmentCol                      ) = $defaultObj.Department;
        if("NULL" -ne $defaultObj.ReportsTo)
        {
            $m365UserWS.Cells.Item($newRow, $ReportsToCol                   ) = $defaultObj.ReportsTo;
        }
        $m365UserWS.Cells.Item($newRow, $OfficeCol                          ) = $defaultObj.Office;
        $m365UserWS.Cells.Item($newRow, $HomePhoneCol                       ) = $chObj.HomePhone;
        $m365UserWS.Cells.Item($newRow, $MobilePhoneCol                     ) = $chObj.MobilePhone
        $m365UserWS.Cells.Item($newRow, $PreferredPhoneCol                  ) = $chObj.PreferredPhone
        $m365UserWS.Cells.Item($newRow, $StreetAddressCol                   ) = $chObj.StreetAddress;
        $m365UserWS.Cells.Item($newRow, $CityCol                            ) = $chObj.City;
        $m365UserWS.Cells.Item($newRow, $StateCol                           ) = $chObj.State;
        $m365UserWS.Cells.Item($newRow, $PostalCodeCol                      ) = $chObj.PostalCode;
        $m365UserWS.Cells.Item($newRow, $CountryCol                         ) = $defaultObj.Country;
        $m365UserWS.Cells.Item($newRow, $MaritalStatusCol                   ) = $defaultObj.MaritalStatus;
        $m365UserWS.Cells.Item($newRow, $SpouseFirstNameCol                 ) = $chObj.SpouseFirstName;
        $m365UserWS.Cells.Item($newRow, $SpouseLastNameCol                  ) = $chObj.SpouseLastName;
        $m365UserWS.Cells.Item($newRow, $JoinDateCol                        ) = $chObj.JoinDate;
        $m365UserWS.Cells.Item($newRow, $DateResignedCol                    ) = $chObj.DateResigned;
        $m365UserWS.Cells.Item($newRow, $BirthDateDayMonthCol               ).Value = $chObj.BirthDateDayMonth;
        $m365UserWS.Cells.Item($newRow, $CH_UserLastModifiedDateCol         ).Value = $chObj.CH_UserLastModifiedDate;
        $m365UserWS.Cells.Item($newRow, $m365LastModifiedDateCol            ).Value = (Get-Date).ToString("MM/dd/yyyy");
        $m365UserWS.Cells.Item($newRow, $m365StatusCol                      ) = $defaultObj.m365Status;
        $m365UserWS.Cells.Item($newRow, $LOAStartDateCol                    ) = $chObj.LOAStartDate;
        $m365UserWS.Cells.Item($newRow, $LOAEndDateCol                      ) = $chObj.LOAEndDate;
        $m365UserWS.Cells.Item($newRow, $LOAStatusCol                       ) = $chObj.LOAStatus;
        $m365UserWS.Cells.Item($newRow, $LOADetailsCol                      ) = $chObj.LOADetails;
        $m365UserWS.Cells.Item($newRow, $MinTrainingLevelNeededCol          ) = $chObj.MinTrainingLevelNeeded;
        $m365UserWS.Cells.Item($newRow, $ActualTrainingLevelCol             ) = $chObj.ActualTrainingLevel;
        $m365UserWS.Cells.Item($newRow, $TrainingStatusCol                  ) = $chObj.TrainingStatus;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactNameCol            ) = $chObj.EmergencyContactName;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactRelationshipCol    ) = $chObj.EmergencyContactRelationship;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactPhoneCol           ) = $chObj.EmergencyContactPhone;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactEmailCol           ) = $chObj.EmergencyContactEmail;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactAltNameCol         ) = $chObj.EmergencyContactAltName;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactAltRelationshipCol ) = $chObj.EmergencyContactAltRelationship;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactAltPhoneCol        ) = $chObj.EmergencyContactAltPhone;
        $m365UserWS.Cells.Item($newRow, $EmergencyContactAltEmailCol        ) = $chObj.EmergencyContactAltEmail;
        $m365UserWS.Cells.Item($newRow, $m365UserTypeCol                    ) = $defaultObj.M365UserType;
        $m365UserWS.Cells.Item($newRow, $CHRecordTypeCol                    ) = $chObj.CHRecordType;
        $m365UserWS.Cells.Item($newRow, $CHRoleCol                          ) = $chObj.CHRole;
        $m365UserWS.Cells.Item($newRow, $CHMembershipClassificationCol      ) = $chObj.CHMembershipClassification;
        $m365UserWS.Cells.Item($newRow, $CHTypeCol                          ) = $chObj.CHType;
        $m365UserWS.Cells.Item($newRow, $ForceChangePasswordCol             ) = $defaultObj.ForceChangePassword;
        $m365UserWS.Cells.Item($newRow, $BlockCredentialCol                 ) = $defaultObj.BlockCredential;
        $m365UserWS.Cells.Item($newRow, $LicenseAssignmentCol               ) = $defaultObj.LicenseAssignment;

        # now resort the whole Sheet
        [string]$lNameSort = $LastNameLetter + "1:" + $LastNameLetter + $newRow.ToString(); 
        [string]$fNameSort = $FirstNameLetter + "1:" + $FirstNameLetter + $newRow.ToString(); 

        # variable needed for UsedRangeSort
        $empty_Var = [System.Type]::Missing
        $sort_colLastName = $m365UserWS.Range($lNameSort)
        $sort_colFirstName = $m365UserWS.Range($fNameSort)

        $m365UserWS.UsedRange.Sort($sort_colLastName,1,$sort_colFirstName,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1)  | Out-Null;
    }

    function UpdateRowAsNeeded($m365UserWS, [int]$rowNumberFound, [PSObject]$chObj, [bool]$testingOnly)
    {
        [bool]$rowUpdate = $false;
        [string]$chapterEmail = $chObj.ChapterEmail;
        if($null -eq $chObj.ChapterEmail)
        {
            $chapterEmail = $chObj.PersonalEmail;
        }

        #$testValue = $m365UserWS.Cells.Item($rowNumberFound, $DisplayNameCol).Text;
        if($m365UserWS.Cells.Item($rowNumberFound, $DisplayNameCol).Text -ne $chObj.DisplayName)
        {
            $logMessage = "Changing " + $chapterEmail + " DisplayName: " + $m365UserWS.Cells.Item($rowNumberFound, $DisplayNameCol).Text + " to " + $chObj.DisplayName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $DisplayNameCol) = $chObj.DisplayName;
                $rowUpdate = $true;
            }
        }

        if($m365UserWS.Cells.Item($rowNumberFound, $ChapterEmailCol).Text -ne $chObj.ChapterEmail)
        {
            $logMessage = "Changing " + $chapterEmail + " ChapterEmail: " + $m365UserWS.Cells.Item($rowNumberFound, $ChapterEmailCol).Text  + " to " + $chObj.ChapterEmail;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $ChapterEmailCol) = $chObj.ChapterEmail;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $PersonalEmailCol).Text -ne $chObj.PersonalEmail)
        {
            $logMessage = "Changing " + $chapterEmail + " PersonalEmail: " + $m365UserWS.Cells.Item($rowNumberFound, $PersonalEmailCol).Text  + " to " + $chObj.PersonalEmail;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $PersonalEmailCol) = $chObj.PersonalEmail;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmailUsedCol).Text -ne $chObj.EmailUsed)
        {
            $logMessage = "Changing " + $chapterEmail + " EmailUsed: " + $m365UserWS.Cells.Item($rowNumberFound, $EmailUsedCol).Text  + " to " + $chObj.EmailUsed;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmailUsedCol) = $chObj.EmailUsed;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $ALStatusCol).Text -ne $chObj.ALStatus)
        {
            $logMessage = "Changing " + $chapterEmail + " ALStatus: " + $m365UserWS.Cells.Item($rowNumberFound, $ALStatusCol).Text  + " to " + $chObj.ALStatus;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $ALStatusCol) = $chObj.ALStatus;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $FirstNameCol).Text -ne $chObj.FirstName)
        {
            $logMessage = "Changing " + $chapterEmail + " FirstName: " + $m365UserWS.Cells.Item($rowNumberFound, $FirstNameCol).Text  + " to " + $chObj.FirstName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $FirstNameCol) = $chObj.FirstName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $InformalNameCol).Text -ne $chObj.InformalName)
        {
            $logMessage = "Changing " + $chapterEmail + " InformalName: " + $m365UserWS.Cells.Item($rowNumberFound, $InformalNameCol).Text  + " to " + $chObj.InformalName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $InformalNameCol) = $chObj.InformalName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $LastNameCol).Text -ne $chObj.LastName)
        {
            $logMessage = "Changing " + $chapterEmail + " LastName: " + $m365UserWS.Cells.Item($rowNumberFound, $LastNameCol).Text  + " to " + $chObj.LastName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LastNameCol) = $chObj.LastName;
                $rowUpdate = $true;
            }
        }
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $TitleCol).Text -ne $chObj.Title)
        {
            $logMessage = "Changing " + $chapterEmail + " Title: " + $m365UserWS.Cells.Item($rowNumberFound, $TitleCol).Text  + " to " + $chObj.Title;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $TitleCol) = $chObj.Title;
                $rowUpdate = $true;
            }
        }
        #>
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $DepartmentCol).Text -ne $chObj.Department)
        {
            $logMessage = "Changing " + $chapterEmail + " Department: " + $m365UserWS.Cells.Item($rowNumberFound, $DepartmentCol).Text  + " to " + $chObj.Department;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $DepartmentCol) = $chObj.Department;
                $rowUpdate = $true;
            }
        }
        #>
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $ReportsToCol).Text -ne $chObj.ReportsTo)
        {
            $logMessage = "Changing " + $chapterEmail + " ReportsTo: " + $m365UserWS.Cells.Item($rowNumberFound, $ReportsToCol).Text  + " to " + $chObj.ReportsTo;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $ReportsToCol) = $chObj.ReportsTo;
                $rowUpdate = $true;
            }
        }
        #>
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $OfficeCol).Text -ne $chObj.Office)
        {
            $logMessage = "Changing " + $chapterEmail + " Office: " + $m365UserWS.Cells.Item($rowNumberFound, $OfficeCol).Text  + " to " + $chObj.Office;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $OfficeCol) = $chObj.Office;
                $rowUpdate = $true;
            }
        }
        #>
        if($m365UserWS.Cells.Item($rowNumberFound, $HomePhoneCol).Text -ne $chObj.HomePhone)
        {
            $logMessage = "Changing " + $chapterEmail + " HomePhone: " + $m365UserWS.Cells.Item($rowNumberFound, $HomePhoneCol).Text  + " to " + $chObj.HomePhone;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $HomePhoneCol) = $chObj.HomePhone;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $MobilePhoneCol).Text -ne $chObj.MobilePhone)
        {
            $logMessage = "Changing " + $chapterEmail + " MobilePhone: " + $m365UserWS.Cells.Item($rowNumberFound, $MobilePhoneCol).Text  + " to " + $chObj.MobilePhone;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $MobilePhoneCol) = $chObj.MobilePhone;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $PreferredPhoneCol).Text -ne $chObj.PreferredPhone)
        {
            $logMessage = "Changing " + $chapterEmail + " PreferredPhone: " + $m365UserWS.Cells.Item($rowNumberFound, $PreferredPhoneCol).Text  + " to " + $chObj.PreferredPhone;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $PreferredPhoneCol) = $chObj.PreferredPhone;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $StreetAddressCol).Text -ne $chObj.StreetAddress)
        {
            $logMessage = "Changing " + $chapterEmail + " StreetAddress: " + $m365UserWS.Cells.Item($rowNumberFound, $StreetAddressCol).Text  + " to " + $chObj.StreetAddress;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $StreetAddressCol) = $chObj.StreetAddress;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CityCol).Text -ne $chObj.City)
        {
            $logMessage = "Changing " + $chapterEmail + " City: " + $m365UserWS.Cells.Item($rowNumberFound, $CityCol).Text  + " to " + $chObj.City;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CityCol) = $chObj.City;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $StateCol).Text -ne $chObj.State)
        {
            $logMessage = "Changing " + $chapterEmail + " State: " + $m365UserWS.Cells.Item($rowNumberFound, $StateCol).Text  + " to " + $chObj.State;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $StateCol) = $chObj.State;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $PostalCodeCol).Text -ne $chObj.PostalCode)
        {
            $logMessage = "Changing " + $chapterEmail + " PostalCode: " + $m365UserWS.Cells.Item($rowNumberFound, $PostalCodeCol).Text  + " to " + $chObj.PostalCode;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $PostalCodeCol) = $chObj.PostalCode;
                $rowUpdate = $true;
            }
        }
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $CountryCol).Text -ne $chObj.Country)
        {
            $logMessage = "Changing " + $chapterEmail + " Country: " + $m365UserWS.Cells.Item($rowNumberFound, $CountryCol).Text  + " to " + $chObj.Country;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CountryCol) = $chObj.Country;
                $rowUpdate = $true;
            }
        }
        #>
        if($m365UserWS.Cells.Item($rowNumberFound, $MaritalStatusCol).Text -ne $chObj.MaritalStatus)
        {
            $logMessage = "Changing " + $chapterEmail + " MaritalStatus: " + $m365UserWS.Cells.Item($rowNumberFound, $MaritalStatusCol).Text  + " to " + $chObj.MaritalStatus;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $MaritalStatusCol) = $chObj.MaritalStatus;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $SpouseFirstNameCol).Text -ne $chObj.SpouseFirstName)
        {
            $logMessage = "Changing " + $chapterEmail + " SpouseFirstName: " + $m365UserWS.Cells.Item($rowNumberFound, $SpouseFirstNameCol).Text  + " to " + $chObj.SpouseFirstName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $SpouseFirstNameCol) = $chObj.SpouseFirstName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $SpouseLastNameCol).Text -ne $chObj.SpouseLastName)
        {
            $logMessage = "Changing " + $chapterEmail + " SpouseLastName: " + $m365UserWS.Cells.Item($rowNumberFound, $SpouseLastNameCol).Text  + " to " + $chObj.SpouseLastName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $SpouseLastNameCol) = $chObj.SpouseLastName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $JoinDateCol).Text -ne $chObj.JoinDate)
        {
            $logMessage = "Changing " + $chapterEmail + " JoinDate: " + $m365UserWS.Cells.Item($rowNumberFound, $JoinDateCol).Text  + " to " + $chObj.JoinDate;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $JoinDateCol) = $chObj.JoinDate;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $DateResignedCol).Text -ne $chObj.DateResigned)
        {
            $logMessage = "Changing " + $chapterEmail + " DateResigned: " + $m365UserWS.Cells.Item($rowNumberFound, $DateResignedCol).Text  + " to " + $chObj.DateResigned;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $DateResignedCol) = $chObj.DateResigned;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $BirthDateDayMonthCol).Text -ne $chObj.BirthDateDayMonth)
        {
            $logMessage = "Changing " + $chapterEmail + " BirthDateDayMonth: " + $m365UserWS.Cells.Item($rowNumberFound, $BirthDateDayMonthCol).Text  + " to " + $chObj.BirthDateDayMonth;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $BirthDateDayMonthCol) = $chObj.BirthDateDayMonth;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CH_UserLastModifiedDateCol).Text -ne $chObj.CH_UserLastModifiedDate)
        {
            $logMessage = "Changing " + $chapterEmail + " CH_UserLastModifiedDate: " + $m365UserWS.Cells.Item($rowNumberFound, $CH_UserLastModifiedDateCol).Text  + " to " + $chObj.CH_UserLastModifiedDate;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CH_UserLastModifiedDateCol) = $chObj.CH_UserLastModifiedDate;
                $rowUpdate = $true;
            }
        }
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $m365StatusCol).Text -ne $chObj.M365Status)
        {
            $logMessage = "Changing " + $chapterEmail + " M365Status: " + $m365UserWS.Cells.Item($rowNumberFound, $m365StatusCol).Text  + " to " + $chObj.M365Status;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $m365StatusCol) = $chObj.M365Status;
                $rowUpdate = $true;
            }
        }
        #>
        if($m365UserWS.Cells.Item($rowNumberFound, $LOAStartDateCol).Text -ne $chObj.LOAStartDate)
        {
            $logMessage = "Changing " + $chapterEmail + " LOAStartDate: " + $m365UserWS.Cells.Item($rowNumberFound, $LOAStartDateCol).Text  + " to " + $chObj.LOAStartDate;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LOAStartDateCol) = $chObj.LOAStartDate;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $LOAEndDateCol).Text -ne $chObj.LOAEndDate)
        {
            $logMessage = "Changing " + $chapterEmail + " LOAEndDate: " + $m365UserWS.Cells.Item($rowNumberFound, $LOAEndDateCol).Text  + " to " + $chObj.LOAEndDate;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LOAEndDateCol) = $chObj.LOAEndDate;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $LOAStatusCol).Text -ne $chObj.LOAStatus)
        {
            $logMessage = "Changing " + $chapterEmail + " LOAStatus: " + $m365UserWS.Cells.Item($rowNumberFound, $LOAStatusCol).Text  + " to " + $chObj.LOAStatus;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LOAStatusCol) = $chObj.LOAStatus;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $LOADetailsCol).Text -ne $chObj.LOADetails)
        {
            $logMessage = "Changing " + $chapterEmail + " LOADetails: " + $m365UserWS.Cells.Item($rowNumberFound, $LOADetailsCol).Text  + " to " + $chObj.LOADetails;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LOADetailsCol) = $chObj.LOADetails;
                $rowUpdate = $true;
            }
        }
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $MinTrainingLevelNeededCol).Text -ne $chObj.MinTrainingLevelNeeded)
        {
            $logMessage = "Changing " + $chapterEmail + " MinTrainingLevelNeeded: " + $m365UserWS.Cells.Item($rowNumberFound, $MinTrainingLevelNeededCol).Text  + " to " + $chObj.MinTrainingLevelNeeded;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $MinTrainingLevelNeededCol) = $chObj.MinTrainingLevelNeeded;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $ActualTrainingLevelCol).Text -ne $chObj.ActualTrainingLevel)
        {
            $logMessage = "Changing " + $chapterEmail + " ActualTrainingLevel: " + $m365UserWS.Cells.Item($rowNumberFound, $ActualTrainingLevelCol).Text  + " to " + $chObj.ActualTrainingLevel;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $ActualTrainingLevelCol) = $chObj.ActualTrainingLevel;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $TrainingStatusCol).Text -ne $chObj.TrainingStatus)
        {
            $logMessage = "Changing " + $chapterEmail + " TrainingStatus: " + $m365UserWS.Cells.Item($rowNumberFound, $TrainingStatusCol).Text  + " to " + $chObj.TrainingStatus;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $TrainingStatusCol) = $chObj.TrainingStatus;
                $rowUpdate = $true;
            }
        }
        #>
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactNameCol).Text -ne $chObj.EmergencyContactName)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactName: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactNameCol).Text  + " to " + $chObj.EmergencyContactName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactNameCol) = $chObj.EmergencyContactName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactRelationshipCol).Text -ne $chObj.EmergencyContactRelationship)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactRelationship: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactRelationshipCol).Text  + " to " + $chObj.EmergencyContactRelationship;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactRelationshipCol) = $chObj.EmergencyContactRelationship;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactPhoneCol).Text -ne $chObj.EmergencyContactPhone)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactPhone: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactPhoneCol).Text  + " to " + $chObj.EmergencyContactPhone;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactPhoneCol) = $chObj.EmergencyContactPhone;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactEmailCol).Text -ne $chObj.EmergencyContactEmail)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactEmail: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactEmailCol).Text  + " to " + $chObj.EmergencyContactEmail;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactEmailCol) = $chObj.EmergencyContactEmail;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltNameCol).Text -ne $chObj.EmergencyContactAltName)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactAltName: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltNameCol).Text  + " to " + $chObj.EmergencyContactAltName;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltNameCol) = $chObj.EmergencyContactAltName;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltRelationshipCol).Text -ne $chObj.EmergencyContactAltRelationship)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactAltRelationship: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltRelationshipCol).Text  + " to " + $chObj.EmergencyContactAltRelationship;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltRelationshipCol) = $chObj.EmergencyContactAltRelationship;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltPhoneCol).Text -ne $chObj.EmergencyContactAltPhone)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactAltPhone: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltPhoneCol).Text  + " to " + $chObj.EmergencyContactAltPhone;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltPhoneCol) = $chObj.EmergencyContactAltPhone;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltEmailCol).Text -ne $chObj.EmergencyContactAltEmail)
        {
            $logMessage = "Changing " + $chapterEmail + " EmergencyContactAltEmail: " + $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltEmailCol).Text  + " to " + $chObj.EmergencyContactAltEmail;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $EmergencyContactAltEmailCol) = $chObj.EmergencyContactAltEmail;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CHRecordTypeCol).Text -ne $chObj.CHRecordType)
        {
            $logMessage = "Changing " + $chapterEmail + " CHRecordType: " + $m365UserWS.Cells.Item($rowNumberFound, $CHRecordTypeCol).Text  + " to " + $chObj.CHRecordType;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CHRecordTypeCol) = $chObj.CHRecordType;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CHRoleCol).Text -ne $chObj.CHRole)
        {
            $logMessage = "Changing " + $chapterEmail + " CHRole: " + $m365UserWS.Cells.Item($rowNumberFound, $CHRoleCol).Text  + " to " + $chObj.CHRole;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CHRoleCol) = $chObj.CHRole;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CHMembershipClassificationCol).Text -ne $chObj.CHMembershipClassification)
        {
            $logMessage = "Changing " + $chapterEmail + " CHMembershipClassification: " + $m365UserWS.Cells.Item($rowNumberFound, $CHMembershipClassificationCol).Text  + " to " + $chObj.CHMembershipClassification;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CHMembershipClassificationCol) = $chObj.CHMembershipClassification;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $CHTypeCol).Text -ne $chObj.CHType)
        {
            $logMessage = "Changing " + $chapterEmail + " CHType: " + $m365UserWS.Cells.Item($rowNumberFound, $CHTypeCol).Text  + " to " + $chObj.CHType;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $CHTypeCol) = $chObj.CHType;
                $rowUpdate = $true;
            }
        }
        <# Not in Chapter Hub
        if($m365UserWS.Cells.Item($rowNumberFound, $m365UserTypeCol).Text -ne $chObj.M365UserType)
        {
            $logMessage = "Changing " + $chapterEmail + " M365UserType: " + $m365UserWS.Cells.Item($rowNumberFound, $m365UserTypeCol).Text  + " to " + $chObj.M365UserType;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $m365UserTypeCol) = $chObj.M365UserType;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $ForceChangePasswordCol).Text -ne $chObj.ForceChangePassword)
        {
            $logMessage = "Changing " + $chapterEmail + " ForceChangePassword: " + $m365UserWS.Cells.Item($rowNumberFound, $ForceChangePasswordCol).Text  + " to " + $chObj.ForceChangePassword;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $ForceChangePasswordCol) = $chObj.ForceChangePassword;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $BlockCredentialCol).Text -ne $chObj.BlockCredential)
        {
            $logMessage = "Changing " + $chapterEmail + " BlockCredential: " + $m365UserWS.Cells.Item($rowNumberFound, $BlockCredentialCol).Text  + " to " + $chObj.BlockCredential;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $BlockCredentialCol) = $chObj.BlockCredential;
                $rowUpdate = $true;
            }
        }
        if($m365UserWS.Cells.Item($rowNumberFound, $LicenseAssignmentCol).Text -ne $chObj.LicenseAssignment)
        {
            $logMessage = "Changing " + $chapterEmail + " LicenseAssignment: " + $m365UserWS.Cells.Item($rowNumberFound, $LicenseAssignmentCol).Text  + " to " + $chObj.LicenseAssignment;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
            if(!$testingOnly)
            {
                $m365UserWS.Cells.Item($rowNumberFound, $LicenseAssignmentCol) = $chObj.LicenseAssignment;
                $rowUpdate = $true;
            }
        }
        #>
        if($rowUpdate)
        {
            $m365UserWS.Cells.Item($rowNumberFound, $m365LastModifiedDateCol) = (Get-Date).ToString("MM/dd/yyyy");
        }
    }
    #endregion function definitions

    #start working here
    $thisScriptName = $MyInvocation.MyCommand.Name
    $logMessage = "Starting " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;


    #region updating APStatus List with Start values
    # define APStatusTable  variables
    #[string]$listName = "APStatuses";
    [string]$ProcessName = "UpdateTenantALChapterSchema";
    [string]$ProcessCategory = "CHWorkFlow";
    [DateTime]$ProcessStartDate = Get-date;
    #[DateTime]$ProcessStartDate = (Get-date).ToUniversalTime();
    [DateTime]$ProcessStopDateBegins = (Get-date("1/1/1900 00:01")).ToUniversalTime();
    [DateTime]$SProcessStopDateFinished =( Get-date("1/1/2099 00:01")).ToUniversalTime();
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
                                                           -masterLogFilePathAndName $masterLogFile;
    #endregion updating APStatus List with Start values

    #region setup Excel objects
    # we want to reuse the Excel object in more than one function
    # Create an Object Excel.Application using Windows Com interface; This only works on Windows Operating System
    $xl = New-Object -ComObject Excel.Application;

    # stop the confirmations
    $xl.DisplayAlerts = $False;

    # hides the spreadsheet from the screen
    $xl.Visible = $false;

    # if you uncomment this you can see the spreadsheet on the screen as the script runs
    #   This should not be uncommented in production runs
    #$xl.Visible = $true;
    #endregion setup Excel objects

    #region NALSchema.xlsx M365Users Sheet column definitions
    [int]$DisplayNameCol                     = 01;  # A
    [int]$IndividualIDCol                    = 02;  # B
    [int]$ChapterEmailCol                    = 03;  # C
    [int]$PersonalEmailCol                   = 04;  # D 
    [int]$EmailUsedCol                       = 05;  # E
    [int]$ALStatusCol                        = 06;  # F
    [int]$FirstNameCol                       = 07;  # G
    [int]$InformalNameCol                    = 08;  # H
    [int]$LastNameCol                        = 09;  # I
    [int]$TitleCol                           = 10;  # J
    [int]$DepartmentCol                      = 11;  # K
    [int]$ReportsToCol                       = 12;  # L
    [int]$OfficeCol                          = 13;  # M
    [int]$HomePhoneCol                       = 14;  # N
    [int]$MobilePhoneCol                     = 15;  # O
    [int]$PreferredPhoneCol                  = 16;  # P
    [int]$StreetAddressCol                   = 17;  # Q
    [int]$CityCol                            = 18;  # R
    [int]$StateCol                           = 19;  # S
    [int]$PostalCodeCol                      = 20;  # T
    [int]$CountryCol                         = 21;  # U
    [int]$MaritalStatusCol                   = 22;  # V
    [int]$SpouseFirstNameCol                 = 23;  # W
    [int]$SpouseLastNameCol                  = 24;  # X
    [int]$JoinDateCol                        = 25;  # Y
    [int]$DateResignedCol                    = 26;  # Z
    [int]$BirthDateDayMonthCol               = 27;  # AA
    [int]$CH_UserLastModifiedDateCol         = 28;  # AB
    [int]$m365LastModifiedDateCol            = 29;  # AC
    [int]$m365StatusCol                      = 30;  # AD
    [int]$LOAStartDateCol                    = 31;  # AE
    [int]$LOAEndDateCol                      = 32;  # AF
    [int]$LOAStatusCol                       = 33;  # AG
    [int]$LOADetailsCol                      = 34;  # AH
    [int]$MinTrainingLevelNeededCol          = 35;  # AI
    [int]$ActualTrainingLevelCol             = 36;  # AJ
    [int]$TrainingStatusCol                  = 37;  # AK
    [int]$EmergencyContactNameCol            = 38;  # AL
    [int]$EmergencyContactRelationshipCol    = 39;  # AM
    [int]$EmergencyContactPhoneCol           = 40;  # AN
    [int]$EmergencyContactEmailCol           = 41;  # AO
    [int]$EmergencyContactAltNameCol         = 42;  # AP
    [int]$EmergencyContactAltRelationshipCol = 43;  # AQ
    [int]$EmergencyContactAltPhoneCol        = 44;  # AR
    [int]$EmergencyContactAltEmailCol        = 45;  # AS
    [int]$CHRecordTypeCol                    = 46;  # AT
    [int]$CHRoleCol                          = 47;  # AU
    [int]$CHMembershipClassificationCol      = 48;  # AV
    [int]$CHTypeCol                          = 49;  # AW
    [int]$m365UserTypeCol                    = 50;  # AX
    [int]$ForceChangePasswordCol             = 51;  # AY
    [int]$BlockCredentialCol                 = 52;  # AZ
    [int]$LicenseAssignmentCol               = 53;  # BA

    [string]$FirstNameLetter = Convert-ToLetter $FirstNameCol;
    [string]$LastNameLetter = Convert-ToLetter $LastNameCol;

    #endregion NALSchema.xlsx column definitions


    #region NALSchema.xlsx UpdateStatus Sheet column definitions
    $LastUpdateDateTimeCol = 01;  # A
    $LastProcessCol        = 02;  # B
    #endregion NALSchema.xlsx column definitions

    #region MonthAbreviations
    [string[]]$MonthAbreviations = @("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec");
    #end region MonthAbreviations

    #region load object arrays from Excel and CSV file

    [string]$m365UsersSheetName         = "M365Users";
    [string]$m365DefaultTenantSheetName = "M365TenantDefaults";
    [int]$m365DefaultTenantStartRow     = 1;
    [string]$m365UpdateStatusSheetName  = "UpdateStatus";


    [PSObject]$siteDefaultObjs = LoadSDataObjs -filePathName $m365SchemaFilePathAndName $m365DefaultTenantSheetName $m365DefaultTenantStartRow;
    [PSObject]$chListObjs = Import-Csv -Path $chCSVFilePathAndName;
    #endregion load object arrays

    # open NALSchema spreadsheet
    $wb = $xl.Workbooks.Open($m365SchemaFilePathAndName);

    # open m365Users page
    $m365UsersWS = $wb.WorkSheets | where-object {$_.Name -eq $m365UsersSheetName};
    $updateStatusWS = $wb.WorkSheets | where-object {$_.Name -eq $m365UpdateStatusSheetName};

    # ******************************************************** Delete this after determining it is not needed  ****************
    #connect to SharePoint On-line
    #Connect-MsolService -Credential $tenantCredentials;

    #region update Schema sheet
    foreach($chListObj in $chListObjs)
    {
        [string]$localCHType   =  $chListObj."Type";
        [string]$localCHRole   =  $chListObj."Role";
        [string]$localCMMC     = $chListObj."Membership Classification";
        [string]$localALStatus =  SetALStatus $chListObj;
        [string]$alStatus = $localALStatus;

        if($alStatus -ne "UK")
        {
            [PSObject]$chObj = LoadCHObj -tenantDefaultObj $siteDefaultObjs[0] -chCsvObj $chListObj -alStatus $alStatus;
            $retObj = FindMatchingRowNumber -m365UserWS $m365UsersWS -chObj $chObj;
            if($retObj.RowNumber -gt 0)
            {
                # if any column compare fails then update
                $logMessage = "ID: " + $chListObj."Individual ID" + " - User: " + $chListObj."Last Name" + ", " + $chListObj."First Name" + " Checking for Updates";
                .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
                UpdateRowAsNeeded -m365UserWS $m365UsersWS -rowNumberFound $retObj.RowNumber -chObj $chObj -testingOnly $testingOnly;
            }
            else
            {
                if($retObj.RowNumber -eq -1)
                {
                    # InsertRow using Chapter Hub object and Site Default Object
                    $logMessage = "ID: " + $chListObj."Individual ID" + " - Adding User " + $chListObj."Last Name" + ", " + $chListObj."First Name";
                    .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
                    if(!$testingOnly)
                    {
                        InsertRowAtEnd -m365UserWS $m365UsersWS -defaultObj $siteDefaultObjs[0] -chObj $chObj;
                    }
                }
                else
                {
                    $logMessage = "Error " +  $chListObj."Last Name" + ", " + $chListObj."First Name" + " - RowNumber Returned: " + $retObj.RowNumber.ToString();
                    .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
                }
            }
        }
        else
        {
            $logMessage = "UK Type " + $chListObj."Last Name" + ", " + $chListObj."First Name";
            $logMessage += " localCHType: " + $localCHType;
            $logMessage += " localCHRole: " + $localCHRole;
            $logMessage += " localCMMC: " + $localCMMC;
            $logMessage += " localALStatus: " + $localALStatus;
            .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
        }
    }

    #endregion update Schema sheet
    
    #region resort the sheet    
    [int]$lastRow = $m365UsersWS.UsedRange.Rows.Count

    # now resort the whole Sheet
    [string]$lNameSort = $LastNameLetter + "1:" + $LastNameLetter + $lastRow.ToString(); 
    [string]$fNameSort = $FirstNameLetter + "1:" + $FirstNameLetter + $lastRow.ToString(); 

    # variable needed for UsedRangeSort
    $empty_Var = [System.Type]::Missing
    $sort_colLastName = $m365UsersWS.Range($lNameSort)
    $sort_colFirstName = $m365UsersWS.Range($fNameSort)

    $m365UsersWS.UsedRange.Sort($sort_colLastName,1,$sort_colFirstName,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1) | Out-Null;
    #endregion resort the sheet    

    # UpdateStatus sheet
    $curDateTime = (Get-Date).ToString('MM/dd/yy HH:mm:ss');
    $thisProcess = "UpdateALChapterSchema";
    $updateStatusRowNumber = 2;
    $updateStatusWS.Cells.Item($updateStatusRowNumber, $LastUpdateDateTimeCol) = $curDateTime;
    $updateStatusWS.Cells.Item($updateStatusRowNumber, $LastProcessCol) = $thisProcess;

    #region clean up Excel stuff still in memory
    $logMessage = "Closing Spreadsheet";
    .\LogManagement\WriteToLogFile -logFile $masterLogFile -message $logMessage;
    $wb.Close($true); # Close workbook and save changes
    $xl.quit(); # Quit Excel...

    Get-Process -Name "Excel" -ErrorAction Ignore | Stop-Process -Force

    if($false)
    {
        Release_Ref($m365UsersWS) | Out-Null;
        Release_Ref($wb) | Out-Null;
        Release_Ref($xl) | Out-Null;
    }
    #endregion clean up Excel stuff still in memory

    #region updating APStatus List with finish values
    $SProcessStopDateFinished = (Get-Date).ToLocalTime();
    #$SProcessStopDateFinished = (Get-Date).ToUniversalTime();

    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $tenantCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $SProcessStopDateFinished `
                                                           -ProcessStatus $ProcessStatusFinish `
                                                           -ProcessProgress $ProcessProgressFinish `
                                                           -masterLogFilePathAndName $masterLogFile;
    #endregion APStatus List with finish values
}
end{}
