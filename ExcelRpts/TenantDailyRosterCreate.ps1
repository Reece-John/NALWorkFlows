#
# TaskTenantDailyRosterCreate.ps1
#
<#
        Author: Mike John
  Date Created: 9/20/2020

    Updated By: Mike John
       Updated: 05/29/2020
       Reason Updated: Added ActiveDate & InactiveDate to role-based emails
                This allows setting up changes before the PCOP (Peaceful Change Of Power
    Updated By: Mike John
       Updated: 12/31/2020
Reason Updated: Added separate entries for Process Status List
                Created PDF for Emergency Contact List
                Change parameters to accommodate PDFs
    Updated By: Mike John
       Updated: 12/18/2020
Reason Updated: Added Role-base Email sheet and updating Process Status List in SharePoint
    Updated By: Mike John
       Updated: 9/20/2020
Reason Updated: Original

Preconditions:
    $m365SchemaFilePathAndName must exist
    NALSchema File must have sheets inside it named $m365UsersSheetName, $m365roleBasedEmailAddressesSheetName
    $masterLogFilePathAndName directory must exist

Load $m365Schema objects
Open $chapterRosterTemplate.xlsx
foreach $userObj in userObjs
    
endforeach

if making PDF delete all non-relevant sheets and save as PDF
Else
Save as XLSX and protect workbook and all sheets

#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1 )][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2 )][PSObject]$tenantObj
     ,[Parameter(Mandatory=$True,Position=3)][string]$workbookFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4)][string]$pdfFilePathAndName
     ,[Parameter(Mandatory=$True,Position=5)][string]$rosterMembershipSheetName
     ,[Parameter(Mandatory=$True,Position=6)][string]$rosterBirthdaysThisMonthSheetName
     ,[Parameter(Mandatory=$True,Position=7)][string]$rosterBirthdaysNextMonthSheetName
     ,[Parameter(Mandatory=$True,Position=8)][string]$rosterEmergencyContactsSheetName
     ,[Parameter(Mandatory=$True,Position=9)][string]$rosterRolebasedEmailsSheetName
     ,[Parameter(Mandatory=$True,Position=10)][string]$m365SchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=11)][string]$m365UsersSheetName
     ,[Parameter(Mandatory=$True,Position=12)][string]$m365roleBasedEmailAddressesSheetName
     ,[Parameter(Mandatory=$True,Position=13)][bool]$protectWorkBookAndSheets
     ,[Parameter(Mandatory=$True,Position=14)][ValidateSet("ChapterRoster","PDFRoster","PDFEmergencyContacts")][string]$typeOfReport
     ,[Parameter(Mandatory=$True,Position=15)][string]$masterLogFilePathAndName
)
begin {}
process {

    #region function definitions
    function IsDateValidAndGt1900($dateIn)
    {
        $dateChecking = $dateIn -as [DateTime];
        if(!$dateChecking)
        {
            return $false;
        }
        else
        {
            $startOfCentury = Get-Date -Date 'January 1, 1900 0:00:00 AM'
            if($dateChecking -gt $startOfCentury)
            {
                return $true;
            }
            else
            {
                return $false;
            }
        }
    }
    function ConverToPSDateTime($dateToConvert)
    {
        $bd = $null;
        if($null -ne $dateToConvert)
        {
            $bdType = $dateToConvert.GetType();
            if($bdType.Name -ne "DateTime")
            {
                $bd = [datetime]::FromOADate($dateToConvert);
            }
            else
            {
                $bd = $dateToConvert;
            }
        }
        return $bd;
    }

    function IsVolunteersBirthdayThisMonth([PSObject]$userObj)
    {
        [bool]$isSelected = $false;
        $bd = ConverToPSDateTime -dateToConvert $userObj.BirthDateDayMonth;
        if($null -ne $bd)
        {
            if($bd.Year -gt 1900)
            {
                [dateTime]$curDateTime = Get-Date;
                if($bd.Month -eq $curDateTime.Month)
                {
                    $isSelected = $true;
                }
            }
        }
        return $isSelected;
    }

    function IsVolunteersBirthdayNextMonth([PSObject]$userObj)
    {
        [bool]$isSelected = $false;
        $bd = ConverToPSDateTime -dateToConvert $userObj.BirthDateDayMonth;
        if($null -ne $bd)
        {
            if($bd.Year -gt 1900)
            {
                [dateTime]$curDateTime = Get-Date;
                [int]$nextMonth = $curDateTime.Month + 1;
                if($nextMonth -eq 13)
                {
                    $nextMonth = 1;
                }
                if($bd.Month -eq $nextMonth)
                {
                    $isSelected = $true;
                }
            }
        }
        return $isSelected;
    }

    function IsLeapYear([int]$YearToValidate)
    {
        [bool]$isYearLeapYear = $false;
        if ($YearToValidate / 400 -is [int])
        {
            $isYearLeapYear = $true;
        }
        else
        {
            if ($YearToValidate / 100 -is [int])
            {
                $isYearLeapYear = $true;
            }
            else
            {
                if ($YearToValidate / 4 -is [int])
                {
                    $isYearLeapYear = $true;
                }
            }
        }
        return $isYearLeapYear;
    }

    function AddRowToMembership($memWS, [PSObject]$userObj, [int]$memRowNum)
    {
        [string]$email = "";
        if($userObj.EmailUsed -eq "ChapterEmail")
        {
            $email = $userObj.ChapterEmail;
        }
        else
        {
            $email = $userObj.PersonalEmail;
        }
        $birthDay = ConverToPSDateTime -dateToConvert $userObj.BirthDateDayMonth;
        if($null -ne $birthDay)
        {
            if($birthDay.Year -le 1900)
            {
                $birthDay = $null;
            }
        }

        $memWS.Cells.Item($memRowNum, $M_LastNameCol    ) = $userObj.LastName;
        $memWS.Cells.Item($memRowNum, $M_InformalNameCol) = $userObj.InformalName;
        $memWS.Cells.Item($memRowNum, $M_StatusCol      ) = $userObj.ALStatus;
        $memWS.Cells.Item($memRowNum, $M_EmailAddressCol) = $email;
        $memWS.Cells.Item($memRowNum, $M_HomeAddressCol ) = $userObj.StreetAddress;
        $memWS.Cells.Item($memRowNum, $M_CityCol        ) = $userObj.City;
        $memWS.Cells.Item($memRowNum, $M_STCol          ) = $userObj.State;
        $memWS.Cells.Item($memRowNum, $M_ZipCol         ) = $userObj.PostalCode;
        $memWS.Cells.Item($memRowNum, $M_HomePhoneCol   ) = $userObj.HomePhone;
        $memWS.Cells.Item($memRowNum, $M_CellCol        ) = $userObj.MobilePhone;
        $memWS.Cells.Item($memRowNum, $M_BirthDateCol   ) = $birthDay;
        $memWS.Cells.Item($memRowNum, $M_SpouseCol      ) = $userObj.SpouseFirstName + " " + $userObj.SpouseLastName;
        $memWS.Cells.Item($memRowNum, $M_JoinDateCol    ) = $userObj.JoinDate;
    }

    function AddRowToBirthdaysThisMonth($btmWS, [PSObject]$userObj, [int]$btmRowNum)
    {
        $btmWS.Cells.Item($btmRowNum, $BTM_InformalLastNameCol) = $userObj.InformalName + " " + $userObj.LastName;
        $btmWS.Cells.Item($btmRowNum, $BTM_BirthDateCol       ) = ConverToPSDateTime -dateToConvert $userObj.BirthDateDayMonth;
    }

    function AddRowToBirthdaysNextMonth($bnmWS, [PSObject]$userObj, [int]$bnmRowNum)
    {
        $bnmWS.Cells.Item($bnmRowNum, $BNM_InformalLastNameCol) = $userObj.InformalName + " " + $userObj.LastName;
        $bnmWS.Cells.Item($bnmRowNum, $BNM_BirthDateCol       ) = ConverToPSDateTime -dateToConvert $userObj.BirthDateDayMonth;
        $bnmWS.Cells.Item($bnmRowNum, $BNM_HomeAddressCol     ) = $userObj.StreetAddress;
        $bnmWS.Cells.Item($bnmRowNum, $BNM_CityCol            ) = $userObj.City;
        $bnmWS.Cells.Item($bnmRowNum, $BNM_STCol              ) = $userObj.State;
        $bnmWS.Cells.Item($bnmRowNum, $BNM_ZipCol             ) = $userObj.PostalCode;
    }

    function AddRowToEmergencyContacts($ecWS, [PSObject]$userObj, [int]$ecRowNum)
    {
        $ecWS.Cells.Item($ecRowNum, $EC_InformalLastNameCol      ) = $userObj.InformalName + " " +$userObj.LastName;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactNameCol           ) = $userObj.EmergencyContactName;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactRelationshipCol   ) = $userObj.EmergencyContactRelationship;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactPhoneCol          ) = $userObj.EmergencyContactPhone;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactEmailCol          ) = $userObj.EmergencyContactEmail;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactAltNameCol        ) = $userObj.EmergencyContactAltName;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactAltRelationshipCol) = $userObj.EmergencyContactAltRelationship;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactAltPhoneCol       ) = $userObj.EmergencyContactAltPhone;
        $ecWS.Cells.Item($ecRowNum, $EC_ContactAltEmailCol       ) = $userObj.EmergencyContactAltEmail;
    }

    function AddRowToRolebasedEmails($reWS, [PSObject]$reObj, [int]$reRowNum)
    {
        $reWS.Cells.Item($reRowNum, $RE_RolebasedEmailAddressCol     ) = $reObj.RolebasedEmailAddress;
        $reWS.Cells.Item($reRowNum, $RE_DisplayNameCol               ) = $reObj.DisplayName;
        $reWS.Cells.Item($reRowNum, $RE_EmailAddressBeingMonitoredCol) = $reObj.EmailAddressBeingMonitored;
        $reWS.Cells.Item($reRowNum, $RE_TypeEmailCol                 ) = $reObj.TypeEmail;
        $reWS.Cells.Item($reRowNum, $RE_DepartmentCol                ) = $reObj.Department;
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

    function CreateVolunteerYearString([datetime]$todaysDate, [PSCustomObject]$tenantObj)
    {
        # if start month is Jan then display only one year else return from and to year
        [string]$volunteerYearDateStr;
        if($tenantObj.VolunteerYearStartMonth -eq 1)
        {
            # return just one year
            $volunteerYearDateStr = $todaysDate.ToString("yyyy");
        }
        else
        {
            # return two years
            [DateTime]$thisDate = $todaysDate;
            [string]$rightNowYear = $thisDate.ToString("yyyy");
            [datetime]$startVolunteerYear = $tenantObj.VolunteerYearStartDate;
            [string]$volunteerYearStartStr = $startVolunteerYear.ToString("yyyy");
            if($startVolunteerYear -ge $thisDate)
            {
                $thisDate = $thisDate.AddYears(-1);
                $firstYearStr = $thisDate.ToString("yyyy");
                $volunteerYearDateStr = $firstYearStr + "-" + $volunteerYearStartStr;
            }
            else
            {
                $thisDate = $thisDate.AddYears(1);
                $nextYearStr = $thisDate.ToString("yyyy");
                $volunteerYearDateStr = $rightNowYear + "-" + $nextYearStr;
            }
        }
        return $volunteerYearDateStr;
    }

    function CreateMembershipTitle($memWS, [PSCustomObject]$titleObj)
    {
        [int]$memTitleRow = $titleObj.TitleRow;
        [int]$memTitleCol = $titleObj.TitleCol;
        [String]$volunteerYearStr = $titleObj.VolunteerYearStr;
        [String]$monthDayStr = $titleObj.MonthDayStr;
        [String]$lastUpdateDateTime = $titleObj.LastUpdateDateTime;
        # this goes in near the end because we do not want to AutoFit the first column to fit the title
        [string]$memSheetTitle = $tenantObj.ChapterRosterTitle + " -" + $volunteerYearStr + " Membership Roster - Rpt Date: " + $monthDayStr + " - Data as of: " + $lastUpdateDateTime;
        $memWS.Cells.Item($memTitleRow,$memTitleCol) = $memSheetTitle;
        $memWS.Cells.Item($memTitleRow,$memTitleCol).Font.Size = 18;
        $memWS.Cells.Item($memTitleRow,$memTitleCol).Font.Bold = $true;
        $memWS.Cells.Item($memTitleRow,$memTitleCol).font.Name = "Arial";
        $memWS.Cells.Item($memTitleRow,$memTitleCol).Font.ThemeFont = 1;
        $memWS.Cells.Item($memTitleRow,$memTitleCol).Font.ThemeColor = 4;
        $memWS.Cells.Item($memTitleRow,$memTitleCol).Font.ColorIndex = $titleColorIndex;
        $membershipTitleRange = $memWS.Range(("$M_LastNameLetter{0}"  -f $memTitleRow),("$M_JoinDateLetter{0}"  -f $memTitleRow));
        $membershipTitleRange.Select() | Out-Null;
        $membershipTitleRange.MergeCells = $true;
        $memWS.Cells($memTitleRow, $memTitleCol).HorizontalAlignment = $xlHAlignCenter; # center constant; used to center the title
    }

    function CreateBirthDaysThisMonthTitle($btmWS, [PSCustomObject]$titleObj)
    {
        [int]$btmTitleRow = $titleObj.TitleRow;
        [int]$btmTitleCol = $titleObj.TitleCol;
        [String]$volunteerYearStr = $titleObj.VolunteerYearStr;
        [String]$monthDayStr = $titleObj.MonthDayStr;
        [String]$lastUpdateDateTime = $titleObj.LastUpdateDateTime;
        # this goes in near the end because we do not want to AutoFit the first column to fit the title
        [string]$btmSheetTitle = $tenantObj.ChapterRosterTitle + " -" + $volunteerYearStr + " Birthdays This Month - Rpt Date: " + $monthDayStr + " - Data as of: " + $lastUpdateDateTime;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol) = $btmSheetTitle;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).Font.Size = 18;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).Font.Bold = $true;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).font.Name = "Arial";
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).Font.ThemeFont = 1;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).Font.ThemeColor = 4;
        $btmWS.Cells.Item($btmTitleRow,$btmTitleCol).Font.ColorIndex = $titleColorIndex;

        <#
        $btmTitleRange = $btmWS.Range(("$BTM_FirstLastNameLetter{0}"  -f $btmTitleRow),("$BTM_BirthDateLetter{0}"  -f $btmTitleRow));
        $btmTitleRange.Select() | Out-Null;
        $btmTitleRange.MergeCells = $true;
        $btmWS.Cells($btmTitleRow, $btmTitleCol).HorizontalAlignment = $xlHAlignCenter; # center constant; used to center the title
        #>
    }

    function CreateBirthDaysNextMonthTitle($bnmWS, [PSCustomObject]$titleObj)
    {
        [int]$bnmTitleRow = $titleObj.TitleRow;
        [int]$bnmTitleCol = $titleObj.TitleCol;
        [String]$volunteerYearStr = $titleObj.VolunteerYearStr;
        [String]$monthDayStr = $titleObj.MonthDayStr;
        [String]$lastUpdateDateTime = $titleObj.LastUpdateDateTime;
        # this goes in near the end because we do not want to AutoFit the first column to fit the title
        [string]$bnmSheetTitle = $tenantObj.ChapterRosterTitle + " -" + $volunteerYearStr + " Birthdays Next Month - Rpt Date: " + $monthDayStr + " - Data as of: " + $lastUpdateDateTime;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol) = $bnmSheetTitle;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).Font.Size = 18;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).Font.Bold = $true;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).font.Name = "Arial";
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).Font.ThemeFont = 1;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).Font.ThemeColor = 4;
        $bnmWS.Cells.Item($bnmTitleRow,$bnmTitleCol).Font.ColorIndex = $titleColorIndex;

        <#
        $bnmTitleRange = $bnmWS.Range(("$BNM_FirstLastNameLetter{0}"  -f $bnmTitleRow),("$BNM_BirthDateLetter{0}"  -f $bnmTitleRow));
        $bnmTitleRange.Select() | Out-Null;
        $bnmTitleRange.MergeCells = $true;
        $bnmWS.Cells($bnmTitleRow, $bnmTitleCol).HorizontalAlignment = $xlHAlignCenter; # center constant; used to center the title
        #>
    }

    function CreateEmergencyContactsTitle($ecWS, [PSCustomObject]$titleObj)
    {
        [int]$ecTitleRow = $titleObj.TitleRow;
        [int]$ecTitleCol = $titleObj.TitleCol;
        [String]$volunteerYearStr = $titleObj.VolunteerYearStr;
        [String]$monthDayStr = $titleObj.MonthDayStr;
        [String]$lastUpdateDateTime = $titleObj.LastUpdateDateTime;
        # this goes in near the end because we do not want to AutoFit the first column to fit the title
        [string]$ecSheetTitle = $tenantObj.ChapterRosterTitle + " -" + $volunteerYearStr + " Emergency Contacts - Rpt Date: " + $monthDayStr + " - Data as of: " + $lastUpdateDateTime;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol) = $ecSheetTitle;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).Font.Size = 16;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).Font.Bold = $true;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).font.Name = "Arial";
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).Font.ThemeFont = 1;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).Font.ThemeColor = 4;
        $ecWS.Cells.Item($ecTitleRow,$ecTitleCol).Font.ColorIndex = $titleColorIndex;

        <#
        $ecTitleRange = $ecWS.Range(("$EC_FirstLastNameLetter{0}"  -f $ecTitleRow),("$EC_AltEmailLetter{0}"  -f $ecTitleRow));
        $ecTitleRange.Select() | Out-Null;
        $ecTitleRange.MergeCells = $true;
        $ecWS.Cells($ecTitleRow, $ecTitleCol).HorizontalAlignment = $xlHAlignCenter; # center constant; used to center the title
        #>
    }

    function CreaterolebasedEmailsTitle($reWS, [PSCustomObject]$titleObj)
    {
        [int]$reTitleRow = $titleObj.TitleRow;
        [int]$reTitleCol = $titleObj.TitleCol;
        [String]$volunteerYearStr = $titleObj.VolunteerYearStr;
        [String]$monthDayStr = $titleObj.MonthDayStr;
        [String]$lastUpdateDateTime = $titleObj.LastUpdateDateTime;
        # this goes in near the end because we do not want to AutoFit the first column to fit the title
        [string]$reSheetTitle = $tenantObj.ChapterRosterTitle + " -" + $volunteerYearStr + " Role-based Emails - Rpt Date: " + $monthDayStr + " - Data as of: " + $lastUpdateDateTime;
        $reWS.Cells.Item($reTitleRow,$reTitleCol) = $reSheetTitle;
        $reWS.Cells.Item($reTitleRow,$reTitleCol).Font.Size = 18;
        $reWS.Cells.Item($reTitleRow,$reTitleCol).Font.Bold = $true;
        $reWS.Cells.Item($reTitleRow,$reTitleCol).font.Name = "Arial";
        $reWS.Cells.Item($reTitleRow,$reTitleCol).Font.ThemeFont = 1;
        $reWS.Cells.Item($reTitleRow,$reTitleCol).Font.ThemeColor = 4;
        $reWS.Cells.Item($reTitleRow,$reTitleCol).Font.ColorIndex = $titleColorIndex;

        <#
        $reTitleRange = $reWS.Range(("$RE_RolebasedEmailAddressLetter{0}"  -f $reTitleRow),("$RE_DepartmentLetter{0}"  -f $reTitleRow));
        $reTitleRange.Select() | Out-Null;
        $reTitleRange.MergeCells = $true;
        $reWS.Cells($ecTitleRow, $reTitleCol).HorizontalAlignment = $xlHAlignCenter; # center constant; used to center the title
        #>
    }

    function CreateMembershipHeaders($memWS, [int]$memHeaderRow)
    {
        $memWS.Cells.Item($memHeaderRow, $M_LastNameCol    ) = "Last Name";
        $memWS.Cells.Item($memHeaderRow, $M_InformalNameCol) = "Informal Name";
        $memWS.Cells.Item($memHeaderRow, $M_StatusCol      ) = "Status";
        $memWS.Cells.Item($memHeaderRow, $M_EmailAddressCol) = "E-mail Address";
        $memWS.Cells.Item($memHeaderRow, $M_HomeAddressCol ) = "Home Address";
        $memWS.Cells.Item($memHeaderRow, $M_CityCol        ) = "City";
        $memWS.Cells.Item($memHeaderRow, $M_STCol          ) = "State";
        $memWS.Cells.Item($memHeaderRow, $M_ZipCol         ) = "Zip Code";
        $memWS.Cells.Item($memHeaderRow, $M_HomePhoneCol   ) = "Home Phone    ";
        $memWS.Cells.Item($memHeaderRow, $M_CellCol        ) = "Mobile Phone  ";
        $memWS.Cells.Item($memHeaderRow, $M_BirthDateCol   ) = "B / Date";
        $memWS.Cells.Item($memHeaderRow, $M_SpouseCol      ) = "Spouse";
        $memWS.Cells.Item($memHeaderRow, $M_JoinDateCol    ) = "Joined Date";

        $membershipHeaderRange = $memWS.Range(("$M_LastNameLetter{0}"  -f $memHeaderRow),("$M_JoinDateLetter{0}"  -f $memHeaderRow));
        $membershipHeaderRange.Font.Bold = $true;
    }

    function CreateBirthdaysThisMonthHeaders($btmWS, [int]$btmHeaderRow)
    {
        $btmWS.Cells.Item($btmHeaderRow, $BTM_InformalLastNameCol) = "Name";
        $btmWS.Cells.Item($btmHeaderRow, $BTM_BirthDateCol       ) = "B / Date";

        $btmHeaderRange = $btmWS.Range(("$BTM_FirstLastNameLetter{0}"  -f $btmHeaderRow),("$BTM_BirthDateLetter{0}"  -f $btmHeaderRow));
        $btmHeaderRange.Font.Bold = $true;
    }

    function CreateBirthdaysNextMonthHeaders($bnmWS, [int]$bnmHeaderRow)
    {
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_InformalLastNameCol) = "Name";
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_BirthDateCol       ) = "B / Date";
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_HomeAddressCol     ) = "Home Address";
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_CityCol            ) = "City";
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_STCol              ) = "State";
        $bnmWS.Cells.Item($bnmHeaderRow, $BNM_ZipCol             ) = "Zip Code";

        $bnmHeaderRange = $bnmWS.Range(("$BNM_FirstLastNameLetter{0}"  -f $bnmHeaderRow),("$BNM_ZipLetter{0}"  -f $bnmHeaderRow));
        $bnmHeaderRange.Font.Bold = $true;
    }

    function CreateEmergencyContactsHeaders($ecWS, [int]$ecHeaderRow)
    {
        $ecWS.Cells.Item($ecHeaderRow, $EC_InformalLastNameCol      ) = "Volunteer Name";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactNameCol           ) = "Emergency Contact Name";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactRelationshipCol   ) = "Emergency Contact Relationship";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactPhoneCol          ) = "Emergency Contact Phone";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactEmailCol          ) = "Emergency Contact Email";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactAltNameCol        ) = "Emergency Contact Alt - Name";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactAltRelationshipCol) = "Emergency Contact Alt - Relationship";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactAltPhoneCol       ) = "Emergency Contact Alt - Phone";
        $ecWS.Cells.Item($ecHeaderRow, $EC_ContactAltEmailCol       ) = "Emergency Contact Alt - Email";

        $ecHeaderRange = $ecWS.Range(("$EC_FirstLastNameLetter{0}"  -f $ecHeaderRow),("$EC_AltEmailLetter{0}"  -f $ecHeaderRow));
        $ecHeaderRange.Font.Bold = $true;
    }

    function CreateRolebasedEmailHeaders($reWS, [int]$reHeaderRow)
    {
        # strings are padded right to display correctly in cell; longest value
        $reWS.Cells.Item($reHeaderRow, $RE_RolebasedEmailAddressCol      ) = "Role-base Email Address            ";
        $reWS.Cells.Item($reHeaderRow, $RE_DisplayNameCol                ) = "Display Name      ";
        $reWS.Cells.Item($reHeaderRow, $RE_EmailAddressBeingMonitoredCol ) = "Email Being Monitored / Used";
        $reWS.Cells.Item($reHeaderRow, $RE_TypeEmailCol                  ) = "TypeEmail  ";
        $reWS.Cells.Item($reHeaderRow, $RE_DepartmentCol                 ) = "Department                  ";

        $reHeaderRange = $reWS.Range(("$RE_RolebasedEmailAddressLetter{0}"  -f $reHeaderRow),("$RE_DepartmentLetter{0}"  -f $reHeaderRow));
        $reHeaderRange.Font.Bold = $true;
    }

    function CreateLoaBoxTitleAndHeaders($memWS, [int]$lRow)
    {
        [int]$tmpRow = $lRow; 
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol) = "On Leave of Absence";
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).font.Name = "Arial";
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).Font.Size = 14;
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).Font.Bold = $true;
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).Font.ThemeFont = 1;
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).Font.ThemeColor = 4;
        $memWS.Cells.Item($tmpRow,$LOA_LastNameCol).Font.ColorIndex = $titleColorIndex;

        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $LOA_LastNameCol ) = "Last Name";
        $memWS.Cells.Item($tmpRow, $LOA_FirstNameCol) = "First Name";
        $memWS.Cells.Item($tmpRow, $LOA_StartDateCol) = "Start Date";
        $memWS.Cells.Item($tmpRow, $LOA_StartDateCol).HorizontalAlignment = $xlHAlignRight;
        $memWS.Cells.Item($tmpRow, $LOA_EndDateCol  ) = "End Date";
        $memWS.Cells.Item($tmpRow, $LOA_EndDateCol  ).HorizontalAlignment = $xlHAlignRight;
        $loaHeaderRange = $memWS.Range(("$LOA_LastNameLetter{0}"  -f $tmpRow),("$LOA_EndDateLetter{0}"  -f $tmpRow));
        $loaHeaderRange.font.Name = "Arial";
        $loaHeaderRange.Font.Size = 12;
        $loaHeaderRange.Font.Bold = $true;
    }

    function CreateStatusBox($memWS, [int]$sRow, $cObj)
    {
        [int]$tmpRow = $sRow;
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol) = "Statistics";
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).Font.Size = 14;
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).Font.Bold = $true;
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).font.Name = "Arial";
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).Font.ThemeFont = 1;
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).Font.ThemeColor = 4;
        $memWS.Cells.Item($tmpRow,$Stat_NumberCol).Font.ColorIndex = $titleColorIndex;

        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Chapter Volunteers & Community Volunteers";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "On Leave of Absence";
        #$tmpRow += 1;
        #$memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Total Paid to National Count";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Community Volunteers";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "PALs";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Non-Voting";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Available to Vote";
        $tmpRow += 1;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Quorum (40% of Available)";
        $tmpRow += 1;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Resigned";
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_DescriptionCol ) = "Deceased";

        $statHeaderRange = $memWS.Range(("$Stat_DescriptionLetter{0}"  -f $sRow),("$Stat_DescriptionLetter{0}"  -f $tmpRow));
        $statHeaderRange.font.Name = "Arial";
        $statHeaderRange.Font.Size = 12;
        $statHeaderRange.Font.Bold = $true;

        $vCount = $cObj.AvaliableToVoteCount;
        $cObj.QuorumCount = ([Math]::Round(($vCount * .40) + 0.005, 0));

        $statStartRow = $sRow + 1
        $tmpRow = $statStartRow;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.TotalVolunteerCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.LoaCount;
        #$tmpRow += 1;
        #$memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.TotalPaidToNationalCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.CommunityVolunteerCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.PalCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.NonVotingCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.AvaliableToVoteCount;
        $tmpRow += 1;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.QuorumCount;
        $tmpRow += 1;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.ResignedCount;
        $tmpRow += 1;
        $memWS.Cells.Item($tmpRow, $Stat_NumberCol) = $cObj.DeceasedCount;

        $statRange = $memWS.Range(("$Stat_NumberLetter{0}"  -f $statStartRow),("$Stat_NumberLetter{0}"  -f $tmpRow));
        $statRange.font.Name = "Arial";
        $statRange.Font.Size = 12;

        $rightMostColumn = $Stat_NumberCol + 3
        $rightMostLetter = Convert-ToLetter $rightMostColumn;
        $borderRange = $memWS.Range(("$Stat_NumberLetter{0}"  -f $sRow),("$rightMostLetter{0}"  -f $tmpRow));
        $borderRange.BorderAround(1, $xlThick)  | Out-Null;
    }

    function CreateColorChart($memWS, $cRow)
    {
        [int]$tmpRow = $cRow;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "Color Chart";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Size = 14;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).font.Name = "Arial";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ThemeFont = 1;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ThemeColor = 4;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $titleColorIndex;

        $tmpRow ++;
        $startColorRow = $tmpRow;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "New Member Less Than a Year";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $newMemberFontColor;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;

        $tmpRow ++;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "Community Volunteer";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $communityVolenteerFontColor;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;

        $tmpRow ++;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "Leave of Absence";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $loaFontColor;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;

        $tmpRow ++;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "Deceased";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $deceasedFontColor;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;

        $tmpRow ++;
        $memWS.Cells.Item($tmpRow,$Color_StartCol) = "Resigned";
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.ColorIndex = $resignedFontColor;
        $memWS.Cells.Item($tmpRow,$Color_StartCol).Font.Bold = $true;

        $colorDataRange = $memWS.Range(("A{0}"  -f $startColorRow),("A{0}"  -f $tmpRow));
        $colorDataRange.Font.Size = 12;
        $colorDataRange.font.Name = "Arial";

        $colorBorderRange = $membershipWS.Range(("$Color_StartLetter{0}"  -f $cRow),("$Color_EndLetter{0}"  -f $tmpRow));
        $colorBorderRange.BorderAround(1, $xlThick) | Out-Null;
    }

    function AutoFitMembership($memWS, [int]$hRow, [int]$memEndRow)
    {
        $membershipDataRange = $memWS.Range(("$M_LastNameLetter{0}"  -f $hRow),("$M_JoinDateLetter{0}"  -f $memEndRow));
        $membershipDataRange.font.Name = "Arial";
        $membershipDataRange.Font.Size = 12;
        $memWS.columns.item($M_BirthDateLetter).NumberFormat = "MMM dd";
        $memWS.columns.item($M_JoinDateLetter).NumberFormat = "MM/dd/yyyy";
        $membershipDataRange.Columns.AutoFit() | Out-Null;
    }

    function AutoFitBirthdayThisMonth($btmWS, [int]$hRow, [int]$btmEndRow)
    {
        $birthdaysThisMonthDataRange = $btmWS.Range(("$BTM_FirstLastNameLetter{0}"  -f $hRow),("$BTM_BirthDateLetter{0}"  -f $btmEndRow));
        $birthdaysThisMonthDataRange.font.Name = "Arial";
        $birthdaysThisMonthDataRange.Font.Size = 12;
        $btmWS.columns.item($BTM_BirthDateLetter).NumberFormat = "MMM dd";
        $birthdaysThisMonthDataRange.Columns.AutoFit() | Out-Null;

        # now resort the whole Sheet
        [string]$birthdaySort = $BTM_BirthDateLetter + "2:" + $BTM_BirthDateLetter + $btmEndRow.ToString(); 
        # variable needed for UsedRangeSort
        $empty_Var = [System.Type]::Missing
        $sort_colBirthday = $btmWS.Range($birthdaySort)
        $btmWS.UsedRange.Sort($sort_colBirthday,1,$empty_Var,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1) | Out-Null;
    }

    function AutoFitBirthdayNextMonth($bnmWS, [int]$hRow, [int]$bnmEndRow)
    {
        $birthdaysNextMonthDataRange = $bnmWS.Range(("$BNM_FirstLastNameLetter{0}"  -f $hRow),("$BNM_ZipLetter{0}"  -f $bnmEndRow));
        $birthdaysNextMonthDataRange.font.Name = "Arial";
        $birthdaysNextMonthDataRange.Font.Size = 12;
        $bnmWS.columns.item($BNM_BirthDateLetter).NumberFormat = "MMM dd";
        $birthdaysNextMonthDataRange.Columns.AutoFit() | Out-Null;

        # now resort the whole Sheet
        [string]$birthdaySort = $BNM_BirthDateLetter + "2:" + $BNM_BirthDateLetter + $bnmEndRow.ToString(); 
        # variable needed for UsedRangeSort
        $empty_Var = [System.Type]::Missing
        $sort_colBirthday = $bnmWS.Range($birthdaySort)
        $bnmWS.UsedRange.Sort($sort_colBirthday,1,$empty_Var,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1) | Out-Null;
    }

    function AutoFitEmergencyContacts($ecWS, [int]$hRow, [int]$ecEndRow)
    {
        $emergencyContactsDataRange = $ecWS.Range(("$EC_FirstLastNameLetter{0}"  -f $hRow),("$EC_AltEmailLetter{0}"  -f $ecEndRow));
        $emergencyContactsDataRange.font.Name = "Arial";
        $emergencyContactsDataRange.Font.Size = 12;
        $ecWS.columns.item($BNM_BirthDateLetter).NumberFormat = "MMM dd";
        $emergencyContactsDataRange.Columns.AutoFit() | Out-Null;

    }

    function AutoFitRolebasedEmails($reWS, [int]$hRow, [int]$reEndRow)
    {
        $rolebasedEmailsDataRange = $reWS.Range(("$RE_RolebasedEmailAddressLetter{0}"  -f $hRow),("$RE_DepartmentLetter{0}"  -f $reEndRow));
        $rolebasedEmailsDataRange.font.Name = "Arial";
        $rolebasedEmailsDataRange.Font.Size = 12;
        $rolebasedEmailsDataRange.Columns.AutoFit() | Out-Null;

    }

    function IsANewMember($userObj, [PSCustomObject]$tenantObj)
    {
        $isNewMember = $false;
        [datetime]$joinDate = ConverToPSDateTime -dateToConvert $userObj.JoinDate;
        [datetime]$curDate = Get-Date;
        [datetime]$yearAgoDate = $curDate.AddYears(-1);
        if($null -ne $joinDate)
        {
            if($joinDate -gt $yearAgoDate)
            {
                $isNewMember = $true;
            }
        }
        return $isNewMember;
    }

    function ColorMembershipRow($memWS, $memRowNum, $fontColor)
    {
        $membershipColorRange = $memWS.Range(("$M_LastNameLetter{0}"  -f $memRowNum),("$M_JoinDateLetter{0}"  -f $memRowNum));
        $membershipColorRange.Font.ColorIndex = $fontColor;
        $membershipColorRange.Font.Bold = $true;
    }

    function ColorRolebasedEmailsRow($reWS, $reRowNum, $fontColor)
    {
        $rolebasedEmailsColorRange = $reWS.Range(("$RE_RolebasedEmailAddressLetter{0}"  -f $reRowNum),("$RE_DepartmentLetter{0}"  -f $reRowNum));
        $rolebasedEmailsColorRange.Font.ColorIndex = $fontColor;
        $rolebasedEmailsColorRange.Font.Bold = $true;
    }

    function Release_Ref ($ref)
    {
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    function LogIncomingParameters()
    {
        if($null -eq $masterLogFilePathAndName) {Write-Output("masterLogFile parameter is null"); exit;}
        #region log startup info
        $logMessage = "Start up UpdateNALSchemaFrom_chDocument.ps1 with args:";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "workbookFilePathAndName: " + $workbookFilePathAndName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "pdfFilePathAndName: " + $pdfFilePathAndName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "rosterMembershipSheetName: " + $rosterMembershipSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "rosterBirthdaysThisMonthSheetName: " + $rosterBirthdaysThisMonthSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "rosterBirthdaysNextMonthSheetName: " + $rosterBirthdaysNextMonthSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "rosterEmergencyContactsSheetName: " + $rosterEmergencyContactsSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "rosterRolebasedEmailsSheetName: " + $rosterRolebasedEmailsSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "m365SchemaFilePathAndName: " + $m365SchemaFilePathAndName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "m365UsersSheetName: " + $m365UsersSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "m365roleBasedEmailAddressesSheetName: " + $m365roleBasedEmailAddressesSheetName;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "protectWorkBookAndSheets: " + $protectWorkBookAndSheets;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "typeOfReport: " + $typeOfReport;
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        #endregion log startup info
    }



    #endregion function definitions

    #start working here
    $thisScriptName = $MyInvocation.MyCommand.Name
    $logMessage = "Starting " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    [string]$tenantDomain =$tenantObj.DomainName;

    LogIncomingParameters;

    #region setup Excel application object
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

    #region Excel enumeration constant definitions
    $empty_Var = [System.Type]::Missing
    $true_Var = [System.Type]::true

    # alignment
    $xlHAlignCenter = -4108;
    $xlHAlignRight  = -4152;

    # border thickness
    #$xlHairline	= 1;     # Hairline (thinnest border).
    #$xlMedium	= -4138; # Medium.
    $xlThick	= 4;     # Thick (widest border).
    #$xlThin	    = 2;     # Thin

    # print constants
    #$xlPortrait = 1
    $xlLandscape = 2
    $xlPrintNoComments = -4142
    #$xlPaperLetter = 1
    #$xlPaperLedger=4
    #$xlPaperLegal = 5
    #$xlPaperFolio=14
    #$xlPaper11x17=17
    #$xlDownThenOver = 1
    $xlAutomatic = -4105
    #endregion Excel enumeration constant definitions

    #region ChapterRoster.xlsx column constant definitions
    [int]$M_LastNameCol      = 01; # A
    [int]$M_InformalNameCol  = 02; # B
    [int]$M_StatusCol        = 03; # C
    [int]$M_EmailAddressCol  = 04; # D
    [int]$M_HomeAddressCol   = 05; # E
    [int]$M_CityCol          = 06; # F
    [int]$M_STCol            = 07; # G
    [int]$M_ZipCol           = 08; # H
    [int]$M_HomePhoneCol     = 09; # I
    [int]$M_CellCol          = 10; # J
    [int]$M_BirthDateCol     = 11; # K
    [int]$M_SpouseCol        = 12; # L
    [int]$M_JoinDateCol      = 13; # M

    [string]$M_LastNameLetter     = Convert-ToLetter $M_LastNameCol;
    #[string]$M_StatusLetter       = Convert-ToLetter $M_StatusCol;
    #[string]$M_EmailAddressLetter = Convert-ToLetter $M_EmailAddressCol;
    [string]$M_BirthDateLetter    = Convert-ToLetter $M_BirthDateCol;
    [string]$M_JoinDateLetter     = Convert-ToLetter $M_JoinDateCol;

    [int]$BTM_InformalLastNameCol = 01; # A
    [int]$BTM_BirthDateCol        = 02; # B

    [string]$BTM_FirstLastNameLetter = Convert-ToLetter $BTM_InformalLastNameCol;
    [string]$BTM_BirthDateLetter     = Convert-ToLetter $BTM_BirthDateCol;

    [int]$BNM_InformalLastNameCol = 01; # A
    [int]$BNM_BirthDateCol        = 02; # B
    [int]$BNM_HomeAddressCol      = 04; # D
    [int]$BNM_CityCol             = 05; # E
    [int]$BNM_STCol               = 06; # F
    [int]$BNM_ZipCol              = 07; # G

    [string]$BNM_FirstLastNameLetter = Convert-ToLetter $BNM_InformalLastNameCol;
    [string]$BNM_BirthDateLetter     = Convert-ToLetter $BNM_BirthDateCol;
    [string]$BNM_ZipLetter           = Convert-ToLetter $BNM_ZipCol;

    [int]$EC_InformalLastNameCol       = 01; # A
    [int]$EC_ContactNameCol            = 03; # C
    [int]$EC_ContactRelationshipCol    = 04; # D
    [int]$EC_ContactPhoneCol           = 05; # E
    [int]$EC_ContactEmailCol           = 06; # F
    [int]$EC_ContactAltNameCol         = 07; # G
    [int]$EC_ContactAltRelationshipCol = 08; # H
    [int]$EC_ContactAltPhoneCol        = 09; # I
    [int]$EC_ContactAltEmailCol        = 10; # J

    [string]$EC_FirstLastNameLetter    = Convert-ToLetter $EC_InformalLastNameCol;
    #[string]$EC_ContactAltNameLetter   = Convert-ToLetter $EC_ContactAltEmailCol;
    [string]$EC_AltEmailLetter         = Convert-ToLetter $EC_ContactAltEmailCol;

    [int]$RE_RolebasedEmailAddressCol        = 01; # A
    [int]$RE_DisplayNameCol                  = 02; # B
    [int]$RE_EmailAddressBeingMonitoredCol   = 03; # C
    [int]$RE_TypeEmailCol                    = 04; # D
    [int]$RE_DepartmentCol                   = 05; # E

    [string]$RE_RolebasedEmailAddressLetter  = Convert-ToLetter $RE_RolebasedEmailAddressCol;
    [string]$RE_DepartmentLetter             = Convert-ToLetter $RE_DepartmentCol;

    [int]$LOA_LastNameCol  = 01; # A
    [int]$LOA_FirstNameCol = 02; # B
    [int]$LOA_StartDateCol = 04; # E
    [int]$LOA_EndDateCol   = 05; # F

    [string]$LOA_LastNameLetter = Convert-ToLetter $LOA_LastNameCol;
    [string]$LOA_EndDateLetter  = Convert-ToLetter $LOA_EndDateCol;

    [int]$Stat_NumberCol      = 01;
    [int]$Stat_DescriptionCol = 02;

    [string]$Stat_NumberLetter      = Convert-ToLetter $Stat_NumberCol;
    [string]$Stat_DescriptionLetter = Convert-ToLetter $Stat_DescriptionCol;

    [int]$Color_StartCol = 06;
    [int]$Color_EndCol   = $Color_StartCol + 2;

    [string]$Color_StartLetter = Convert-ToLetter $Color_StartCol;
    [string]$Color_EndLetter   = Convert-ToLetter $Color_EndCol;



    #endregion ChapterRoster.xlsx column constant definitions

    #region color constant definitions
    $titleColorIndex = 41;
    $newMemberFontColor = 10;
    $communityVolenteerFontColor = 33;
    $loaFontColor = 7;
    $deceasedFontColor = 39;
    $resignedFontColor = 3;
    $notMonitoredEmailAddressColor = 3;
    #endregion color constant definitions

    #region APStatus variable definitions
    #[string]$listName = "APStatuses";
    [string]$ProcessName = $typeOfReport;
    [string]$ProcessCategory = "CHWorkFlow";
    [DateTime]$ProcessStartDate = Get-date;
    #[DateTime]$ProcessStartDate = (Get-date).ToUniversalTime();
    [DateTime]$ProcessStopDateBegins = (Get-date("1/1/1900 00:01")).ToUniversalTime();
    [DateTime]$SProcessStopDateFinished =( Get-date("1/1/2099 00:01")).ToUniversalTime();
    [string]$ProcessStatusStart = "Started";
    [string]$ProcessStatusFinish = "Successful";
    [string]$ProcessProgressStart = "In-progress";
    [string]$ProcessProgressFinish = "Completed";
    #endregion

    #region updating APStatusTable with Start values
    .\M365SharePoint\InsertOrUpdateTenantStatusRptList.ps1 -tenantCredentials $tenantCredentials `
                                                           -tenantAbbreviation $tenantAbbreviation `
                                                           -tenantDomain $tenantDomain `
                                                           -ProcessName $ProcessName `
                                                           -ProcessCategory $ProcessCategory `
                                                           -StartDate $ProcessStartDate `
                                                           -StopDate $ProcessStopDateBegins `
                                                           -ProcessStatus $ProcessStatusStart `
                                                           -ProcessProgress $ProcessProgressStart `
                                                           -masterLogFilePathAndName $masterLogFilePathAndName | Out-Null;
    #endregion
    
    #region load Assistant League Chapter Schema user objects and Excel workbook and sheet objects
    $logMessage = "Loading User objects and Excel objects";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    [PSObject]$alcsUserObjs = Import-Excel -Path $m365SchemaFilePathAndName -WorksheetName $m365UsersSheetName -StartRow 1 -DataOnly;
    [PSObject]$rolebasedEmailObjs = Import-Excel -Path $m365SchemaFilePathAndName -WorksheetName $m365roleBasedEmailAddressesSheetName -StartRow 1 -DataOnly;

    #UpdateStatus
    [string]$updateStatusSheetName = "UpdateStatus";
    [PSObject]$schemaUpdateStatusObj = Import-Excel -Path $m365SchemaFilePathAndName -WorksheetName $updateStatusSheetName -StartRow 1 -DataOnly;

    #region add and name worksheets
    $wb = $xl.Workbooks.Add();
    $wb.Worksheets.Add() | Out-Null;
    $wb.Worksheets.Add() | Out-Null;
    $wb.Worksheets.Add() | Out-Null;
    $wb.Worksheets.Add() | Out-Null;
    $membershipWS = $wb.Worksheets.Item(1);
    $membershipWS.Name = $rosterMembershipSheetName;
    $birthdaysThisMonthWS = $wb.Worksheets.Item(2);
    $birthdaysThisMonthWS.Name = $rosterBirthdaysThisMonthSheetName;
    $birthdaysNextMonthWS = $wb.Worksheets.Item(3);
    $birthdaysNextMonthWS.Name = $rosterBirthdaysNextMonthSheetName;
    $emergencyContactsWS = $wb.Worksheets.Item(4);
    $emergencyContactsWS.Name = $rosterEmergencyContactsSheetName;
    $rolebasedEmailsWS = $wb.Worksheets.Item(5);
    $rolebasedEmailsWS.Name = $rosterRolebasedEmailsSheetName;
    #endregion

    #connect to SharePoint On-line
    <#
    $logMessage = "Connecting to SharePoint On-line";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    Connect-MsolService -Credential $tenantCredentials;
    #>

    #region update roster sheets Data

     [PSCustomObject]$countsObj = [PSCustomObject][ordered]@{
                                                             TotalVolunteerCount      = 0;
                                                             AvaliableToVoteCount     = 0;
                                                             TotalPaidToNationalCount = 0;
                                                             ResignedCount            = 0;
                                                             DeceasedCount            = 0;
                                                             CommunityVolunteerCount  = 0
                                                             PalCount                 = 0;
                                                             NonVotingCount           = 0;
                                                             LoaCount                 = 0;
                                                             QuorumCount              = 0;
                                                            }

    # define variables
    #region
    [int]$membershipTitleRow  = 1;
    [int]$membershipTitleCol  = 1;
    [int]$membershipHeaderRow = 2;
    #[int]$membershipStartRow  = $membershipHeaderRow + 1;  # start of membership data; used in range for auto sizing
    [int]$membershipRow       = $membershipHeaderRow;       # starts at 3 because of increment; Leaves room for title and headers
    [int]$membershipEndRow    = -1; # end of membership data; Used in range for auto sizing; Value assigned later

    [int]$birthdaysThisMonthTitleRow   = 1;
    [int]$birthdaysThisMonthTitleCol   = 1;
    [int]$birthdaysThisMonthHeaderRow  = 2
    #[int]$birthdaysThisMonthStartRow   = $birthdaysThisMonthHeaderRow  + 1;  # start of birthdayThisMonth data; used in range for auto sizing
    [int]$birthdaysThisMonthRow        = $birthdaysThisMonthHeaderRow; # starts at 3 because of increment; Leaves room for title and headers
    #[int]$birthdaysThisMonthEndRow     = -1; # end of birthdayThisMonth data; Used in range for auto sizing; Value assigned later

    [int]$birthdaysNextMonthTitleRow  = 1;
    [int]$birthdaysNextMonthTitleCol  = 1;
    [int]$birthdaysNextMonthHeaderRow = 2;
    #[int]$birthdaysNextMonthStartRow  = $birthdaysNextMonthHeaderRow  + 1;  # start of birthdayNextMonth data; used in range for auto sizing
    [int]$birthdaysNextMonthRow       = $birthdaysNextMonthHeaderRow; # starts at 3 because of increment; Leaves room for title and headers
    #[int]$birthdaysNextMonthEndRow    = -1; # end of birthdayNextMonth data; Used in range for auto sizing; Value assigned later

    [int]$emergencyContactsTitleRow  = 1;
    [int]$emergencyContactsTitleCol  = 1;
    [int]$emergencyContactsHeaderRow = 2
    #[int]$emergencyContactsStartRow  = $emergencyContactsHeaderRow  + 1;  # start of emergencyContact data; used in range for auto sizing
    [int]$emergencyContactsRow       = $emergencyContactsHeaderRow;  # starts at 3 because of increment; Leaves room for title and headers
    [int]$emergencyContactsEndRow    = -1; # end of emergencyContact data; Used in range for auto sizing; Value assigned later

    [int]$rolebasedEmailsTitleRow  = 1;
    [int]$rolebasedEmailsTitleCol  = 1;
    [int]$rolebasedEmailsHeaderRow = 2
    #[int]$rolebasedEmailsStartRow  = $rolebasedEmailsHeaderRow  + 1;  # start of role-based email data; used in range for auto sizing
    [int]$rolebasedEmailsRow       = $rolebasedEmailsHeaderRow;  # starts at 3 because of increment; Leaves room for title and headers
    [int]$rolebasedEmailsEndRow    = -1; # end of emergencyContact data; Used in range for auto sizing; Value assigned later

    [bool]$isBirthDayThisMonth    = $false;
    [bool]$isBirthDayNextMonth    = $false;
    #endregion define variables

    #Create header row at top of all 4 worksheets
    #region
    $logMessage = "Creating header row at top of all 4 worksheets";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    CreateMembershipHeaders         -memWS $membershipWS         -memHeaderRow $membershipHeaderRow;
    CreateBirthdaysThisMonthHeaders -btmWS $birthdaysThisMonthWS -btmHeaderRow $birthdaysThisMonthHeaderRow;
    CreateBirthdaysNextMonthHeaders -bnmWS $birthdaysNextMonthWS -bnmHeaderRow $birthdaysNextMonthHeaderRow;
    CreateEmergencyContactsHeaders  -ecWS  $emergencyContactsWS  -ecHeaderRow  $emergencyContactsHeaderRow;
    CreateRolebasedEmailHeaders     -reWS  $rolebasedEmailsWS    -reHeaderRow  $rolebasedEmailsHeaderRow;
    #endregion Create header row at top of all 4 worksheets

    # loop to fill-in data for membership, birthdaysThisMonth, birthdaysNextMonth, and emergencyContacts
    foreach($alcsUserObj in $alcsUserObjs)
    {
        [string]$alStatus = $alcsUserObj.ALStatus;
        $isBirthDayThisMonth = IsVolunteersBirthdayThisMonth -userObj $alcsUserObj;
        $isBirthDayNextMonth = IsVolunteersBirthdayNextMonth -userObj $alcsUserObj;

        switch ($alStatus)
        {
            "V"
            {
                $countsObj.TotalVolunteerCount++;
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.AvaliableToVoteCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                $isANewMember = IsANewMember -userObj $alcsUserObj -tenantObj $tenantObj;
                if($isANewMember)
                {
                    ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $newMemberFontColor;
                }
                if($isBirthDayThisMonth)
                {
                    $birthdaysThisMonthRow++;
                    AddRowToBirthdaysThisMonth -btmWS $birthdaysThisMonthWS -userObj $alcsUserObj -btmRowNum $birthdaysThisMonthRow;
                }
                if($isBirthDayNextMonth)
                {
                    $birthdaysNextMonthRow++;
                    AddRowToBirthdaysNextMonth -bnmWS $birthdaysNextMonthWS -userObj $alcsUserObj -bnmRowNum $birthdaysNextMonthRow;
                }
                $emergencyContactsRow++;
                AddRowToEmergencyContacts -ecWS $emergencyContactsWS -userObj $alcsUserObj -ecRowNum $emergencyContactsRow;
                break;
            }

            "CV"
            {
                $countsObj.TotalVolunteerCount++;
                $countsObj.CommunityVolunteerCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $communityVolenteerFontColor;
                if($isBirthDayThisMonth)
                {
                    $birthdaysThisMonthRow++;
                    AddRowToBirthdaysThisMonth -btmWS $birthdaysThisMonthWS -userObj $alcsUserObj -btmRowNum $birthdaysThisMonthRow;
                }
                if($isBirthDayNextMonth)
                {
                    $birthdaysNextMonthRow++;
                    AddRowToBirthdaysNextMonth -bnmWS $birthdaysNextMonthWS -userObj $alcsUserObj -bnmRowNum $birthdaysNextMonthRow;
                }
                $emergencyContactsRow++;
                AddRowToEmergencyContacts -ecWS $emergencyContactsWS -userObj $alcsUserObj -ecRowNum $emergencyContactsRow;
                break;
            }

            "PAL"
            {
                $countsObj.TotalVolunteerCount++;
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.PalCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                $isANewMember = IsANewMember -userObj $alcsUserObj -tenantObj $tenantObj;
                if($isANewMember)
                {
                    ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $newMemberFontColor;
                }
                if($isBirthDayThisMonth)
                {
                    $birthdaysThisMonthRow++;
                    AddRowToBirthdaysThisMonth -btmWS $birthdaysThisMonthWS -userObj $alcsUserObj -btmRowNum $birthdaysThisMonthRow;
                }
                if($isBirthDayNextMonth)
                {
                    $birthdaysNextMonthRow++;
                    AddRowToBirthdaysNextMonth -bnmWS $birthdaysNextMonthWS -userObj $alcsUserObj -bnmRowNum $birthdaysNextMonthRow;
                }
                $emergencyContactsRow++;
                AddRowToEmergencyContacts -ecWS $emergencyContactsWS -userObj $alcsUserObj -ecRowNum $emergencyContactsRow;
                break;
            }

            "NV"
            {
                $countsObj.TotalVolunteerCount++;
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.NonVotingCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                $isANewMember = IsANewMember -userObj $alcsUserObj -tenantObj $tenantObj;
                if($isANewMember)
                {
                    ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $newMemberFontColor;
                }
                if($isBirthDayThisMonth)
                {
                    $birthdaysThisMonthRow++;
                    AddRowToBirthdaysThisMonth -btmWS $birthdaysThisMonthWS -userObj $alcsUserObj -btmRowNum $birthdaysThisMonthRow;
                }
                if($isBirthDayNextMonth)
                {
                    $birthdaysNextMonthRow++;
                    AddRowToBirthdaysNextMonth -bnmWS $birthdaysNextMonthWS -userObj $alcsUserObj -bnmRowNum $birthdaysNextMonthRow;
                }
                $emergencyContactsRow++;
                AddRowToEmergencyContacts -ecWS $emergencyContactsWS -userObj $alcsUserObj -ecRowNum $emergencyContactsRow;
                break;
            }

            "LOA"
            {
                $countsObj.TotalVolunteerCount++;
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.LoaCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $loaFontColor;
                if($isBirthDayThisMonth)
                {
                    $birthdaysThisMonthRow++;
                    AddRowToBirthdaysThisMonth -btmWS $birthdaysThisMonthWS -userObj $alcsUserObj -btmRowNum $birthdaysThisMonthRow;
                }
                if($isBirthDayNextMonth)
                {
                    $birthdaysNextMonthRow++;
                    AddRowToBirthdaysNextMonth -bnmWS $birthdaysNextMonthWS -userObj $alcsUserObj -bnmRowNum $birthdaysNextMonthRow;
                }
                $emergencyContactsRow++;
                AddRowToEmergencyContacts -ecWS $emergencyContactsWS -userObj $alcsUserObj -ecRowNum $emergencyContactsRow;
                break;
            }

            "D"
            {
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.DeceasedCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $deceasedFontColor;
                break;
            }

            "R"
            {
                $countsObj.TotalPaidToNationalCount++;
                $countsObj.ResignedCount++;
                $membershipRow++;
                AddRowToMembership -memWS $membershipWS -userObj $alcsUserObj -memRowNum $membershipRow;
                ColorMembershipRow -memWS $membershipWS -memRowNum $membershipRow -fontColor $resignedFontColor;
                break;
            }

            default
            {
                $logMessage = "Unknown AL Status: " + $alStatus + "  ***************************";
                .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                break;
            }
        }
    }

    # these are assigned to the last value at the end of the loop
    $membershipEndRow = $membershipRow;
    $birthdayThisMonthEndRow = $birthdaysThisMonthRow;
    $birthdayNextMonthEndRow = $birthdaysNextMonthRow;
    $emergencyContactsEndRow = $emergencyContactsRow;
    #endregion update roster sheets Data

    $logMessage = "Auto-Fitting Membership Sheet";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    AutoFitMembership -memWS $membershipWS  -hRow $membershipHeaderRow  -memEndRow $membershipEndRow;
    
    #region add volunteers on LOA
    [int]$loaTitleRow = $membershipEndRow + 3;
    [int]$loaHeaderRow = $loaTitleRow + 1;
    [int]$loaRow = $loaHeaderRow;
    [int]$loaStartRow = $loaHeaderRow + 1;

    CreateLoaBoxTitleAndHeaders -memWS $membershipWS -lRow $loaTitleRow;

    foreach($alcsUserObj in $alcsUserObjs)
    {
        if($alcsUserObj.ALStatus -eq "LOA")
        {
            $loaRow += 1;
            $startDate = ConverToPSDateTime $alcsUserObj.LOAStartDate;
            $endDate = ConverToPSDateTime -dateToConvert $alcsUserObj.LOAEndDate;
            $membershipWS.Cells.Item($loaRow, $LOA_LastNameCol ) = $alcsUserObj.LastName;
            $membershipWS.Cells.Item($loaRow, $LOA_FirstNameCol) = $alcsUserObj.FirstName;
            $membershipWS.Cells.Item($loaRow, $LOA_StartDateCol) = $startDate;
            $membershipWS.Cells.Item($loaRow, $LOA_EndDateCol  ) = $endDate;
        }
    }
    $loaEndRow = $loaRow;
    $loaBorderRange = $membershipWS.Range(("$LOA_LastNameLetter{0}"  -f $loaStartRow),("$LOA_EndDateLetter{0}"  -f $loaEndRow));
    $loaBorderRange.font.Name = "Arial";
    $loaBorderRange.Font.Size = 12;

    $borderRange = $membershipWS.Range(("$LOA_LastNameLetter{0}"  -f $loaTitleRow),("$LOA_EndDateLetter{0}"  -f $loaEndRow));
    $borderRange.BorderAround(1, $xlThick) | Out-Null;

    #endregion add volunteers on LOA

    $logMessage = "Creating Status Box";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    $statusHeaderRow = $loaRow + 3;
    CreateStatusBox -memWS $membershipWS -sRow $statusHeaderRow -cObj $countsObj;

    $logMessage = "Creating Color Chart";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    $colorChartHeaderRow = $statusHeaderRow;
    CreateColorChart -memWS $membershipWS -cRow $colorChartHeaderRow; 

    #
    foreach($rolebasedEmailObj in $rolebasedEmailObjs)
    {
        if($rolebasedEmailObj.TypeEmail -ne 'Automation')
        {
            [DateTime]$currentDate = Get-Date;
            [DateTime]$activeDate = $rolebasedEmailObj.ActiveDate;
            [DateTime]$inactiveDate = $rolebasedEmailObj.InactiveDate;
            if($activeDate -le $currentDate -and $inactiveDate -ge $currentDate)
            {
                $rolebasedEmailsRow++;
                AddRowToRolebasedEmails -reWS $rolebasedEmailsWS -reObj $rolebasedEmailObj -reRowNum $rolebasedEmailsRow;
                if(!$rolebasedEmailObj.EmailAddressBeingMonitored)
                {
                    ColorRolebasedEmailsRow -reWS $rolebasedEmailsWS -reRowNum $rolebasedEmailsRow -fontColor $notMonitoredEmailAddressColor;
                }
            }
        }
    }
    $rolebasedEmailsEndRow = $rolebasedEmailsRow;

    $volunteerYearStr = CreateVolunteerYearString -todaysDate (Get-Date) -tenantObj $tenantObj;
    $monthDayStr   = (Get-Date).ToString("MMM d");

    $logMessage = "Creating titles and Freezing Panes";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    [PSCustomObject]$titleObj = [PSCustomObject][ordered]@{
        TitleRow           = $membershipTitleRow;
        TitleCol           = $membershipTitleCol;
        ChapterRosterTitle = $tenantObj.ChapterRosterTitle;
        VolunteerYearStr   = $volunteerYearStr;
        MonthDayStr        = $monthDayStr;
        LastUpdateDateTime = $schemaUpdateStatusObj.LastUpdateDateTime.ToString("MM-dd-yy HH:mm");
       }

    CreateMembershipTitle -memWS $membershipWS -titleObj $titleObj;
    $membershipWS.Select() | Out-Null;
    $membershipWS.Activate();
    $membershipWS.Application.ActiveWindow.FreezePanes = $false;
    $membershipWS.Application.ActiveWindow.SplitRow = 2;
    $membershipWS.Application.ActiveWindow.FreezePanes = $true;

    <#
    if($typeOfReport -eq  'PDFRoster' -or $typeOfReport -eq  'PDFEmergencyContacts')
    {
        $xx = 1;
    }
    #>
    
    AutoFitBirthdayThisMonth -btmWS $birthdaysThisMonthWS -hRow $birthdaysThisMonthHeaderRow -btmEndRow $birthdayThisMonthEndRow;
    $titleObj.TitleRow = $birthdaysThisMonthTitleRow;
    $titleObj.TitleRow = $birthdaysThisMonthTitleCol;
    CreateBirthDaysThisMonthTitle -btmWS $birthdaysThisMonthWS -titleObj $titleObj;
    $birthdaysThisMonthWS.Select() | Out-Null;
    $birthdaysThisMonthWS.Activate();
    $birthdaysThisMonthWS.Application.ActiveWindow.FreezePanes = $false;
    $birthdaysThisMonthWS.Application.ActiveWindow.SplitRow = 2;
    $birthdaysThisMonthWS.Application.ActiveWindow.FreezePanes = $true;

    AutoFitBirthdayNextMonth -bnmWS $birthdaysNextMonthWS -hRow $birthdaysNextMonthHeaderRow -bnmEndRow $birthdayNextMonthEndRow;
    $titleObj.TitleRow = $birthdaysNextMonthTitleRow;
    $titleObj.TitleRow = $birthdaysNextMonthTitleCol;
    CreateBirthDaysNextMonthTitle -bnmWS $birthdaysNextMonthWS -titleObj $titleObj;
    $birthdaysNextMonthWS.Select() | Out-Null;
    $birthdaysNextMonthWS.Activate();
    $birthdaysNextMonthWS.Application.ActiveWindow.FreezePanes = $false;
    $birthdaysNextMonthWS.Application.ActiveWindow.SplitRow = 2;
    $birthdaysNextMonthWS.Application.ActiveWindow.FreezePanes = $true;

    AutoFitEmergencyContacts -ecWS $emergencyContactsWS -hRow $emergencyContactsHeaderRow -ecEndRow $emergencyContactsEndRow;
    $titleObj.TitleRow = $emergencyContactsTitleRow;
    $titleObj.TitleRow = $emergencyContactsTitleCol;
    CreateEmergencyContactsTitle -ecWS  $emergencyContactsWS -titleObj $titleObj;
    $emergencyContactsWS.Select() | Out-Null;
    $emergencyContactsWS.Activate();
    $emergencyContactsWS.Application.ActiveWindow.FreezePanes = $false;
    $emergencyContactsWS.Application.ActiveWindow.SplitRow = 2;
    $emergencyContactsWS.Application.ActiveWindow.FreezePanes = $true;

    AutoFitRoleBasedEmails -reWS $rolebasedEmailsWS -hRow $rolebasedEmailsHeaderRow -reEndRow $rolebasedEmailsEndRow;
    $titleObj.TitleRow = $rolebasedEmailsTitleRow;
    $titleObj.TitleRow = $rolebasedEmailsTitleCol;
    CreateRoleBasedEmailsTitle -reWS  $rolebasedEmailsWS -titleObj $titleObj;
    $rolebasedEmailsWS.Select() | Out-Null;
    $rolebasedEmailsWS.Activate();
    $rolebasedEmailsWS.Application.ActiveWindow.FreezePanes = $false;
    $rolebasedEmailsWS.Application.ActiveWindow.SplitRow = 2;
    $rolebasedEmailsWS.Application.ActiveWindow.FreezePanes = $true;
    #endregion

    #region set final point of focus when Excel file is opened
    $membershipWS.Select() | Out-Null;
    $membershipWS.Activate();
    $finalSelectCell = $membershipWS.Range("A3:A3");
    $finalSelectCell.Select() | Out-Null;
    #endregion set final point of focus when Excel file is opened

    if($typeOfReport -eq  'PDFRoster')
    {
        $logMessage = "Creating PDFRoster workbook";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        #region delete all the sheets except for Membership
        $wb.Worksheets.Item(5).Delete()
        $wb.Worksheets.Item(4).Delete()
        $wb.Worksheets.Item(3).Delete()
        $wb.Worksheets.Item(2).Delete()
        $ws = $wb.Worksheets.Item(1);
        $ws.Select() | Out-Null;
        $ws.Activate();
        $ws.PageSetup.Orientation = $xlLandscape;
        $ws.PageSetup.PrintTitleRows = '$1:$2';
        # https://www.litigationsupporttipofthenight.com/single-post/2016/06/20/powershell-script-to-print-excel-files-landscaped-to-pdfs
        $ws.PageSetup.LeftMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.RightMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.TopMargin = $xl.InchesToPoints(0.5);
        $ws.PageSetup.BottomMargin = $xl.InchesToPoints(0.5);
        $ws.PageSetup.HeaderMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.FooterMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.PrintGridlines = $true;
        $ws.PageSetup.PrintComments = $xlPrintNoComments;
        $ws.PageSetup.FirstPageNumber = $xlAutomatic;
        $ws.PageSetup.Order = 1;
        $ws.PageSetup.Zoom = $false;
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = 9999;

        
        $range = $ws.usedRange;
        $r = $range.rows.count;
        $c = $range.Columns.count;
        $S = $ws.Cells.Item($r,$c);
        $logMessage = "Column Count: " + $c.ToString();
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $S = $S.Address();
        $U = $ws.Cells.Item(1, 1);
        $U = $U.Address();
        $T = $U + ":" + $S;
        $ws.PageSetup.PrintArea = $T

        $ws.PageSetup.RightFooter = "Page &P of &N"
        
        #$xlFromPage = 1;
        #$xlToPage = 7;
        #$xlQuality = "Microsoft.Office.Interop.Excel.xlQualityStandard" -as [type];
        $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type];
        $wb.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfFilePathAndName);
        #$wb.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath, $empty_Var, $false, $true, $empty_Var, $empty_Var)
        #endregion
    }

    
    if($typeOfReport -eq  'PDFEmergencyContacts')
    {
        $logMessage = "Creating PDFEmergencyContacts workbook";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        #region delete all the sheets except for Membership
        $wb.Worksheets.Item(5).Delete()
        $wb.Worksheets.Item(3).Delete()
        $wb.Worksheets.Item(2).Delete()
        $wb.Worksheets.Item(1).Delete()
        $ws = $wb.Worksheets.Item(1);
        $ws.Select() | Out-Null;
        $ws.Activate();
        

        #$EC_ContactEmailCol
        $ws.Range("G1:K160").Clear() | Out-Null;

        $ws.PageSetup.Orientation = $xlLandscape;
        $ws.PageSetup.PrintTitleRows = '$1:$2';
        # https://www.litigationsupporttipofthenight.com/single-post/2016/06/20/powershell-script-to-print-excel-files-landscaped-to-pdfs
        $ws.PageSetup.LeftMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.RightMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.TopMargin = $xl.InchesToPoints(0.5);
        $ws.PageSetup.BottomMargin = $xl.InchesToPoints(0.5);
        $ws.PageSetup.HeaderMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.FooterMargin = $xl.InchesToPoints(0.25);
        $ws.PageSetup.PrintGridlines = $true;
        $ws.PageSetup.PrintComments = $xlPrintNoComments;
        $ws.PageSetup.FirstPageNumber = $xlAutomatic;
        $ws.PageSetup.Order = 1;
        $ws.PageSetup.Zoom = $false;
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = 9999;

        
        $range = $ws.usedRange;
        $r = $range.rows.count;
        #$c = $range.Columns.count;
        $c = $EC_ContactEmailCol;
        $S = $ws.Cells.Item($r,$c);
        $logMessage = "Column Count: " + $c.ToString();
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $S = $S.Address();
        $U = $ws.Cells.Item(1, 1);
        $U = $U.Address();
        $T = $U + ":" + $S;
        $ws.PageSetup.PrintArea = $T

        $ws.PageSetup.RightFooter = "Page &P of &N"
        
        #$xlFromPage = 1;
        #$xlToPage = 7;
        #$xlQuality = "Microsoft.Office.Interop.Excel.xlQualityStandard" -as [type];
        $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type];
        $wb.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdfFilePathAndName);
        #endregion
    }

    if($protectWorkBookAndSheets -and $typeOfReport -eq  'ChapterRoster')
    {
        $logMessage = "Protecting Workbook and sheets";
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $empty_Var = [System.Type]::Missing
        $true_Var = [System.Type]::true

        $membershipWS.Protect("NAL900115!", $true);
        $birthdaysThisMonthWS.Protect("NAL900115!", $true_Var);
        $birthdaysNextMonthWS.Protect("NAL900115!", $true_Var);
        $emergencyContactsWS.Protect("NAL900115!", $true_Var);
        $rolebasedEmailsWS.Protect("NAL900115!", $true_Var);
        $wb.Protect("NAL900115!", $empty_Var, $empty_Var);
    }

    #region clean up Excel stuff
    $logMessage = "Closing Workbook";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;



    $wb.SaveAs($workbookFilePathAndName);
    $xl.quit(); # Quit Excel...

    if($false)
    {
        Release_Ref($membershipWS) | Out-Null;
        Release_Ref($birthdaysThisMonthRow) | Out-Null;
        Release_Ref($birthdaysNextMonthWS) | Out-Null;
        Release_Ref($emergencyContactsWS) | Out-Null;
        Release_Ref($wb) | Out-Null;
        Release_Ref($xl) | Out-Null;
    }

    Get-Process -Name "Excel" -ErrorAction Ignore | Stop-Process -Force
    #endregion clean up Excel stuff

    #region updating APStatusTable with finish values
    $SProcessStopDateFinished = Get-Date;
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
                                                           -masterLogFilePathAndName $masterLogFilePathAndName | Out-Null;
    #endregion
    


}
end{}
