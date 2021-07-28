<# Header Information **********************************************************
Name: ReturnTenantSpecificVariables.ps1
Created By: Mike John
Created Date: 03/23/2021
Summary:
    Retrieves Tenant specific variable in a PowerShell object
        DevStartUpDir           - Start up directory when debugging
        DomainName              - Tenant Domain Name without extension
        DomainExtension         - Tenant Domain Name extension
        CommonStorageDir        - This is where the encrypted information is stored
        FiscalYearStartDate     - The start of the Fiscal Year; Used for accounting
        PSAdminUser             - This must be a global admin
        PSStartUpDir            - this is the start up directory in production
        TeamSiteName            - Name of a Tenant Team site
        UserExportFileName      - Name of the export file from CHapter Hub or Chapter Web without extension
        UserExportFileExtension - Name of the export file extension from CHapter Hub or Chapter Web
        TenantSchemaFileName    - Name of the Tenant Schema file
        VolunteerYearStartDate  - The start of the Volunteer Year; used to determine new members and to determine when to purge schema file of resigned and deceased

Prerequisites:
    variables must be set:
                           DevStartUpDir        
                           DomainName           
                           DomainExtension      
                           CommonStorageDir     
                           PSAdminUser          
                           PSStartUpDir         
                           TeamSiteName         
                           UserExportFileName   
                           TenantSchemaFileName 
    Encrypted password must be stored at "CommonStorageDir"
Update History ******************************************************************************************
Updated By: Mike John
UpdatedDate: 03/23/2021
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
       [Parameter(Mandatory=$True,Position=0)][ValidateSet("ALGA","ALSA","ALLV")][string]$tenantAbbreviation
    )
begin {}
process {

    #region function definitions
    function ReturnYearStartDate([int]$yearStartDay, [int]$yearStartMonth, [dateTime]$currentDate)
    {
        # used to get start of Volunteer year and fiscal year; both at midnight
        [int]$curentMonth = $currentDate.Month;
        [int]$currentYear = $currentDate.Year;
        
        if($curentMonth -lt $yearStartMonth)
        {
            $currentYear = $currentYear - 1;
        }
        
        [DateTime]$yearStartDate = Get-Date -Year $currentYear -Month $yearStartMonth -Day $yearStartDay -Hour 0 -Minute 0 -Second 0;

        return $yearStartDate;
    }
    #endregion function definitions

    [string]$chapterRosterTitle      = "";
    [string]$psAdminUserName         = "";
    [string]$DomainName              = "";
    [string]$teamSiteName            = "";
    [string]$devStartUpDir           = "";
    [string]$UserExportFileName      = "";
    [string]$UserExportFileExtension = "";
    [string]$TenantSchemaFileName    = "";
        
    [string]$PSStartUpDir           = "c:\PSScripts";
    [string]$CommonStorageDir       = $PSStartUpDir + "\" + $tenantAbbreviation + "\Common";
    [string]$DomainExtension        = "Org";
    [string]$CurrentUserName        = [Environment]::UserName;
    [int]$fiscalYearStartDay;
    [int]$fiscalYearStartMonth;
    [DateTime]$fiscalYearStartDate = ReturnYearStartDate -yearStartDay 1 -yearStartMonth 1 -currentDate (Get-Date);
    [DateTime]$volunteerYearStartDate = ReturnYearStartDate -yearStartDay 1 -yearStartMonth 1 -currentDate (Get-Date);


    switch ( $tenantAbbreviation )
    {
        'ALGA'
        {
            $chapterRosterTitle = "Assistance League Georgetown Area";
            $psAdminUserName = "MJohn";
            $DomainName = "ALGeorgetownArea";
            $teamSiteName = "ALGA Chapter Team";
            $devStartupDir   = "C:\Users\" + $CurrentUserName + "\Source\Repos\PSWorkFlows";
            $UserExportFileName = "ALGA Nightly User Export";
            $TenantSchemaFileName = "ALChapterSchema";
            $UserExportFileExtension = "csv";

            $notifyIfDailyExportFileIsMissing = $true;

            $fiscalYearStartDay = 1;
            $fiscalYearStartMonth = 6;
            $fiscalYearStartDate = ReturnYearStartDate -yearStartDay $fiscalYearStartDay -yearStartMonth $fiscalYearStartMonth -currentDate (Get-Date);

            $volunteerYearStartDay = 1;
            $volunteerYearStartMonth = 6;
            $volunteerYearStartDate = ReturnYearStartDate -yearStartDay $volunteerYearStartDay -yearStartMonth $volunteerYearStartMonth -currentDate (Get-Date);
        }
        'ALSA'
        {
            $chapterRosterTitle = "Assistance League San Antonio";
            $psAdminUserName = "MJohn";
            $DomainName = "ALSanAntonio";
            $teamSiteName = "ALSA Chapter Team";
            $devStartupDir   = "C:\Users\" + $CurrentUserName + "\Source\Repos\PSWorkFlows";
            $UserExportFileName = "ALSA User Export";
            $TenantSchemaFileName = "ALChapterSchema";
            $UserExportFileExtension = "csv";

            $notifyIfDailyExportFileIsMissing = $false;

            $fiscalYearStartDay = 1;
            $fiscalYearStartMonth = 6;
            $fiscalYearStartDate = ReturnYearStartDate -yearStartDay $fiscalYearStartDay -yearStartMonth $fiscalYearStartMonth -currentDate (Get-Date);

            $volunteerYearStartDay = 1;
            $volunteerYearStartMonth = 1;
            $volunteerYearStartDate = ReturnYearStartDate -yearStartDay $volunteerYearStartDay -yearStartMonth $volunteerYearStartMonth -currentDate (Get-Date);
        }
        'ALLV'
        {
            $chapterRosterTitle = "Assistance League Las Vagas";
            $psAdminUserName = "CWarman";
            $DomainName = "ALLasVegas";
            $teamSiteName = "ALLV Chapter Team";
            $devStartupDir   = "C:\Users\" + $CurrentUserName + "\Source\Repos\PSWorkFlows";
            $UserExportFileName = "ALLV User Export";
            $TenantSchemaFileName = "ALChapterSchema";
            $UserExportFileExtension = "csv";
            
            $notifyIfDailyExportFileIsMissing = $false;

            $fiscalYearStartDay = 1;
            $fiscalYearStartMonth = 6;
            $fiscalYearStartDate = ReturnYearStartDate -yearStartDay $fiscalYearStartDay -yearStartMonth $fiscalYearStartMonth -currentDate (Get-Date);

            $volunteerYearStartDay = 1;
            $volunteerYearStartMonth = 6;
            $volunteerYearStartDate = ReturnYearStartDate -yearStartDay $volunteerYearStartDay -yearStartMonth $volunteerYearStartMonth -currentDate (Get-Date);
        }
        default
        {
            Write-Host("No tenantAbreviation Match found");
            Exit; # No tenantAbreviation Matches
        }
    }
    

    [PSCustomObject]$tenantObj = [PSCustomObject][ordered]@{
        ChapterRosterTitle               = $chapterRosterTitle
        CommonStorageDir                 = $CommonStorageDir
        DevStartUpDir                    = $devStartUpDir
        DomainExtension                  = $DomainExtension
        DomainName                       = $DomainName
        FiscalYearStartDy                = $fiscalYearStartDay
        FiscalYearStartMonth             = $fiscalYearStartMonth
        FiscalYearStartDate              = $fiscalYearStartDate
        NotifyIfDailyExportFileIsMissing = $notifyIfDailyExportFileIsMissing
        PSAdminUser                      = $psAdminUserName
        PSStartUpDir                     = $PSStartUpDir
        TenantSchemaFileName             = $TenantSchemaFileName
        TeamSiteName                     = $teamSiteName
        UserExportFileName               = $UserExportFileName
        UserExportFileExtension          = $UserExportFileExtension
        VolunteerYearStartDay            = $volunteerYearStartDat
        VolunteerYearStartMonth          = $volunteerYearStartMonth
        VolunteerYearStartDate           = $volunteerYearStartDate
       }
       return $tenantObj;
}
