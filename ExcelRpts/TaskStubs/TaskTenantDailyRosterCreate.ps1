#
# TaskTenantDailyRosterCreate.ps1
#
<# Header Information **********************************************************
Name: TaskUpdateALChapterSchema.ps1
Created By: Mike John
Created Date: 02/23/2021
Summary:
    Task stub to run TenantDailyRosterCreate.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 07/06/2021
    Reason Updated: Added $justTesting parameter"
Updated By: Mike John
Updated Date: 02/23/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][bool]$justTesting
      )
begin {}
process {
    # Starts Here
    [bool]$downloadALChapterSchemaFile = $true;
    [bool]$uploadChapterRoster = $true;
    [bool]$uploadPDFRoster = $true;
    [bool]$uploadPDFEmergencyContacts = $true;
    [bool]$SendArchiveEmailEmail = $true;
    if($justTesting)
    {
        $uploadChapterRoster = $false;
        $uploadPDFRoster = $false;
        $uploadPDFEmergencyContacts = $false;
        $SendArchiveEmailEmail = $false;
    }

    # get tenant specific variable values
    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    [string]$tenantDomain = $tenantObj.DomainName;


    Get-Process -Name "Excel" -ErrorAction Ignore | Stop-Process -Force

    # create the log file name
    [DateTime]$dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\' + $tenantAbbreviation + ' RosterCreate_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';


    [string]$thisScriptName = $MyInvocation.MyCommand.Name
    [string]$logMessage = "Starting " + $thisScriptName;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    # Assign values to variables
    # SharePoint File download variables
    [string]$downloadSchemaFileName = "ALChapterSchema.xlsx";
    [string]$technologySiteURL =     "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam-M365Data";
    [string]$downLoadschemaFileRelativeURL = "/sites/TechnologyTeam-M365Data/Shared Documents/M365Data/" + $downloadSchemaFileName;
    
    [string]$downloadPath = $tenantObj.PSStartUpDir + "\" + $tenantAbbreviation + "\ExcelDataFiles"
    [string]$existingALChapterSchemaFile = $downloadPath + "\" + $downloadSchemaFileName;

    # Clear file to begin
    Remove-Item $existingALChapterSchemaFile -ErrorAction SilentlyContinue;

    # get administrator credentials
    [system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    # Get latest copy of Chapter Schema Excel Spreadsheet from SharePoint
    $logMessage = "Getting latest copy of Chapter Schema Excel Spreadsheet from SharePoint";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    if($downloadALChapterSchemaFile)
    {
        .\M365SharePoint\DownLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                    -SharePointSiteURL $technologySiteURL `
                                                    -SharePointFileRelativeURL $downLoadschemaFileRelativeURL `
                                                    -LocalFileDownloadPath $downloadPath `
                                                    -FileName $downloadSchemaFileName `
                                                    -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }

    <#
    Make sure the ChapterSchema.xlsx file has today's date in it
    if()
    #>

    # define parts of spreadsheets we are creating
    [string]$workbookFileName= "Chapter Roster";
    [string]$workbookFileExtension= "xlsx";
    [string]$workbookFileFullName= $workbookFileName + "." + $workbookFileExtension;
    [string]$workbookFilePathAndName = $downloadPath + "\" + $workbookFileFullName;
    Remove-Item $workbookFilePathAndName -ErrorAction SilentlyContinue;

    $logMessage = "Creating ChapterRoster XLSX version";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    # Assign more values to variables
    [string]$schemaFilePathName = $DownloadPath + "\" + $downloadSchemaFileName;
    [string]$templateMembershipPageName = "Membership";
    [string]$templateBirthdaysThisMonthPageName = "Birthdays This Month";
    [string]$templateBirthdaysNextMonthPageName = "Birthdays Next Month";
    [string]$templateEmergencyContactPageName = "Emergency Contact Information";
    [string]$templateRolebasedEmailsSheetName = "Role-based Emails Information";
    [string]$NALUsersPageName = "M365Users";
    [string]$NALRoleBasedEmailsSheetName = "RoleBasedEmailAddresses";
    [bool]$protectWbAndSheets = $true;
    [string]$typeOfReportToMake = "ChapterRoster";
    $chapterRosterPdfFilePathAndName = "Not Used";

    # Create ChapterRoster.xlsx
    .\ExcelRpts\TenantDailyRosterCreate.ps1 -tenantCredentials $psAdminCredentials `
                                            -tenantAbbreviation $tenantAbbreviation `
                                            -tenantObj $tenantObj `
                                            -workbookFilePathAndName $workbookFilePathAndName `
                                            -pdfFilePathAndName $chapterRosterPdfFilePathAndName `
                                            -rosterMembershipSheetName $templateMembershipPageName `
                                            -rosterBirthdaysThisMonthSheetName $templateBirthdaysThisMonthPageName `
                                            -rosterBirthdaysNextMonthSheetName $templateBirthdaysNextMonthPageName `
                                            -rosterEmergencyContactsSheetName $templateEmergencyContactPageName `
                                            -rosterRolebasedEmailsSheetName $templateRolebasedEmailsSheetName `
                                            -m365SchemaFilePathAndName $schemaFilePathName `
                                            -m365UsersSheetName $NALUsersPageName `
                                            -m365roleBasedEmailAddressesSheetName $NALRoleBasedEmailsSheetName  `
                                            -protectWorkBookAndSheets $protectWbAndSheets `
                                            -typeOfReport $typeOfReportToMake `
                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;

    # upload "Chapter Roster.xlsx" to Chapter Team Storage Site
    # Configuration Variables
    $uploadSiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/NALChapterTeam";
    $uploadFileRelativeURL = "/Shared Documents/General/ChapterRoster";
    # one of the parameters is set above for the copy command to use also
    if($uploadChapterRoster)
    {
        .\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                  -SharePointSiteURL $uploadSiteURL `
                                                  -SharePointFileRelativeURL $uploadFileRelativeURL `
                                                  -UploadPathAndFileName $workbookFilePathAndName `
                                                  -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }

    $logMessage = "Creating ChapterRoster PDF version";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    $protectWbAndSheets = $false;
    [string]$typeOfReportToMake = "PDFRoster";

    [string]$excelToPDFrosterFile = "PDF Chapter Roster.xlsx";
    $workbookFilePathAndName = $downloadPath + "\" + $excelToPDFrosterFile;

    [string]$chapterRosterPdfFile = "Chapter Roster.pdf";
    [string]$chapterRosterPdfFilePathAndName = $downloadPath + "\" + $chapterRosterPdfFile;
    # Create ChapterRoster.pdf
    .\ExcelRpts\TenantDailyRosterCreate.ps1 -tenantCredentials $psAdminCredentials `
                                            -tenantAbbreviation $tenantAbbreviation `
                                            -tenantObj $tenantObj `
                                            -workbookFilePathAndName $workbookFilePathAndName `
                                            -pdfFilePathAndName $chapterRosterPdfFilePathAndName `
                                            -rosterMembershipSheetName $templateMembershipPageName `
                                            -rosterBirthdaysThisMonthSheetName $templateBirthdaysThisMonthPageName `
                                            -rosterBirthdaysNextMonthSheetName $templateBirthdaysNextMonthPageName `
                                            -rosterEmergencyContactsSheetName $templateEmergencyContactPageName `
                                            -rosterRolebasedEmailsSheetName $templateRolebasedEmailsSheetName `
                                            -m365SchemaFilePathAndName $schemaFilePathName `
                                            -m365UsersSheetName $NALUsersPageName `
                                            -m365roleBasedEmailAddressesSheetName $NALRoleBasedEmailsSheetName  `
                                            -protectWorkBookAndSheets $protectWbAndSheets `
                                            -typeOfReport $typeOfReportToMake `
                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;
      
    # Store Chapter Roster.pdf file in SharePoint storage
    # Configuration Variables
    $uploadSiteURL = "https://algeorgetownarea.sharepoint.com/sites/NALChapterTeam";
    $uploadFileRelativeURL = "/Shared Documents/General/ChapterRoster";
    $mySourceFilePath ="C:\PSScripts\NAL\ExcelDataFiles\Chapter Roster.pdf";
    if($uploadPDFRoster)
    {
        .\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                  -SharePointSiteURL $uploadSiteURL `
                                                  -SharePointFileRelativeURL $uploadFileRelativeURL `
                                                  -UploadPathAndFileName $mySourceFilePath `
                                                  -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }

    $logMessage = "Creating Emergency Contact List PDF version";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    $protectWbAndSheets = $false;
    [string]$typeOfReportToMake = "PDFEmergencyContacts";
    [string]$excelToPDFrosterFile = "PD Emergency Contacts.xlsx";
    $workbookFilePathAndName = $downloadPath + "\" + $excelToPDFrosterFile;

    [string]$chapterRosterPdfFile = "Chapter Emergency Contacts.pdf";
    [string]$chapterRosterPdfFilePathAndName = $downloadPath + "\" + $chapterRosterPdfFile;
    # 3333333333333333333333333333333333333333333333
    .\ExcelRpts\TenantDailyRosterCreate.ps1 -tenantCredentials $psAdminCredentials `
                                            -tenantAbbreviation $tenantAbbreviation `
                                            -tenantObj $tenantObj `
                                            -workbookFilePathAndName $workbookFilePathAndName `
                                            -pdfFilePathAndName $chapterRosterPdfFilePathAndName `
                                            -rosterMembershipSheetName $templateMembershipPageName `
                                            -rosterBirthdaysThisMonthSheetName $templateBirthdaysThisMonthPageName `
                                            -rosterBirthdaysNextMonthSheetName $templateBirthdaysNextMonthPageName `
                                            -rosterEmergencyContactsSheetName $templateEmergencyContactPageName `
                                            -rosterRolebasedEmailsSheetName $templateRolebasedEmailsSheetName `
                                            -m365SchemaFilePathAndName $schemaFilePathName `
                                            -m365UsersSheetName $NALUsersPageName `
                                            -m365roleBasedEmailAddressesSheetName $NALRoleBasedEmailsSheetName  `
                                            -protectWorkBookAndSheets $protectWbAndSheets `
                                            -typeOfReport $typeOfReportToMake `
                                            -masterLogFilePathAndName $myMasterLogFilePathAndName;
      
    # Store excel-to-PDF file in SharePoint storage
    # Configuration Variables
    $uploadSiteURL = "https://algeorgetownarea.sharepoint.com/sites/NALChapterTeam";
    $uploadFileRelativeURL = "/Shared Documents/General/ChapterRoster";
    $mySourceFilePath ="C:\PSScripts\NAL\ExcelDataFiles\Chapter Emergency Contacts.pdf";
    if($uploadPDFEmergencyContacts)
    {
        .\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $psAdminCredentials `
                                                  -SharePointSiteURL $uploadSiteURL `
                                                  -SharePointFileRelativeURL $uploadFileRelativeURL `
                                                  -UploadPathAndFileName $mySourceFilePath `
                                                  -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }

    [int]$emailDayOfMonth = (Get-Date).Day;
    if($emailDayOfMonth -eq 1 -and $SendArchiveEmailEmail)
    {
        $logMessage = "Sending Email to mass mailing coordinator";
        .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

        # email to mass mailing coordinator that PDF file is available and it first of month
        [string]$myEmailSender           = "mjohn@algeorgetownarea.org";
        [string[]]$myEmailRecipientArray = @("librarian@algeorgetownarea.org", "cgraham@algeorgetownarea.org", "mjohn@algeorgetownarea.org");
        [string]$myEmailSubject          = "Monthly Chapter Roster.pdf - $(Get-Date -Format g)";
        [string]$myEmailBody             = "Monthly Chapter Roster PDF process is Complete, Please send out mass mailing.- $(Get-Date -Format g)";
        .\EMailer\NotifyProcessCompleted.ps1 -tenantCredentials $psAdminCredentials `
                                             -emailSender $myEmailSender `
                                             -emailRecipientArray $myEmailRecipientArray `
                                             -emailSubject $myEmailSubject `
                                             -emailBody $myEmailBody `
                                             -masterLogFilePathAndName $myMasterLogFilePathAndName;
    }
    $logMessage = "Finished TaskTenantDailyRosterCreate.ps1";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
}
