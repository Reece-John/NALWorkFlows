#
# UpdateSharePointDocumentList.ps1
#

#log starting here

$userCredential = Get-Credential -Message "Type the password.";
$WebUrl = "https://SaggieHaim.sharepoint.com/sites/Portal";
Connect-PnPOnline –Url $WebUrl –Credentials $userCredential;

# Monthly Birthdays
$listName = "MonthlyBirthdays";

$employeesBirthday = Import-Csv -Path "D:\Tasks\Update SharePoint Events list\WeeklyBirthdays.CSV";

Get-PnPListItem -List "$listName" | foreach { Remove-PnPListItem -List "$listName" -Identity $_.Id -Force};

foreach ($employee in $employeesBirthday) 
{
    ## Replacing first and last names
    $EventAuthor = $employee.name.Split(" ")[1] + " " + $employee.name.Split(" ")[0];
    $expire = $employee.Expires;
    Add-PnPListItem -List "LumenisGreetingsList" -Values @{
                    "Title" = "IMF" ; 
                    "EventDay" = "$employee.EventDay"; 
                    "EventMonth" = "$employee.EventMonth"; 
                    "EventType" = "birthday";
                    "Role" = "$employee.Role"; 
                    "EventAuthor" = $EventAuthor; 
                    "Expires" = "$employee.expire"
                    };
}

#log success here
