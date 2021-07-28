<#
Name: UnblockScripts.ps1
Created By: Mike John
Created Date: 07/29/2020
Summary:
    Unblocks files downloaded from the internet
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file?view=powershell-7

Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 07/29/2020
    Reason Updated: Original version
#>

# starts here
cls;
# This lists file that are blocked
Get-Item * -Stream "Zone.Identifier" -ErrorAction SilentlyContinue;

# This unblocks file that are blocked in the specified directory command

$dirPath = "C:\PSScripts\NAL\*.ps1";

dir $dirPath | Unblock-File;