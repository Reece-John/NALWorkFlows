<# Header Information **********************************************************
Name: RemoteTenantCredentialsStorage.ps1
Created By: Mike John
Created Date: 01/13/2021
Summary:
    Encrypts and stores password
    http://social.technet.microsoft.com/wiki/contents/articles/4546.working-with-passwords-secure-strings-and-credentials-in-windows-powershell.aspx
    http://www.techrepublic.com/blog/networking/powershell-code-to-store-user-credentials-encrypted-for-re-use/5817
    http://poshcode.org/3528
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 01/13/2021
    Reason Updated: Original version adapted from a previous version
#>
[cmdletbinding()]
Param()
begin {}
process {
    # starts here
    clear-host;

    $tenantAbbreviation = Read-Host 'Enter Abbreviation for Tenant (All Caps):'
    $tenantDomainName = Read-Host 'Enter Domain Name for Tenant (No extension):'
    $tenantUserName = Read-Host 'Enter Global Administrator User Name (First part of Email) for Tenant:'

    [string]$tenantCStorageName = $tenantAbbreviation + "CStorage";

    $cStorageDirectory = [Environment]::GetEnvironmentVariable($tenantCStorageName,"User");
    if($cStorageDirectory -eq $null)
    {
        Write-Host("Environment variable cStorageDirectory scope user not set");
        exit;
    }

    
    [string]$domainAndCurrentUser = whoami;
    #$domain = $domainAndCurrentUser.Substring(0, $domainAndCurrentUser.IndexOf("\"));
    $CurrentUser = $domainAndCurrentUser.Substring($domainAndCurrentUser.IndexOf("\") + 1 );
    $CurrentUser = $CurrentUser.Replace('.', '');
    $CurrentUser = $CurrentUser.Replace(' ', '');

    $CredsFile = "$cStorageDirectory\" + $tenantDomainName + "_" + $CurrentUser + "_" + $tenantUserName + "_PowershellCreds.txt";
    $FileExists = Test-Path $CredsFile -ErrorAction SilentlyContinue;
    #<#
    # Next 5 lines used for testing
    write-host "  tenantUserName:" $tenantUserName;
    write-host "tenantDomainName:" $tenantDomainName;
    write-host "     CurrentUser:" $CurrentUser;
    write-host "       CredsFile:" $CredsFile;
    write-host "      FileExists:" $FileExists;
    #
    #>
    if(!$FileExists)
    {
        #Write-Host 'Enter your password:'
        Read-Host 'Enter User Password for Remote Tenant:' -AsSecureString | ConvertFrom-SecureString | Out-File $CredsFile
        Write-Host "Password stored at: $CredsFile"
    }
    else
    {
        Write-Host 'Stored credential file for this user already exists'
    }
}
end {}
