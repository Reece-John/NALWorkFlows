<# Header Information **********************************************************
Name: ReturnTenantCredentials.ps1
Created By: Mike John
Created Date: 01/11/2021
Summary:
    Retrieves remote Credentials from Storage it uses the local domain parameter as part of the credentials
Prerequisites:
    Environment variables must be set:
                                    "CStorage", "User"
                                    "TenantUser","User"
                                    "UserDomain","User"
                                    "DomainExtension","User"
    Encrypted password must be stored at "CStorage", "User"
Update History ******************************************************************************************
Updated By: Mike John
UpdatedDate: 01/11/2021
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][PSCustomObject]$tenantObj
)
begin {}
process {
    [string]$storageDirectory = $tenantObj.CommonStorageDir;
    [string]$adminName = $tenantObj.PSAdminUser;
    
    $userDomain = $tenantObj.DomainName;
    $domainExtension = $tenantObj.DomainExtension;
    $userEmail = $adminName + "@" + $userDomain + "." + $domainExtension;

    [string]$CurrentDomainAndUser = whoami;
    $CurrentUser = $CurrentDomainAndUser.Substring($CurrentDomainAndUser.IndexOf("\") + 1 );
    $CurrentUser = $CurrentUser.Replace('.', '');
    $CurrentUser = $CurrentUser.Replace(' ', '');
    $envObj = Get-ChildItem Env:USERNAME;
    $CurrentUser = $envObj.Value;

    $CredsFile = "$storageDirectory\" + $userDomain + "_" + $CurrentUser + "_" + $adminName + "_PowershellCreds.txt";
    $CredsFileExists = Test-Path $CredsFile -ErrorAction SilentlyContinue;
    if($CredsFileExists)
    {
        [SecureString]$password = get-content $CredsFile | ConvertTo-SecureString
        $tempCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $userDomain\$adminName,$password
        $clearTextPassword = $tempCredentials.GetNetworkCredential().password;
        $password = $clearTextPassword | ConvertTo-SecureString -asPlainText -Force
        $SecureCredentials = New-Object System.Management.Automation.PSCredential($userEmail,$password)
        return $SecureCredentials;
    }
    else
    {
        return $null
    }
}
