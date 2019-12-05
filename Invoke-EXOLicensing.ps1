<#
.NOTES
    Name: Invoke-EXOLicensing.ps1
    Author: Dana Garcia
    Requires: Azure Automation Account, Azure Automation Runbook, Azure Active Directory,
    Azure Active Directory Registered Application.
    Version History:
    1.0 - 12/4/2019 - Initial release.
    ###############################################################################
    The sample scripts are not supported under any Microsoft standard support
    program or service. The sample scripts are provided AS IS without warranty
    of any kind. Microsoft further disclaims all implied warranties including,
    without limitation, any implied warranties of merchantability or of fitness
    for a particular purpose. The entire risk arising out of the use or
    performance of the sample scripts and documentation remains with you. In no
    event shall Microsoft, its authors, or anyone else involved in the creation,
    production, or delivery of the scripts be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits,
    business interruption, loss of business information, or other pecuniary
    loss) arising out of the use of or inability to use the sample scripts or
    documentation, even if Microsoft has been advised of the possibility of such
    damages.
    ###############################################################################
.SYNOPSIS
    This script finds all unlicensed user mailboxes in the O365 tenant, and
    adds them to the group used for EXO group based licensing, and notes any
    work performed and/or issues found in a Power BI dataset.
.DESCRIPTION
    This script checks the status of the SKUAssigned attribute on EXO mailboxes,
    and if it isn't set to true then it considers the mailbox unlicensed. It
    then checks the O365 license status of all EXO unlicensed mailboxes, and
    notes which ones are missing the core O365 license and those which Azure AD
    think are EXO licensed (meaning there is a sync issue between Azure AD and
    EXO). Lastly it tries to add the unlicensed mailboxes the group designated
    for EXO group based licensing. All of the discovered information and actions
    taken are exported to a CSV report file at the end.
.PARAMETER OutCSVFile
    Specifies the path and file name of the CSV to     export the collected data
    to. If no value is supplied, then a default file named
    ProcessedMailboxes-MM-dd-yyyy.CSV is created in the current folder, where
    the MM-dd-yyyy are the current date.
.PARAMETER WhatIf
    Optional: Instructs the script to not make any changes, just simulate what
    user accounts would be added to the license group, and the results are
    stored in the default CSV file name format in current directory.
.EXAMPLE
    Invoke-EXOMailboxGroupLicensing.ps1 -OutCSVFile .\NewMailboxLicenseReport.CSV
    The script performs its normal mailbox licensing check, adding regular user
    mailbox accounts to the designated group in Azure AD, and the results are
    stored in the NewMailboxLicenseReport.CSV file in the current directory.
.EXAMPLE
    Invoke-EXOMailboxGroupLicensing.ps1 -WhatIf
    All actions in the script are simulated and the HTML report is sent to the
    predefined SMTP address.
.LINK
    https://www.github.com/danagarcia/EXO-Licensing
#>
param(
    [Parameter(Mandatory=$True,
        Position=0)][String] $TenantDomain,
    [Parameter(Mandatory=$True,
        Position=1)][String] $ClientID,
    [Parameter(Mandatory=$True,
        Position=2)][String] $ClientSecret,
    [Parameter(Mandatory=$True,
        Position=3)][String] $LicensedExchangeUsersGroupID,
    [Parameter(Mandatory=$True,
        Position=4)][String] $LicensedUsersGroupID,
    [Parameter(Mandatory=$True,
        Position=5)][String] $DisabledLicensedUsersGroupID
)

#Declare static variables
$loginURL = 'https://login.microsoft.com'
$resource = 'https://graph.microsoft.com'

Function Get-OAuthToken
{
    param([Parameter(Mandatory=$True)][String]$TenantDomain,
        [Parameter(Mandatory=$True)][String]$ClientID,
        [Parameter(Mandatory=$True)][String]$ClientSecret)

    #Build the OAuth request
    $body = @{grant_type='client_credentials';resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    #Request OAuth token
    try
    {
        $response = Invoke-RestMethod -Method Post -Uri $loginURL/$TenantDomain/oauth2/token?api-version=1.0 -Body $body
    }
    catch
    {
        Write-Error "Unable to obtain OAuth token, Error message: $($_.Exception.Message)"
        exit
    }
    #Return OAuth token
    return $response.access_token
}

#Get OAuth token
$token = Get-OAuthToken -TenantDomain $TenantDomain -ClientID $ClientID -ClientSecret $ClientSecret
#Build Microsoft Graph web request
$headerParams = @{Authorization="Bearer $token"}
$groupMembersIDsURL = "https://graph.microsoft.com/v1.0/groups/{0}/members?`$select=id"
#Request data via web request to Microsoft Graph
$licensedUsersResultSet = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri ($groupMembersIDsURL -f $LicensedUsersGroupID)
$licensedExchangeUsersResultSet = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri ($groupMembersIDsURL -f $LicensedExchangeUsersGroupID)
$disabledLicensedUsersResultSet = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri ($groupMembersIDsURL -f $DisabledLicensedUsersGroupID)
$licensedUserIDs = @()
$licensedExchangeUserIDs = @()
$disabledLicensedUserIDs = @()
ForEach($user in ($licensedUsersResultSet.Content | ConvertFrom-Json).value)
{
    $licensedUserIDs += $user.id
}
ForEach($user in ($licensedExchangeUsersResultSet.Content | ConvertFrom-Json).value)
{
    $licensedExchangeUserIDs += $user.id
}
ForEach($user in ($disabledLicensedUsersResultSet.Content | ConvertFrom-Json).value)
{
    $disabledLicensedUserIDs += $user.id
}
$licensedUserDif = $licensedUserIDs | Where-Object {$licensedExchangeUserIDs -notcontains $_}
$disabledUserDif = $disabledLicensedUserIDs | Where-Object {$licensedExchangeUserIDs -contains $_}
$mailboxSettingsURL = "https://graph.microsoft.com/v1.0/users/{0}/mailboxSettings"
$addMemberURL = "https://graph.microsoft.com/v1.0/groups/{0}/members/`$ref"
$addMemberHeaderParams = ($headerParams += @{'Content-type'='application/json'})
ForEach($user in $licensedUserDif)
{
    try 
    {
        $mailboxSettingResponse = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri ($mailboxSettingsURL -f $user)
        switch ($mailboxSettingResponse.StatusCode)
        {
            200 {
                try {
                    $body = @{'@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$user"} | ConvertTo-Json
                    $addMemberResponse = Invoke-WebRequest -Method Post -UseBasicParsing -Headers $addMemberHeaderParams -Uri ($addMemberURL -f $LicensedExchangeUsersGroupID) -Body $body
                    switch ($addMemberResponse.StatusCode)
                    {
                        204 {Write-Host "User: $user, has been added to the Exchange licensing group."}
                        default {throw}
                    }
                }
                catch
                {
                    Write-Host "Unable to add user $user to Exchange online licensing group."
                    Write-Error "Error: $($_.Exception.Message)"
                } 
            }
            default {throw}
        }
    }
    catch 
    {
        Write-Host "User: $user, has no mailbox settings therefore no mailbox."
    }
}
$removeGroupMemberURL = "https://graph.microsoft.com/v1.0/groups/{0}/members/{1}/`$ref"
ForEach($user in $disabledUserDif)
{
    try 
    {
        $removeMemberResponse = Invoke-WebRequest -Method Delete -UseBasicParsing -Headers $headerParams -Uri ($removeGroupMemberURL -f $LicensedExchangeUsersGroupID,$user)   
        switch ($removeMemberResponse.StatusCode)
        {
            204 {Write-Host "User: $user, removed from Exchange licensing group."}
            default {throw}
        }
    }
    catch {
        Write-Host "Unable to remove user: $user from Exchange licensing group."
    }
}