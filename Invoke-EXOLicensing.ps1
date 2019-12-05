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
    This script check the users who are members of three different groups.
    
    1. Licensed Users (All plans excluding Exchange)
    2. Licensed Exchange Users (Only Exchange plans)
    3. Licensed Disabled Users (Only Exchange and Sharepoint plans)

    It then compares users who are in group 1 (licensed users) with users in
    group 2 (licensed Exchange users). This allows us to identify which users
    are licensed and active who don't have an Exchange license. It then checks
    to see if it can pull those users mailbox settings via Graph API. If it can't
    we can safely assume that they don't have a valid mailbox. If we can pull
    the settings then we now know they have a mailbox without a license. We then
    add them to group 2 for group based licensing. The script also compares users
    who are in group 3 (licensed disabled users) with users in group 2. If there
    are any users from group 2 present in group 3 they are removed. This is to
    ensure proper license hygiene.
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