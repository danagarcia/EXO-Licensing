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
        Position=4)][String] $DisabledExchangeUsersGroupID,
    [Parameter(Mandatory=$True,
        Position=5)][String] $LicensedUsersGroupID,
    [Parameter(Mandatory=$True,
        Position=6)][String] $DisabledLicensedUsersGroupID,
    [Parameter(Mandatory=$False,
        Position=7)][String] $ScopeToSingleUser
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

Function Get-GroupMemberIDs
{
    param([Parameter(Mandatory=$True)][String]$groupID,
        [Parameter(Mandatory=$True)][String]$token)

    #Declare and populate variables
    $requestHeader = @{Authorization="Bearer $token"}
    $requestUri = "https://graph.microsoft.com/v1.0/groups/$groupID/members?`$select=id&`$top=999"
    $hasNext = $true
    $result = @()

    #Retrieve all member IDs from group paging through results (100 at a time)
    while($hasNext)
    {
        try
        {
            $response = Invoke-WebRequest -UseBasicParsing -Headers $requestHeader -Uri $requestUri
            switch($response.StatusCode)
            {
                200 {}
                429 {
                    #Graph throttling, sleep for 5 seconds.
                    Start-Sleep -Seconds 5
                    $response = Invoke-WebRequest -UseBasicParsing -Headers $requestHeader -Uri $requestUri
                }
                default {throw}
            }
            if ($groupMembersResultSet.StatusCode -ne 200)
            {
                #Error occurred throw error
                throw
            }
            $parsedJson = $response.Content | ConvertFrom-Json
            if($parsedJson.'@odata.nextLink')
            {
                $requestUri = $parsedJson.'@odata.nextLink' 
            }
            else
            {
                $hasNext = $false 
            }
            $parsedJson.Value | ForEach-Object {$result += $_.id}
        }
        catch
        {
            #Write error notification and exit
            Write-Host "An error occurred while retreiving group members from Graph API."
            Write-Host $_.Exception.Message
            Write-Host "Exiting"
            exit
        }
        
    }

    #All members retrieved return user id array
    return $result

}

Function Add-GroupMember
{
    param([Parameter(Mandatory=$True)][String]$User,
        [Parameter(Mandatory=$True)][String]$Group,
        [Parameter(Mandatory=$True)][String]$Token)

    #Create web request values
    $requestHeader = @{
        "Authorization" = "Bearer $Token"
        "Content-type" = "application/json"
    }
    $requestUri = "https://graph.microsoft.com/v1.0/groups/$Group/members/`$ref"
    $requestBody = @{'@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$User"} | ConvertTo-Json
    #Add user to group
    try
    {
        $response = Invoke-WebRequest -Method Post -UseBasicParsing -Headers $requestHeader -Uri $requestUri -Body $requestBody
        switch ($response.StatusCode)
        {
            #Successful status code
            204 {
                #Return success
                return @{
                    Status = 1
                    Message = "Success"
                }
            }
            #Unsuccessful status code
            default {throw}
        }
    }
    catch
    {
        #Return failure
        return @{
            Status = 0
            Message = "Error: $($_.Exception.Message)"
        }
    }
}

Function Remove-GroupMember
{
    param([Parameter(Mandatory=$True)][String]$User,
        [Parameter(Mandatory=$True)][String]$Group,
        [Parameter(Mandatory=$True)][String]$Token)
    #Create web request values
    $requestHeader = @{Authorization = "Bearer $Token"}
    $requestUri = "https://graph.microsoft.com/v1.0/groups/$Group/members/$User/`$ref"
    #Remove user from group
    try 
    {
        $response = Invoke-WebRequest -Method Delete -UseBasicParsing -Headers $requestHeader -Uri $requestUri   
        switch ($response.StatusCode)
        {
            #Successful status code
            204 {
                #Return success
                return @{
                    Status = 1
                    Message = "Success"
                }
            }
            #Unsuccessful status code
            default {throw}
        }
    }
    catch {
        #Return failure
        return @{
            Status = 0
            Message = $_.Exception.Message
        }
    }
}

Function Write-LogEntry
{
    param([Parameter(Mandatory=$True)][String]$User,
        [Parameter(Mandatory=$True)][String]$Status,
        [Parameter(Mandatory=$True)][String]$ErrorDetails,
        [Parameter(Mandatory=$True)][String]$Token)

    try
    {
        #Create web request values
        $requestHeader = @{Authorization = "Bearer $Token"}
        $requestUri = "https://graph.microsoft.com/v1.0/users/$User`?`$select=userPrincipalName"
        #Request user information from Graph
        $response = Invoke-WebRequest -UseBasicParsing -Headers $requestHeader -Uri $requestUri
        #Ensure user was found
        If($response.StatusCode -ne 200)
        {
            throw
        }
        #Pull User Principal Name from Graph response
        $userPrincipalName = ($response.Content | ConvertFrom-Json).userPrincipalName
        #Configure web request values to push data into Power BI streaming dataset
        $endpoint = "https://api.powerbigov.us/beta/66cf5074-5afe-48d1-a691-a12b2121f44b/datasets/e1875004-7410-4bab-9bfa-c3c3ea7d0df6/rows?tenant=66cf5074-5afe-48d1-a691-a12b2121f44b&UPN=EmlaeluZ%40state.gov&key=qPfyBSbPUrb6bET38jwSQRHfZ4QoIwxi4jLndOonJqeHE5ND2gVGhQpAiKJ%2FMbWJExyT8GY7JglfhBZu3NjO6A%3D%3D"
        $payload = @{
            "Time Stamp" = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss.000Z").ToString()
            "User Principal" = $userPrincipalName
            "Status" = $Status
            "Error Details" = $ErrorDetails
        }
        #Push data into Power BI streaming dataset
        Invoke-RestMethod -Method Post -Uri $endpoint -Body (ConvertTo-Json @($payload))
    }
    catch
    {
        Write-Host $_.Exception.Message
    }
}

#Get OAuth token
$token = Get-OAuthToken -TenantDomain $TenantDomain -ClientID $ClientID -ClientSecret $ClientSecret
#Request data via web request to Microsoft Graph
#Get all licensed users
$licensedUserIDs = Get-GroupMemberIDs -groupID $LicensedUsersGroupID -token $token
#Refresh token
$token = Get-OAuthToken -TenantDomain $TenantDomain -ClientID $ClientID -ClientSecret $ClientSecret
#Get all licensed Exchange users
$licensedExchangeUserIDs = Get-GroupMemberIDs -groupID $LicensedExchangeUsersGroupID -token $token
#Refresh token
$token = Get-OAuthToken -TenantDomain $TenantDomain -ClientID $ClientID -ClientSecret $ClientSecret
#Get all disabled Exchange users
$disabledExchangeUserIDs = Get-GroupMemberIDs -groupID $DisabledExchangeUsersGroupID -token $token
#Get all disabled users
$disabledLicensedUserIDs = Get-GroupMemberIDs -groupID $DisabledLicensedUsersGroupID -token $token
#Refresh token
$token = Get-OAuthToken -TenantDomain $TenantDomain -ClientID $ClientID -ClientSecret $ClientSecret
#Get licensed users that aren't licensed for Exchange
$licensedUserDif = $licensedUserIDs | Where-Object {$licensedExchangeUserIDs -notcontains $_}
#Dispose of licensed user IDs array
$licensedUserIDs = $null
#Start garbage collection
[GC]::Collect()
#Get licensed users that aren't licensed for Exchange but have been before
$reenabledUserDif = $licensedUserDif | Where-Object {$disabledExchangeUserIDs -contains $_}
#Dispose of disabled Exchange user IDs
$disabledExchangeUserIDs = $null
#Start garbage collection
[GC]::Collect()
#Filter re-enabled users from licensed users that aren't licensed for Exchange
$licensedUserDif = $licensedUserDif | Where-Object {$reenabledUserDif -notcontains $_}
#Get disabled users that are still licensed for Exchange
$disabledUserDif = $disabledLicensedUserIDs | Where-Object {$licensedExchangeUserIDs -contains $_}
#Dispose of licensed Exchange user IDs
$licensedExchangeUserIDs = $null
#Start garbage collection
[GC]::Collect()
If($ScopeToSingleUser)
{
    $disabledUserDif = $disabledUserDif | Where-Object {$_ -like $ScopeToSingleUser}
    $licensedUserDif = $licensedUserDif | Where-Object {$_ -like $ScopeToSingleUser}
    $reenabledUserDif = $reenabledUserDif | Where-Object {$_ -like $ScopeToSingleUser}
}
#Configure URL for mailbox settings
$mailboxSettingsURL = "https://graph.microsoft.com/v1.0/users/{0}/mailboxSettings"
#Add re-enabled users to Exchange licensing group, remove them from Exchange disabled group
ForEach($user in $reenabledUserDif)
{
    #Add user to Exchange licensing group
    $result = Add-GroupMember -User $user -Group $LicensedExchangeUsersGroupID -Token $token
    Switch ($result.Status)
    {
        0 {Write-LogEntry -User $user -Token $token -Status "Failed" -ErrorDetails $result.Message}
        1 {
            #Remove user from disabled Exchange group
            $result = Remove-GroupMember -User $user -Group $DisabledExchangeUsersGroupID -Token $token
            Switch ($result.Status)
            {
                0 {Write-LogEntry -User $user -Token $token -Status "Failed" -ErrorDetails $result.Message}
                1 {Write-LogEntry -User $user -Token $token -Status "Success" -ErrorDetails "N/A"}
                
            }
        }
        
    }
}
ForEach($user in $licensedUserDif)
{
    #Get mailbox settings of each user to determine if any of the mailboxes are active but unlicensed
    try 
    {
        $mailboxSettingResponse = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri ($mailboxSettingsURL -f $user)
        switch ($mailboxSettingResponse.StatusCode)
        {
            200 {
                #Add user to Exchange licensing group
                $result = Add-GroupMember -User $user -Group $LicensedExchangeUsersGroupID -Token $token
                Switch ($result.Status)
                {
                    0 {Write-LogEntry -User $user -Token $token -Status "Failed" -ErrorDetails $result.Message}
                    1 {Write-LogEntry -User $user -Token $token -Status "Success" -ErrorDetails "N/A"}
                                    } 
            }
            default {throw}
        }
    }
    catch 
    {
        #No need to handle this exception as it means the user doesn't have a mailbox we can see.
    }
}
ForEach($user in $disabledUserDif)
{
    #Remove user to Exchange licensing group
    $result = Remove-GroupMember -User $user -Group $LicensedExchangeUsersGroupID -Token $token
    Switch ($result.Status)
    {
        0 {Write-LogEntry -User $user -Token $token -Status "Failed - Disabled" -ErrorDetails $result.Message}
        1 {
            #Add user from disabled Exchange group
            $result = Add-GroupMember -User $user -Group $DisabledExchangeUsersGroupID -Token $token
            Switch ($result.Status)
            {
                0 {Write-LogEntry -User $user -Token $token -Status "Failed - Disabled" -ErrorDetails $result.Message}
                1 {Write-LogEntry -User $user -Token $token -Status "Removed - Disabled" -ErrorDetails "N/A"}
            }
        }
    }
} 
