<#//***********************************************************************
//
// Demo-App.ps1
// Modified 02 March 2023
// Last Modifier:  Jim Martin
// Project Owner:  Jim Martin
// .VERSION 20230302.0001
//
// .SYNOPSIS
//  Test application permissions using PowerShell script
// 
// .DESCRIPTION
//  This script will run Get commands in your Exchange Management Shell to collect configuration data via PowerShell
//
// .PARAMETERS
//    MailboxName - The mailbox you want to test against and retrieve items
//    FolderName - The folder within the mailbox to retrieve items
//    Protocol - The protocol the script should use to access the mailbox
//    Scope - The type of application permission to use
//    Permission - The type of permissions the application should have to the mailbox
//
//.EXAMPLES
// .\Demo-App.ps1 -MailboxName thanos@thejimmartin.com -FolderName Inbox -Protocol Graph -Scope AppAccessPolicy -Permission Application
// This example attempts to access the Inbox using Graph with an application scoped using Application Access Policy
//
// .\Demo-App.ps1 -MailboxName venom@thejimmartin.com -FolderName "Sent Items" -Protocol EWS -Scope RBAC -Permission Application
// This example attempts to access the Sent Items folder using EWS with an application scoped using RBAC
//
// .\Demo-App.ps1 -MailboxName ronan@thejimmartin.com -FolderName Calendar -Protocol EWS -Scope None -Permission Application
// This example attempts to access the Calendar using EWS with an application that is not scoped
//
//.NOTES
//  This script requires three separate applications registered 
// 
//
#>
<#
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
//
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="The mailbox where the folder where be renamed.")] [string] $MailboxName,
    [Parameter(Mandatory=$true, HelpMessage="The folder to be renamed.")] [string] $FolderName,
    [Parameter(Mandatory = $true, HelpMessage="Protocol the script should use to access mailbox")] [ValidateSet("Graph", "EWS")] [String]$Protocol,
    [Parameter(Mandatory = $true, HelpMessage="Application the script should use to access mailbox")] [ValidateSet("None", "RBAC", "AppAccessPolicy","Impersonation")] [String]$Scope,
    [Parameter(Mandatory = $false, HelpMessage="Application permission type of either Delegated or Application")] [ValidateSet("Delegated", "Application")] [String]$Permission="Application"
)

function Get-ApplicationOAuthToken {
    param([string]$ScriptProtocol)
    #Change the AppId, AppSecret, and TenantId to match your registered application
    switch ($Scope) {
        "None" {
            $AppId = "2f79178b-54c3-4e81-83a0-a7d16010a424"
            $AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxx"
        }
        "AppAccessPolicy" {
            $AppId = "431791f8-7511-4167-86bd-f14d69a50b9a"
            $AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
        }
        "RBAC" {
            $AppId = "02d3aaee-8dd9-4bec-b4a4-e5bf7f03c802"
            $AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
        }
    }
    $TenantId = "00001111-5be5-4438-a1d7-xxxxyyyyzzzz"
    $Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    switch($ScriptProtocol) {
        "Graph" { $Scope = "https://graph.microsoft.com/.default" }
        "EWS" { $Scope = "https://outlook.office365.com/.default" }
    }
    $Body = @{
        client_id     = $AppId
        scope         = $Scope
        client_secret = $AppSecret
        grant_type    = "client_credentials"
    }
    $TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
    #Unpack the access token
    $Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
    return $Token
}

function Get-DelegatedOAuthToken {
    param([string]$ScriptProtocol)
    #Check and install Microsoft Authentication Library module
    if(!(Get-Module -Name MSAL.PS -ListAvailable -ErrorAction Ignore)){
        try { 
            #Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module -Name MSAL.PS -Repository PSGallery -Force
        }
        catch {
            Write-Warning "Failed to install the Microsoft Authentication Library module."
            exit
        }
        try {
            Import-Module -Name MSAL.PS
        }
        catch {
            Write-Warning "Failed to import the Microsoft Authentication Library module."
        }
    }
    switch($Scope) {
        "None" {
            $ClientID = "2f79178b-54c3-4e81-83a0-a7d16010a424"
            $RedirectUri = "msal2f79178b-54c3-4e81-83a0-a7d16010a424://auth"
        }
        "AppAccessPolicy" {
            $ClientID = "431791f8-7511-4167-86bd-f14d69a50b9a"
            $RedirectUri = "msal431791f8-7511-4167-86bd-f14d69a50b9a://auth"
        }
        "RBAC" {
            $ClientID = "02d3aaee-8dd9-4bec-b4a4-e5bf7f03c802"
            $RedirectUri = "msal02d3aaee-8dd9-4bec-b4a4-e5bf7f03c802://auth"
        }
        "Impersonation" {
            $ClientID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            $RedirectUri = "ms-appx-web://Microsoft.AAD.BrokerPlugin/d3590ed6-52b3-4102-aeff-aad2292ab01c"
        }
    }
    switch($Protocol) {
        "EWS" {$ScopeUri = "https://outlook.office365.com/.default"}
        "Graph" {$ScopeUri = "https://graph.microsoft.com/.default"}
    }
    $Token = Get-MsalToken -ClientId $ClientID -RedirectUri $RedirectUri -Scopes $ScopeUri -Interactive
    return $Token.AccessToken
}

#region Disclaimer
Write-Host -ForegroundColor Yellow '//***********************************************************************'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// Copyright (c) 2018 Microsoft Corporation. All rights reserved.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR'
Write-Host -ForegroundColor Yellow '// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,'
Write-Host -ForegroundColor Yellow '// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE'
Write-Host -ForegroundColor Yellow '// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER'
Write-Host -ForegroundColor Yellow '// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,'
Write-Host -ForegroundColor Yellow '// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '//**********************************************************************​'
#endregion

#region Get OAuthToken
switch($Permission) {
    "Application" {$OAuthToken = Get-ApplicationOAuthToken -ScriptProtocol $Protocol}
    "Delegated" {$OAuthToken = Get-DelegatedOAuthToken }
}
#endregion

switch($Protocol) {
    "Graph" {
        #region Find folder
        Write-Host "Attempting to find folder with the name $FolderName..." -ForegroundColor Cyan -NoNewline
        $Headers = @{ 
            'Authorization' = "Bearer $OAuthToken" 
            'Content-type' = "application/json"
        }
        $MessageParams = @{
            "URI"         = "https://graph.microsoft.com/v1.0/users/$MailboxName/mailFolders/delta?`$select=DisplayName"
            "Headers"     = $Headers
            "Method"      = "GET"
        }
        try { $Folders = Invoke-RestMethod @Messageparams }
        catch { 
            Write-Host "FAILED" -ForegroundColor Red
            exit
        }
        foreach($folder in $Folders.value) {
            if($folder.displayName -eq $FolderName) {$FindFolder = $folder.Id}
        }
        if($FindFolder -like $null) {
            Write-Host "FAILED" -ForegroundColor Red
            Write-Warning "Unable to locate the folder $FolderName in the mailbox for $MailboxName."
            exit
        }
        Write-Host "COMPLETE"
        #endregion

        $Headers = @{ 
            'Authorization' = "Bearer $OAuthToken" 
        }
        $MessageParams = @{
            "URI"         = "https://graph.microsoft.com/v1.0/users/$MailboxName/mailFolders/$FindFolder/messages?`$top=15"
            "Headers"     = $Headers
            "Method"      = "GET"
        }
        $global:Messages = Invoke-RestMethod @Messageparams
        $global:Messages.Value | ft Subject
}
    "EWS" {
        #region LoadEwsManagedAPI
        $ewsDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        if (Test-Path $ewsDLL) {
            Import-Module $ewsDLL
        }
        else {
            Write-Warning "This script requires the EWS Managed API 1.2 or later."
            exit
        }
        #endregion

        #region EwsService
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
        ## Create Exchange Service Object
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
        $service.HttpHeaders.Clear()
        $service.UserAgent = "EwsPowerShellScript"
        $OAuthToken = "Bearer {0}" -f $OAuthToken
        $service.HttpHeaders.Add("Authorization", " $($OAuthToken)")
        $service.Url = "https://outlook.office365.com/ews/exchange.asmx"
        $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
        #endregion

        $WellKnownFolderNames = @("ArchiveDeletedItems",
            "ArchiveMsgFolderRoot",
            "ArchiveRecoverableItemsDeletions",
            "ArchiveRecoverableItemsPurges",
            "ArchiveRecoverableItemsRoot",
            "ArchiveRecoverableItemsVersions",
            "ArchiveRoot",
            "Calendar",
            "Conflicts",
            "Contacts",
            "ConversationHistory",
            "DeletedItems",
            "Drafts",
            "Inbox",
            "Journal",
            "JunkEmail",
            "LocalFailures",
            "MsgFolderRoot",
            "Notes",
            "Outbox",
            "PublicFoldersRoot",
            "QuickContacts",
            "RecipientCache",
            "RecoverableItemsDeletions",
            "RecoverableItemsPurges",
            "RecoverableItemsRoot",
            "RecoverableItemsVersions",
            "Root",
            "SearchFolders",
            "SentItems",
            "ServerFailures",
            "SyncIssues",
            "Tasks",
            "ToDoSearch",
            "VoiceMail"
        )
        $FolderCheck = $FolderName.Replace(" ","")
    
        if($WellKnownFolderNames -notcontains $FolderCheck) {
            Write-Host "Searching for $FolderName in the mailbox..." -ForegroundColor Cyan
            $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
            $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
            $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
            $SfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)
            $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)
            if ($findFolderResults.TotalCount -gt 0){ 
                foreach($folder in $findFolderResults.Folders){ 
                    $folderid = $folder.Id
                } 
            } 
            else{ 
                Write-Warning "$FolderName was not found in the mailbox for $MailboxName"  
                #$tfTargetFolder = $null  
                exit  
            }
        }
        #region ConnectToFolder
        else {
            $folderid= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderCheck,$MailboxName)
        }
        Write-Host "Connecting to the $FolderName for $MailboxName..." -ForegroundColor Cyan -NoNewline
        try { 
            $MailboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid) 
            Write-Host "COMPLETE"
        }
        catch { 
            Write-Host "FAILED" -ForegroundColor Red 
            exit
        }
        #endregion
        #region GetItems
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(15)  
        #    do{
        $fiResult = $MailboxFolder.FindItems($ivItemView)
        foreach($Item in $fiResult.Items){  
            $Item.Subject
        }
        $ivItemView.offset += $fiResult.Items.Count  
        #    }
        #    while($fiResult.MoreAvailable -eq $true)
        #endregion
    }
}