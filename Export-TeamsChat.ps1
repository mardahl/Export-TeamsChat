<#
.SYNOPSIS
Exports Microsoft Teams chat conversations to TXT, JSON, HTML, or CSV using the Microsoft Graph API.

.DESCRIPTION
Retrieves chat metadata, members, and messages for a specified Microsoft Teams chat (provided as a Teams chat URL) using Microsoft Graph (v1.0) and exports them in the chosen format. Supports non-interactive parameter input, a guided -Interactive mode, and an optional configuration template file stored next to the script.

.PARAMETER TenantId
The Microsoft Entra ID tenant ID (GUID) of your app registration.

.PARAMETER ClientId
The application (client) ID of your app registration in Microsoft Entra ID.

.PARAMETER ClientSecret
A client secret for the app registration. Used with the OAuth 2.0 client credentials grant to obtain an access token.

.PARAMETER TeamsUrl
A Microsoft Teams chat URL that contains the chat ID (for example: https://teams.microsoft.com/l/chat/...). The script automatically extracts the chat ID from the URL.

.PARAMETER ExportFormat
The output format for the export. Valid values: TXT, JSON, HTML, CSV. Default: TXT.

.PARAMETER OutputPath
Destination directory for the exported file. Default: current directory (.).

.PARAMETER ConfigFile
Creates a TeamsExportConfig.json file in the script folder with setup instructions and placeholders for TenantId, ClientId, and ClientSecret.

.PARAMETER Interactive
Runs a guided interactive setup to collect settings and start the export.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -ConfigFile
Creates the configuration template file TeamsExportConfig.json next to the script.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -Interactive
Starts the guided mode, prompts for required values, and exports the chat.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -ClientSecret "<secret>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
Authenticates with Microsoft Graph using app credentials and exports the specified chat to the default TXT format.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..." -ExportFormat HTML -OutputPath "C:\Exports"
Exports the specified chat to HTML in the given output directory. TenantId/ClientId/ClientSecret are read from TeamsExportConfig.json when present.

.OUTPUTS
String. Returns the full file path of the exported file.

.REMARKS
- Exports chat metadata, members, and messages using Microsoft Graph v1.0.
- Supports TXT, JSON, HTML, and CSV formats. HTML preserves basic message formatting.
- Uses application permissions and OAuth 2.0 client credentials flow.
- Accepts a Teams chat deep link; the script extracts the 19:...@thread.v2 or ...@unq chat ID.
- Handles pagination to retrieve all messages for large chats.

.NOTES
Author: Michael Mardahl (GitHub: https://github.com/mardahl)
Version: 1.0.0
Last Updated: 2025-09-01
LLM: ChatGPT 5 and Claude 4
Work: Consultant for hire via inciro.com
License: Prosperity Public License 3.0.0 (noncommercial + 30-day commercial trial). Commercial licensing and consulting: https://inciro.com

Requirements:
- PowerShell 7+
- Microsoft Graph application permissions with admin consent:
  Chat.Read.All (required), ChatMessage.Read.All (optional)
- The script uses client credentials flow against https://graph.microsoft.com/v1.0
Config file path: $PSScriptRoot\TeamsExportConfig.json

.LINK
https://learn.microsoft.com/graph/api/resources/chatmessage?view=graph-rest-1.0
.LINK
https://learn.microsoft.com/graph/permissions-reference
.LINK
https://github.com/mardahl
.LINK
https://prosperitylicense.com/versions/3.0.0
#>

#requires -Version 7.0

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [string]$TeamsUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("TXT", "JSON", "HTML", "CSV")]
    [string]$ExportFormat = "TXT",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",
    
    [Parameter(Mandatory = $false)]
    [switch]$ConfigFile,
    
    [Parameter(Mandatory = $false)]
    [switch]$Interactive
)

# Configuration file path
$ConfigFilePath = Join-Path $PSScriptRoot "TeamsExportConfig.json"

# Function to create/update configuration file
function New-ConfigFile {
    $config = @{
        TenantId = ""
        ClientId = ""
        ClientSecret = ""
        Instructions = @{
            Setup = @(
                "1. Go to Microsoft Entra admin center → App registrations",
                "2. Create a new app registration",
                "3. Under API permissions, add these APPLICATION permissions:",
                "   - Microsoft Graph → Chat.Read.All",
                "   - Microsoft Graph → ChatMessage.Read.All (optional)",
                "4. Click 'Grant admin consent'",
                "5. Under Certificates & secrets, create a new client secret",
                "6. Copy the Application (client) ID, Directory (tenant) ID, and Client Secret to this config"
            )
            RequiredPermissions = @(
                "Chat.Read.All",
                "ChatMessage.Read.All"
            )
        }
    }
    
    $config | ConvertTo-Json -Depth 3 | Out-File $ConfigFilePath -Encoding UTF8
    Write-Host "✅ Configuration file created at: $ConfigFilePath" -ForegroundColor Green
    Write-Host "📝 Please edit the file and add your Microsoft Entra ID app registration details" -ForegroundColor Yellow
}

# Function to load configuration
function Get-Configuration {
    if (Test-Path $ConfigFilePath) {
        try {
            return Get-Content $ConfigFilePath | ConvertFrom-Json
        }
        catch {
            Write-Error "❌ Failed to parse configuration file: $($_.Exception.Message)"
            return $null
        }
    }
    return $null
}

# Function to get access token
function Get-AccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    
    try {
        Write-Host "🔐 Authenticating with Microsoft Graph..." -ForegroundColor Cyan
        $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Host "✅ Authentication successful!" -ForegroundColor Green
        return $response.access_token
    }
    catch {
        Write-Error "❌ Authentication failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $errorBody = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorBody)
            $errorContent = $reader.ReadToEnd()
            Write-Error "Error details: $errorContent"
        }
        throw
    }
}

function Get-ChatIdFromUrl {
    param([string]$TeamsUrl)

    try {
        if ([string]::IsNullOrWhiteSpace($TeamsUrl)) { throw "Empty TeamsUrl" }

        # Decode once to turn %3A -> :, %40 -> @, etc.
        try {
            Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
            $decoded = [System.Web.HttpUtility]::UrlDecode($TeamsUrl)
        } catch { $decoded = $TeamsUrl }

        # Try several known patterns:
        $patterns = @(
            '/l/chat/(?<id>19:[^/?]+@(thread\.v2|unq))',            # /l/chat/19:...@thread.v2/...
            '/conversations/(?<id>19:[^/?]+@(thread\.v2|unq))',     # .../conversations/19:...@thread.v2?
            'chatid=(?<id>19:[^&]+@(thread\.v2|unq))',              # ...chatid=19:...@unq
            '(?<id>19:[A-Za-z0-9\-_]+@(thread\.v2|unq))'            # bare fallback
        )

        foreach ($p in $patterns) {
            $m = [regex]::Match($decoded, $p, 'IgnoreCase')
            if ($m.Success) { return $m.Groups['id'].Value }
        }

        throw "Could not extract chat ID from URL:`n$decoded"
    }
    catch {
        throw "❌ Invalid Teams URL format: $($_.Exception.Message)"
    }
}

# Function to make Graph API requests
function Invoke-GraphRequest {
    param(
        [string]$Endpoint,
        [string]$AccessToken,
        [string]$Method = "GET"
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $uri = "https://graph.microsoft.com/v1.0$Endpoint"
    
    try {
        return Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode
        $errorBody = ""
        
        if ($_.Exception.Response) {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd()
        }
        
        Write-Error "❌ Graph API request failed: $statusCode - $errorBody"
        throw
    }
}

# Function to get all messages with pagination
function Get-AllChatMessages {
    param(
        [string]$ChatId,
        [string]$AccessToken
    )
    
    $allMessages = @()
    $nextLink = "/chats/$([uri]::EscapeDataString($ChatId))/messages?`$top=50&`$orderby=createdDateTime desc"
    
    Write-Host "📥 Fetching chat messages..." -ForegroundColor Cyan
    
    do {
        $response = Invoke-GraphRequest -Endpoint $nextLink -AccessToken $AccessToken
        $allMessages += $response.value
        
        Write-Host "📨 Retrieved $($response.value.Count) messages (Total: $($allMessages.Count))" -ForegroundColor Gray
        
        # Handle pagination
        $nextLink = $null
        if ($response.'@odata.nextLink') {
            $nextLink = $response.'@odata.nextLink' -replace 'https://graph.microsoft.com/v1.0', ''
        }
    } while ($nextLink)
    
    Write-Host "✅ Total messages retrieved: $($allMessages.Count)" -ForegroundColor Green
    return $allMessages
}

# Function to strip HTML tags
function Remove-HtmlTags {
    param([string]$HtmlString)
    
    if ([string]::IsNullOrEmpty($HtmlString)) {
        return ""
    }
    
    # Simple regex to remove HTML tags
    $cleanText = $HtmlString -replace '<[^>]*>', ''
    # Decode common HTML entities
    $cleanText = $cleanText -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&' -replace '&quot;', '"'
    return $cleanText.Trim()
}

# Function to format date for display
function Format-DisplayDate {
    param([string]$DateString)
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return $DateString
    }
}

# Export functions
function Export-ToText {
    param(
        [object]$ChatData,
        [array]$Messages,
        [string]$OutputPath
    )
    
    $fileName = "teams-chat-export-$(Get-Date -Format 'yyyy-MM-dd-HHmm').txt"
    $filePath = Join-Path $OutputPath $fileName
    
    $content = @"
Microsoft Teams Chat Export
================================

Chat Information:
- Chat Type: $($ChatData.chatType)
- Created: $(Format-DisplayDate $ChatData.createdDateTime)
- Participants: $($ChatData.members.displayName -join ', ')
- Total Messages: $($Messages.Count)
- Chat ID: $($ChatData.id)

Messages:
----------

"@
    
    # Sort messages chronologically (oldest first)
    $sortedMessages = $Messages | Sort-Object createdDateTime
    
    foreach ($msg in $sortedMessages) {
        $sender = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $timestamp = Format-DisplayDate $msg.createdDateTime
        $content += "[${timestamp}] ${sender}:`n"
        
        if ($msg.messageType -eq "unknownFutureValue" -or $msg.messageType -eq "systemEventMessage") {
            $content += "   System: $($msg.eventDetail.'@odata.type' -replace '#microsoft.graph.', '')`n"
        }
        else {
            $messageContent = Remove-HtmlTags $msg.body.content
            $content += "   $messageContent`n"
        }
        $content += "`n"
    }
    
    $content += "`nExported on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    
    $content | Out-File $filePath -Encoding UTF8
    return $filePath
}

function Export-ToJSON {
    param(
        [object]$ChatData,
        [array]$Messages,
        [string]$OutputPath
    )
    
    $fileName = "teams-chat-export-$(Get-Date -Format 'yyyy-MM-dd-HHmm').json"
    $filePath = Join-Path $OutputPath $fileName
    
    $exportData = @{
        chatInfo = $ChatData
        messages = $Messages | Sort-Object createdDateTime
        exportedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        exportedBy = "PowerShell Script"
        totalMessages = $Messages.Count
    }
    
    $exportData | ConvertTo-Json -Depth 10 | Out-File $filePath -Encoding UTF8
    return $filePath
}

function Export-ToHTML {
    param(
        [object]$ChatData,
        [array]$Messages,
        [string]$OutputPath
    )
    
    $fileName = "teams-chat-export-$(Get-Date -Format 'yyyy-MM-dd-HHmm').html"
    $filePath = Join-Path $OutputPath $fileName
    
    $sortedMessages = $Messages | Sort-Object createdDateTime
    
    $messagesHtml = ""
    foreach ($msg in $sortedMessages) {
        $sender = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $timestamp = Format-DisplayDate $msg.createdDateTime
        $isSystemMessage = $msg.messageType -eq "unknownFutureValue" -or $msg.messageType -eq "systemEventMessage"
        $messageClass = if ($isSystemMessage) { "message system-message" } else { "message" }
        
        if ($isSystemMessage) {
            $content = "System: $($msg.eventDetail.'@odata.type' -replace '#microsoft.graph.', '')"
        }
        else {
            $content = $msg.body.content
        }
        
        $messagesHtml += @"
        <div class="$messageClass">
            <div class="message-header">
                <span class="sender">$sender</span>
                <span class="timestamp">$timestamp</span>
            </div>
            <div class="message-content">$content</div>
        </div>
"@
    }
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Teams Chat Export - $(Format-DisplayDate $ChatData.createdDateTime)</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .chat-info { background: #f0f8ff; padding: 20px; border-radius: 8px; margin-bottom: 25px; border-left: 5px solid #4c63d2; }
        .message { border-left: 4px solid #4c63d2; padding: 15px; margin-bottom: 15px; background: #fafafa; border-radius: 8px; }
        .message-header { font-weight: bold; margin-bottom: 8px; color: #4c63d2; display: flex; justify-content: space-between; }
        .timestamp { color: #666; font-size: 0.9em; font-weight: normal; }
        .message-content { line-height: 1.6; color: #333; }
        .system-message { background: #fff3cd; border-left-color: #ffc107; color: #856404; font-style: italic; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 0.9em; text-align: center; }
        h1 { color: #4c63d2; margin-bottom: 20px; }
        h3 { color: #333; margin-bottom: 15px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📱 Microsoft Teams Chat Export</h1>
        
        <div class="chat-info">
            <h3>Chat Information</h3>
            <p><strong>Chat Type:</strong> $($ChatData.chatType)</p>
            <p><strong>Created:</strong> $(Format-DisplayDate $ChatData.createdDateTime)</p>
            <p><strong>Participants:</strong> $($ChatData.members.displayName -join ', ')</p>
            <p><strong>Total Messages:</strong> $($Messages.Count)</p>
            <p><strong>Chat ID:</strong> <code>$($ChatData.id)</code></p>
        </div>
        
        <div class="messages">
            $messagesHtml
        </div>
        
        <div class="footer">
            <p>Exported on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') using Microsoft Graph API</p>
        </div>
    </div>
</body>
</html>
"@
    
    $html | Out-File $filePath -Encoding UTF8
    return $filePath
}

function Export-ToCSV {
    param(
        [object]$ChatData,
        [array]$Messages,
        [string]$OutputPath
    )
    
    $fileName = "teams-chat-export-$(Get-Date -Format 'yyyy-MM-dd-HHmm').csv"
    $filePath = Join-Path $OutputPath $fileName
    
    $csvData = @()
    
    # Sort messages chronologically
    $sortedMessages = $Messages | Sort-Object createdDateTime
    
    foreach ($msg in $sortedMessages) {
        $sender = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $isSystemMessage = $msg.messageType -eq "unknownFutureValue" -or $msg.messageType -eq "systemEventMessage"
        
        if ($isSystemMessage) {
            $content = "System: $($msg.eventDetail.'@odata.type' -replace '#microsoft.graph.', '')"
        }
        else {
            $content = Remove-HtmlTags $msg.body.content
        }
        
        $csvData += [PSCustomObject]@{
            Timestamp = Format-DisplayDate $msg.createdDateTime
            Sender = $sender
            MessageType = $msg.messageType
            Content = $content
            MessageId = $msg.id
        }
    }
    
    $csvData | Export-Csv $filePath -NoTypeInformation -Encoding UTF8
    return $filePath
}

# Function to get user input securely
function Get-SecureInput {
    param(
        [string]$Prompt,
        [switch]$IsSecret
    )
    
    if ($IsSecret) {
        $secureString = Read-Host $Prompt -AsSecureString
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
        $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
        return $plainText
    }
    else {
        return Read-Host $Prompt
    }
}

# Function to run interactive mode
function Start-InteractiveMode {
    Write-Host "`n🚀 Microsoft Teams Chat Exporter - Interactive Mode" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    Write-Host "`n📋 Microsoft Entra ID app registration setup required:" -ForegroundColor Yellow
    Write-Host "1. Go to Microsoft Entra admin center → App registrations"
    Write-Host "2. Create a new app registration"
    Write-Host "3. Under API permissions, add these APPLICATION permissions:"
    Write-Host "   - Microsoft Graph → Chat.Read.All"
    Write-Host "   - Microsoft Graph → ChatMessage.Read.All (optional)"
    Write-Host "4. Click 'Grant admin consent'"
    Write-Host "5. Under Certificates & secrets, create a new client secret"
    Write-Host ""
    
    # Get configuration
    $script:TenantId = Get-SecureInput "Enter your Tenant ID"
    $script:ClientId = Get-SecureInput "Enter your Client ID"
    $script:ClientSecret = Get-SecureInput "Enter your Client Secret" -IsSecret
    $script:TeamsUrl = Get-SecureInput "Enter the Teams chat URL"
    
    Write-Host "`nSelect export format:" -ForegroundColor Yellow
    Write-Host "1. TXT (Text file)"
    Write-Host "2. JSON (Structured data)"
    Write-Host "3. HTML (Web page)"
    Write-Host "4. CSV (Spreadsheet)"
    
    do {
        $formatChoice = Read-Host "Enter choice (1-4) [Default: 1]"
        if ([string]::IsNullOrEmpty($formatChoice)) { $formatChoice = "1" }
    } while ($formatChoice -notin @("1", "2", "3", "4"))
    
    $script:ExportFormat = @("TXT", "JSON", "HTML", "CSV")[[int]$formatChoice - 1]
    
    $script:OutputPath = Get-SecureInput "Enter output directory [Default: current directory]"
    if ([string]::IsNullOrEmpty($script:OutputPath)) {
        $script:OutputPath = "."
    }
}

# Main execution logic
function Start-TeamsExport {
    Write-Host "`n🗨️ Microsoft Teams Chat Exporter" -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    # Handle configuration file creation
    if ($ConfigFile) {
        New-ConfigFile
        return
    }
    
    # Load configuration from file if exists
    $config = Get-Configuration
    
    # Use parameters or config file values or interactive mode
    if ($Interactive -or (-not $TenantId -and -not $config)) {
        Start-InteractiveMode
    }
    else {
        $script:TenantId = $TenantId ?? $config.TenantId
        $script:ClientId = $ClientId ?? $config.ClientId
        $script:ClientSecret = $ClientSecret ?? $config.ClientSecret
    }
    
    # Validate required parameters
    if (-not $script:TenantId -or -not $script:ClientId -or -not $script:ClientSecret) {
        Write-Error "❌ Missing required parameters. Use -Interactive mode or provide TenantId, ClientId, and ClientSecret"
        Write-Host "`n💡 Tip: Run with -ConfigFile to create a configuration template" -ForegroundColor Yellow
        return
    }
    
    if (-not $TeamsUrl -and -not $script:TeamsUrl) {
        $script:TeamsUrl = Get-SecureInput "Enter the Teams chat URL"
    }
    elseif ($TeamsUrl) {
        $script:TeamsUrl = $TeamsUrl
    }
    
    try {
        # Extract chat ID
        Write-Host "`n🔍 Extracting chat ID from URL..." -ForegroundColor Cyan
        $chatId = Get-ChatIdFromUrl $script:TeamsUrl
        Write-Host "✅ Chat ID: $chatId" -ForegroundColor Green
        
        # Get access token
        $accessToken = Get-AccessToken -TenantId $script:TenantId -ClientId $script:ClientId -ClientSecret $script:ClientSecret
        
        # Get chat information
        Write-Host "`n📊 Retrieving chat information..." -ForegroundColor Cyan
        $chatData = Invoke-GraphRequest -Endpoint "/chats/$([uri]::EscapeDataString($chatId))" -AccessToken $accessToken
        
        # Get chat members
        $membersResponse = Invoke-GraphRequest -Endpoint "/chats/$([uri]::EscapeDataString($chatId))/members" -AccessToken $accessToken
        $chatData | Add-Member -NotePropertyName "members" -NotePropertyValue $membersResponse.value
        
        Write-Host "✅ Chat Type: $($chatData.chatType)" -ForegroundColor Green
        Write-Host "✅ Participants: $($chatData.members.displayName -join ', ')" -ForegroundColor Green
        
        # Get all messages
        $messages = Get-AllChatMessages -ChatId $chatId -AccessToken $accessToken
        
        # Export based on format
        Write-Host "`n📤 Exporting to $ExportFormat format..." -ForegroundColor Cyan
        
        $exportedFile = switch ($ExportFormat.ToUpper()) {
            "TXT" { Export-ToText -ChatData $chatData -Messages $messages -OutputPath $OutputPath }
            "JSON" { Export-ToJSON -ChatData $chatData -Messages $messages -OutputPath $OutputPath }
            "HTML" { Export-ToHTML -ChatData $chatData -Messages $messages -OutputPath $OutputPath }
            "CSV" { Export-ToCSV -ChatData $chatData -Messages $messages -OutputPath $OutputPath }
        }
        
        Write-Host "`n🎉 Export completed successfully!" -ForegroundColor Green
        Write-Host "📁 File saved: $exportedFile" -ForegroundColor Green
        Write-Host "📊 Total messages exported: $($messages.Count)" -ForegroundColor Green
        
        # Emit the exported file path to the pipeline
        Write-Output $exportedFile
        
        # Open the output directory
        if ($IsWindows) {
            Write-Host "`n💡 Opening output directory..." -ForegroundColor Yellow
            Start-Process explorer.exe -ArgumentList (Split-Path $exportedFile -Parent)
        }
        
    }
    catch {
        Write-Error "❌ Export failed: $($_.Exception.Message)"
        Write-Host "`n🔧 Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "- Verify your app registration has the correct permissions"
        Write-Host "- Ensure admin consent has been granted"
        Write-Host "- Check that the Teams URL is valid and accessible"
        Write-Host "- Confirm your client secret hasn't expired"
    }
}

# Script header
Write-Host @"

 ████████╗███████╗ █████╗ ███╗   ███╗███████╗
 ╚══██╔══╝██╔════╝██╔══██╗████╗ ████║██╔════╝
    ██║   █████╗  ███████║██╔████╔██║███████╗
    ██║   ██╔══╝  ██╔══██║██║╚██╔╝██║╚════██║
    ██║   ███████╗██║  ██║██║ ╚═╝ ██║███████║
    ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝
                                              
   ██████╗██╗  ██╗ █████╗ ████████╗           
  ██╔════╝██║  ██║██╔══██╗╚══██╔══╝           
  ██║     ███████║███████║   ██║              
  ██║     ██╔══██║██╔══██║   ██║              
  ╚██████╗██║  ██║██║  ██║   ██║              
   ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝              
                                              
  ███████╗██╗  ██╗██████╗  ██████╗ ██████╗ ████████╗███████╗██████╗ 
  ██╔════╝╚██╗██╔╝██╔══██╗██╔═══██╗██╔══██╗╚══██╔══╝██╔════╝██╔══██╗
  █████╗   ╚███╔╝ ██████╔╝██║   ██║██████╔╝   ██║   █████╗  ██████╔╝
  ██╔══╝   ██╔██╗ ██╔═══╝ ██║   ██║██╔══██╗   ██║   ██╔══╝  ██╔══██╗
  ███████╗██╔╝ ██╗██║     ╚██████╔╝██║  ██║   ██║   ███████╗██║  ██║
  ╚══════╝╚═╝  ╚═╝╚═╝      ╚═════╝ ╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Magenta

# Display help if no parameters
if (-not $PSBoundParameters.Count -and -not $Interactive) {
    Write-Host @"
📖 USAGE EXAMPLES:

1. 🔧 Create configuration file:
   .\Export-TeamsChat.ps1 -ConfigFile

2. 🖱️ Interactive mode (guided setup):
   .\Export-TeamsChat.ps1 -Interactive

3. 📝 Direct parameters:
   .\Export-TeamsChat.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret" -TeamsUrl "https://teams.microsoft.com/l/chat/..."

4. 📄 Using config file:
   .\Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..." -ExportFormat JSON

5. 📁 Custom output location:
   .\Export-TeamsChat.ps1 -TeamsUrl "..." -OutputPath "C:\Exports" -ExportFormat HTML

🎯 PARAMETERS:
   -TenantId       : Microsoft Entra ID tenant ID
   -ClientId       : Microsoft Entra ID app registration Client ID
   -ClientSecret   : Microsoft Entra ID app registration Client Secret
   -TeamsUrl       : Microsoft Teams chat URL
   -ExportFormat   : TXT, JSON, HTML, or CSV (default: TXT)
   -OutputPath     : Export directory (default: current directory)
   -ConfigFile     : Create configuration template
   -Interactive    : Run in guided interactive mode

"@ -ForegroundColor White

    $choice = Read-Host "`n🤔 Would you like to run in interactive mode? (y/N)"
    if ($choice -eq 'y' -or $choice -eq 'Y') {
        $Interactive = $true
    }
    else {
        Write-Host "👋 Run the script with -Interactive or provide the required parameters. Use -ConfigFile to create a configuration template." -ForegroundColor Yellow
        return
    }
}

# Execute main function
Start-TeamsExport