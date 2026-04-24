<#
.SYNOPSIS
Exports Microsoft Teams chat conversations to TXT, JSON, HTML, or CSV using the Microsoft Graph API.

.DESCRIPTION
Retrieves chat metadata, members, and messages for a specified Microsoft Teams chat (provided as a Teams chat URL) using Microsoft Graph (v1.0) and exports them in the chosen format. Supports non-interactive parameter input, a guided -Interactive mode, and an optional configuration template file stored next to the script.

Compatible with PowerShell 5.1 and PowerShell 7+.

Two authentication modes are supported:

  Delegated (default for -Interactive / no-params):
    Supports two sign-in flows — device code or interactive browser (PKCE).
    Only requires TenantId and ClientId. No client secret is needed.
    The default ClientId is the well-known Microsoft Graph Command Line Tools app
    (14d82eec-204b-4c2f-b7e8-296a70dab67e), which has Chat.Read pre-consented
    in most tenants.
    Device code flow: displays a short code to enter at https://microsoft.com/devicelogin.
    Browser flow (-BrowserAuth): opens a browser window for direct sign-in —
      useful when device code flow is blocked by Conditional Access policies.

  App-only (used when ClientSecret is supplied):
    Uses the OAuth 2.0 client credentials flow. Requires TenantId, ClientId,
    and ClientSecret with Chat.Read.All application permission granted by an admin.

.PARAMETER TenantId
The Microsoft Entra ID tenant ID (GUID). Optional in delegated mode — if omitted or set to 'common',
the script authenticates via the common endpoint and auto-detects the tenant ID from the returned token.

.PARAMETER ClientId
The application (client) ID of your app registration in Microsoft Entra ID.
In delegated mode, defaults to the Microsoft Graph Command Line Tools app
(14d82eec-204b-4c2f-b7e8-296a70dab67e) if omitted.

.PARAMETER ClientSecret
A client secret for the app registration. When provided, the script uses the
OAuth 2.0 client credentials (app-only) flow. Omit to use delegated device code flow.

.PARAMETER TeamsUrl
A Microsoft Teams chat URL that contains the chat ID (for example:
https://teams.microsoft.com/l/chat/...). The script automatically extracts the
chat ID from the URL.

.PARAMETER ExportFormat
The output format for the export. Valid values: TXT, JSON, HTML, CSV. Default: TXT.

.PARAMETER OutputPath
Destination directory for the exported file. Default: current directory (.).

.PARAMETER ConfigFile
Creates a TeamsExportConfig.json file in the script folder with setup instructions
and placeholders for TenantId, ClientId, ClientSecret, and AuthMode.

.PARAMETER Interactive
Runs a guided interactive setup using delegated authentication.
Only TenantId and ClientId are required — no client secret.
By default uses device code flow; add -BrowserAuth to use interactive browser sign-in.

.PARAMETER BrowserAuth
When combined with -Interactive or -Delegated, uses the OAuth 2.0 Authorization Code flow
with PKCE (opens a real browser window) instead of the device code flow.
Useful on tenants that block device code / legacy authentication.
A local HTTP listener is started on a loopback port (8400–8420) to receive the redirect.
Requires the redirect URI http://localhost:<port> to be registered on the app registration,
or use the default Microsoft Graph Command Line Tools client ID which supports loopback URIs.

.PARAMETER Delegated
Forces delegated authentication even when running non-interactively
(i.e. when TenantId and ClientId are passed as parameters but ClientSecret is not).
Pair with -BrowserAuth to use the interactive browser flow instead of device code.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -ConfigFile
Creates the configuration template file TeamsExportConfig.json next to the script.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -Interactive
Starts the guided mode with delegated sign-in. You will be prompted to choose between
device code flow and interactive browser sign-in.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -Interactive -BrowserAuth
Starts the guided mode and uses browser-based interactive sign-in (no device code).

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -BrowserAuth -TeamsUrl "https://teams.microsoft.com/l/chat/..."
Authenticates via interactive browser sign-in (PKCE) and exports the specified chat to TXT.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
Authenticates via device code flow (delegated) and exports the specified chat to TXT.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -ClientSecret "<secret>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
Authenticates using app-only client credentials and exports the specified chat to TXT.

.EXAMPLE
PS> .\Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..." -ExportFormat HTML -OutputPath "C:\Exports"
Exports the specified chat to HTML in the given output directory. Auth credentials
are read from TeamsExportConfig.json when present.

.OUTPUTS
String. Returns the full file path of the exported file.

.REMARKS
- Exports chat metadata, members, and messages using Microsoft Graph v1.0.
- Supports TXT, JSON, HTML, and CSV formats. HTML preserves basic message formatting.
- Delegated mode uses Chat.Read (no admin consent required for most tenants).
- App-only mode uses Chat.Read.All and requires admin consent.
- Accepts a Teams chat deep link; the script extracts the 19:...@thread.v2 or ...@unq chat ID.
- Handles pagination to retrieve all messages for large chats.
- Compatible with PowerShell 5.1 and PowerShell 7+.

.NOTES
Author: Michael Mardahl (GitHub: https://github.com/mardahl)
Version: 1.2.1
Last Updated: 2026-04-24
LLM: ChatGPT 5 and Claude 4
Work: Consultant for hire via inciro.com
License: Prosperity Public License 3.0.0 (noncommercial + 30-day commercial trial). Commercial licensing and consulting: https://inciro.com

Requirements:
- PowerShell 5.1+ (both Windows PowerShell and PowerShell 7+ are supported)
- Delegated mode: Chat.Read delegated permission (pre-consented on the default app ID in most tenants)
- App-only mode: Chat.Read.All application permission with admin consent
- Browser auth (-BrowserAuth): redirect URI http://localhost:<port> must be registered on the app,
  or use the default Microsoft Graph Command Line Tools client ID
- The script uses Microsoft Graph v1.0 at https://graph.microsoft.com/v1.0
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

#requires -Version 5.1

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
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$BrowserAuth,

    [Parameter(Mandatory = $false)]
    [switch]$Delegated
)

# Default ClientId for delegated auth (Microsoft Graph Command Line Tools enterprise app)
$script:DefaultDelegatedClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# Configuration file path
$ConfigFilePath = Join-Path $PSScriptRoot "TeamsExportConfig.json"

# ---------------------------------------------------------------------------
# Configuration helpers
# ---------------------------------------------------------------------------

function New-ConfigFile {
    $config = [ordered]@{
        AuthMode     = "AppOnly"
        TenantId     = ""
        ClientId     = ""
        ClientSecret = ""
        Instructions = [ordered]@{
            Overview          = @(
                "This configuration file supports two authentication modes:",
                "",
                "  AppOnly (default in this file):",
                "    Requires TenantId, ClientId, and ClientSecret.",
                "    Uses the OAuth 2.0 client credentials flow (app-only).",
                "    Requires Chat.Read.All application permission with admin consent.",
                "",
                "  Delegated (for interactive / device code sign-in):",
                "    Requires TenantId and optionally ClientId.",
                "    Uses the OAuth 2.0 device code flow — no secret needed.",
                "    Requires Chat.Read delegated permission (pre-consented on the",
                "    default Microsoft Graph Command Line Tools app in most tenants).",
                "    Set AuthMode to 'Delegated' and leave ClientSecret blank to use this mode."
            )
            AppOnlySetup      = @(
                "1. Go to Microsoft Entra admin center → App registrations",
                "2. Create a new app registration",
                "3. Under API permissions, add these APPLICATION permissions:",
                "   - Microsoft Graph → Chat.Read.All",
                "   - Microsoft Graph → ChatMessage.Read.All (optional)",
                "4. Click 'Grant admin consent'",
                "5. Under Certificates & secrets, create a new client secret",
                "6. Copy the Application (client) ID, Directory (tenant) ID, and Client Secret into this file",
                "7. Set AuthMode to 'AppOnly'"
            )
            DelegatedSetup    = @(
                "1. Leave ClientSecret blank (or omit it)",
                "2. Set AuthMode to 'Delegated'",
                "3. Set TenantId to your Directory (tenant) ID",
                "4. Optionally set ClientId — if blank the script uses the well-known",
                "   Microsoft Graph Command Line Tools app (14d82eec-204b-4c2f-b7e8-296a70dab67e)",
                "5. Run the script; you will be prompted to sign in via a browser"
            )
            RequiredPermissions = [ordered]@{
                AppOnly   = @("Chat.Read.All", "ChatMessage.Read.All (optional)")
                Delegated = @("Chat.Read")
            }
        }
    }

    $config | ConvertTo-Json -Depth 5 | Out-File $ConfigFilePath -Encoding UTF8
    Write-Host "✅ Configuration file created at: $ConfigFilePath" -ForegroundColor Green
    Write-Host "📝 Please edit the file and add your Microsoft Entra ID details" -ForegroundColor Yellow
}

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

# ---------------------------------------------------------------------------
# Authentication: app-only (client credentials)
# ---------------------------------------------------------------------------

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
        Write-Host "🔐 Authenticating with Microsoft Graph (app-only)..." -ForegroundColor Cyan
        $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Host "✅ Authentication successful!" -ForegroundColor Green
        return $response.access_token
    }
    catch {
        Write-Error "❌ Authentication failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $errorContent = if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $_.ErrorDetails.Message
            } elseif ($_.Exception.Response | Get-Member -Name GetResponseStream -MemberType Method -ErrorAction SilentlyContinue) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $reader.ReadToEnd()
            } else { "" }
            if ($errorContent) { Write-Error "Error details: $errorContent" }
        }
        throw
    }
}

# ---------------------------------------------------------------------------
# Authentication: delegated (device code flow)
# ---------------------------------------------------------------------------

function Get-DelegatedAccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId
    )

    $deviceCodeUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
    $tokenUrl      = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope         = "https://graph.microsoft.com/Chat.Read offline_access"

    # Step 1 — request device code
    try {
        Write-Host "🔐 Requesting device code from Microsoft..." -ForegroundColor Cyan
        $dcResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -ContentType "application/x-www-form-urlencoded" -Body @{
            client_id = $ClientId
            scope     = $scope
        }
    }
    catch {
        Write-Error "❌ Device code request failed: $($_.Exception.Message)"
        throw
    }

    # Step 2 — instruct the user
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Sign-in required" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. Open a browser and go to:" -ForegroundColor White
    Write-Host "     $($dcResponse.verification_uri)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  2. Enter the code:" -ForegroundColor White
    Write-Host "     $($dcResponse.user_code)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  3. Sign in with your Microsoft 365 account." -ForegroundColor White
    Write-Host ""
    Write-Host "  Waiting for sign-in (expires in $($dcResponse.expires_in)s)..." -ForegroundColor Gray
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    # Step 3 — poll for token
    $interval   = if ($null -ne $dcResponse.interval)   { [int]$dcResponse.interval }   else { 5 }
    $expiresSec = if ($null -ne $dcResponse.expires_in) { [int]$dcResponse.expires_in } else { 900 }
    $deadline   = (Get-Date).AddSeconds($expiresSec)
    $deviceCode = $dcResponse.device_code

    $pollBody = @{
        client_id   = $ClientId
        device_code = $deviceCode
        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
    }

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval

        try {
            $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -ContentType "application/x-www-form-urlencoded" -Body $pollBody
            Write-Host "✅ Sign-in successful!" -ForegroundColor Green
            return $tokenResponse.access_token
        }
        catch {
            # Parse the error from the response body.
            # PowerShell 7+ surfaces the body in $_.ErrorDetails.Message; PS 5.x
            # requires reading the response stream directly.
            $rawError = $null
            try {
                $errorBody = if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $_.ErrorDetails.Message
                } elseif ($_.Exception.Response -and ($_.Exception.Response | Get-Member -Name GetResponseStream -MemberType Method -ErrorAction SilentlyContinue)) {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $reader.ReadToEnd()
                } else { "" }
                $rawError = $errorBody | ConvertFrom-Json
            }
            catch { <# ignore parse errors #> }

            $errorCode = if ($rawError -and $rawError.error) { $rawError.error } else { "unknown" }

            switch ($errorCode) {
                "authorization_pending" {
                    # Normal — user hasn't signed in yet; keep polling
                    Write-Host "." -NoNewline -ForegroundColor Gray
                }
                "slow_down" {
                    # Server asked us to slow down
                    $interval += 5
                    Write-Host "." -NoNewline -ForegroundColor Gray
                }
                "authorization_declined" {
                    Write-Host ""
                    throw "❌ The user declined the sign-in request."
                }
                "expired_token" {
                    Write-Host ""
                    throw "❌ The device code has expired. Please run the script again."
                }
                default {
                    Write-Host ""
                    $detail = if ($rawError -and $rawError.error_description) { $rawError.error_description } else { $_.Exception.Message }
                    throw "❌ Token request failed ($errorCode): $detail"
                }
            }
        }
    }

    Write-Host ""
    throw "❌ Sign-in timed out. Please run the script again and complete sign-in within the time limit."
}

# ---------------------------------------------------------------------------
# Authentication: delegated (interactive browser / authorization code + PKCE)
# Compatible with PowerShell 5.1 and PowerShell 7+
# ---------------------------------------------------------------------------

function Get-InteractiveBrowserToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [int]$TimeoutSeconds = 300
    )

    $scope    = "https://graph.microsoft.com/Chat.Read offline_access"
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    # --- PKCE: generate a cryptographically random code_verifier ---
    $rng           = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $verifierBytes = New-Object byte[] 32
    $rng.GetBytes($verifierBytes)
    $rng.Dispose()
    $codeVerifier = [Convert]::ToBase64String($verifierBytes) -replace '\+', '-' -replace '/', '_' -replace '=', ''

    # Derive code_challenge = BASE64URL( SHA256( ASCII(code_verifier) ) )
    $sha256         = New-Object System.Security.Cryptography.SHA256Managed
    $challengeBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($codeVerifier))
    $sha256.Dispose()
    $codeChallenge = [Convert]::ToBase64String($challengeBytes) -replace '\+', '-' -replace '/', '_' -replace '=', ''

    # --- Find an available loopback TCP port (8400–8420) ---
    $listener = $null
    $port     = 8400
    while ($port -le 8420) {
        try {
            $candidate = New-Object System.Net.HttpListener
            $candidate.Prefixes.Add("http://localhost:$port/")
            $candidate.Start()
            $listener = $candidate
            break
        }
        catch {
            if ($null -ne $candidate) { $candidate.Close() }
            $port++
        }
    }

    if (-not $listener) {
        throw "❌ Could not start a local HTTP listener on ports 8400–8420. Free one of those ports and try again."
    }

    $redirectUri = "http://localhost:$port"

    # --- Build authorization URL ---
    $queryParts = @(
        "client_id=$([uri]::EscapeDataString($ClientId))",
        "response_type=code",
        "redirect_uri=$([uri]::EscapeDataString($redirectUri))",
        "scope=$([uri]::EscapeDataString($scope))",
        "code_challenge=$codeChallenge",
        "code_challenge_method=S256",
        "response_mode=query"
    )
    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?" + ($queryParts -join '&')

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Browser sign-in" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Opening your browser for Microsoft sign-in..." -ForegroundColor White
    Write-Host "  If the browser does not open automatically, visit:" -ForegroundColor White
    Write-Host "  $authUrl" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Waiting for sign-in (timeout: ${TimeoutSeconds}s)..." -ForegroundColor Gray
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    # --- Open the default browser (PS5.1 on Windows + PS7 cross-platform) ---
    $onWindows = if ($null -ne $IsWindows) { $IsWindows } else { $true }
    $onMacOS   = if ($null -ne $IsMacOS)   { $IsMacOS }   else { $false }

    try {
        if ($onWindows) {
            Start-Process $authUrl
        } elseif ($onMacOS) {
            Start-Process "open" -ArgumentList $authUrl
        } else {
            Start-Process "xdg-open" -ArgumentList $authUrl
        }
    }
    catch {
        Write-Host "⚠️  Could not open browser automatically. Please open the URL above manually." -ForegroundColor Yellow
    }

    # --- Wait for the OAuth redirect callback with timeout ---
    $asyncResult = $listener.BeginGetContext($null, $null)
    $signaled    = $asyncResult.AsyncWaitHandle.WaitOne([System.TimeSpan]::FromSeconds($TimeoutSeconds))

    if (-not $signaled) {
        $listener.Stop()
        throw "❌ Browser sign-in timed out after ${TimeoutSeconds} seconds. Please run the script again."
    }

    $context  = $listener.EndGetContext($asyncResult)
    $request  = $context.Request
    $response = $context.Response

    # Return a friendly page to the browser
    $successHtml   = "<html><body style='font-family:sans-serif;text-align:center;padding:40px'><h2>&#x2705; Sign-in complete!</h2><p>You can close this browser tab and return to your terminal.</p></body></html>"
    $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($successHtml)
    $response.ContentType       = "text/html; charset=utf-8"
    $response.ContentLength64   = $responseBytes.Length
    $response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
    $response.Close()
    $listener.Stop()

    # --- Extract authorization code or error from the redirect URL ---
    $code      = $request.QueryString["code"]
    $authError = $request.QueryString["error"]

    if ($authError) {
        $errDesc = $request.QueryString["error_description"]
        throw "❌ Authorization failed: $authError — $errDesc"
    }

    if (-not $code) {
        throw "❌ No authorization code received from Microsoft. Ensure the redirect URI '$redirectUri' is registered on your app registration."
    }

    # --- Exchange the authorization code for an access token ---
    $tokenBody = @{
        client_id     = $ClientId
        code          = $code
        redirect_uri  = $redirectUri
        grant_type    = "authorization_code"
        code_verifier = $codeVerifier
        scope         = $scope
    }

    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -ContentType "application/x-www-form-urlencoded" -Body $tokenBody
        Write-Host "✅ Browser sign-in successful!" -ForegroundColor Green
        return $tokenResponse.access_token
    }
    catch {
        $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "❌ Token exchange failed: $detail"
    }
}


function Get-TenantIdFromToken {
    # Decodes a JWT access token and returns the 'tid' (tenant ID) claim.
    # Works on PowerShell 5.1 and 7+. Returns $null on any failure.
    param([string]$AccessToken)
    try {
        $parts = $AccessToken.Split('.')
        if ($parts.Count -lt 2) { return $null }
        $payload = $parts[1]
        # Base64url → standard base64
        $padded  = $payload.PadRight($payload.Length + (4 - $payload.Length % 4) % 4, '=')
        $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($padded))
        $claims  = $decoded | ConvertFrom-Json
        return $claims.tid
    }
    catch { return $null }
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
        $errorBody  = ""

        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorBody = $_.ErrorDetails.Message
        } elseif ($_.Exception.Response -and ($_.Exception.Response | Get-Member -Name GetResponseStream -MemberType Method -ErrorAction SilentlyContinue)) {
            $stream    = $_.Exception.Response.GetResponseStream()
            $reader    = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd()
        }

        Write-Error "❌ Graph API request failed: $statusCode - $errorBody"
        throw
    }
}

function Get-AllChatMessages {
    param(
        [string]$ChatId,
        [string]$AccessToken
    )

    $allMessages = @()
    $nextLink    = "/chats/$([uri]::EscapeDataString($ChatId))/messages?`$top=50&`$orderby=createdDateTime desc"

    Write-Host "📥 Fetching chat messages..." -ForegroundColor Cyan

    do {
        $response     = Invoke-GraphRequest -Endpoint $nextLink -AccessToken $AccessToken
        $allMessages += $response.value

        Write-Host "📨 Retrieved $($response.value.Count) messages (Total: $($allMessages.Count))" -ForegroundColor Gray

        $nextLink = $null
        if ($response.'@odata.nextLink') {
            $nextLink = $response.'@odata.nextLink' -replace 'https://graph.microsoft.com/v1.0', ''
        }
    } while ($nextLink)

    Write-Host "✅ Total messages retrieved: $($allMessages.Count)" -ForegroundColor Green
    return $allMessages
}

# ---------------------------------------------------------------------------
# Text utilities (unchanged)
# ---------------------------------------------------------------------------

function Remove-HtmlTags {
    param([string]$HtmlString)

    if ([string]::IsNullOrEmpty($HtmlString)) { return "" }

    $cleanText = $HtmlString -replace '<[^>]*>', ''
    $cleanText = $cleanText -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&' -replace '&quot;', '"'
    return $cleanText.Trim()
}

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

# ---------------------------------------------------------------------------
# Export functions (unchanged)
# ---------------------------------------------------------------------------

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

    $sortedMessages = $Messages | Sort-Object createdDateTime

    foreach ($msg in $sortedMessages) {
        $sender    = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $timestamp = Format-DisplayDate $msg.createdDateTime
        $content  += "[${timestamp}] ${sender}:`n"

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
        chatInfo       = $ChatData
        messages       = $Messages | Sort-Object createdDateTime
        exportedAt     = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        exportedBy     = "PowerShell Script"
        totalMessages  = $Messages.Count
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
        $sender          = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $timestamp       = Format-DisplayDate $msg.createdDateTime
        $isSystemMessage = $msg.messageType -eq "unknownFutureValue" -or $msg.messageType -eq "systemEventMessage"
        $messageClass    = if ($isSystemMessage) { "message system-message" } else { "message" }

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

    $sortedMessages = $Messages | Sort-Object createdDateTime

    foreach ($msg in $sortedMessages) {
        $sender          = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "System" }
        $isSystemMessage = $msg.messageType -eq "unknownFutureValue" -or $msg.messageType -eq "systemEventMessage"

        if ($isSystemMessage) {
            $content = "System: $($msg.eventDetail.'@odata.type' -replace '#microsoft.graph.', '')"
        }
        else {
            $content = Remove-HtmlTags $msg.body.content
        }

        $csvData += [PSCustomObject]@{
            Timestamp   = Format-DisplayDate $msg.createdDateTime
            Sender      = $sender
            MessageType = $msg.messageType
            Content     = $content
            MessageId   = $msg.id
        }
    }

    $csvData | Export-Csv $filePath -NoTypeInformation -Encoding UTF8
    return $filePath
}

# ---------------------------------------------------------------------------
# Input helpers (unchanged)
# ---------------------------------------------------------------------------

function Get-SecureInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue,
        [switch]$IsSecret,
        [switch]$Required,
        [switch]$HasSavedValue
    )

    $displayPrompt = $Prompt
    if ($IsSecret -and $HasSavedValue) {
        $displayPrompt = "$Prompt [Press Enter to keep saved value]"
    }
    elseif (-not $IsSecret -and -not [string]::IsNullOrWhiteSpace($DefaultValue)) {
        $displayPrompt = "$Prompt [$DefaultValue]"
    }

    do {
        if ($IsSecret) {
            $secureString = Read-Host $displayPrompt -AsSecureString
            $ptr          = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
            $plainText    = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            $value = $plainText
        }
        else {
            $value = Read-Host $displayPrompt
        }

        if ([string]::IsNullOrWhiteSpace($value)) {
            $value = $DefaultValue
        }

        if (-not $Required -or -not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }

        Write-Host "⚠️ This value is required." -ForegroundColor Yellow
    } while ($true)
}

function Get-ChoiceInput {
    param(
        [string]$Prompt,
        [array]$Options,
        [string]$DefaultKey,
        [string]$DefaultValue
    )

    if (-not [string]::IsNullOrWhiteSpace($DefaultValue) -and [string]::IsNullOrWhiteSpace($DefaultKey)) {
        $defaultOption = $Options | Where-Object { $_.Value -eq $DefaultValue } | Select-Object -First 1
        if ($defaultOption) {
            $DefaultKey = $defaultOption.Key
        }
    }

    foreach ($option in $Options) {
        $defaultMarker = if ($option.Key -eq $DefaultKey) { " (default)" } else { "" }
        Write-Host "$($option.Key). $($option.Label)$defaultMarker"
    }

    do {
        $choice = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($choice)) {
            $choice = $DefaultKey
        }

        $selectedOption = $Options | Where-Object { $_.Key -eq $choice } | Select-Object -First 1
        if ($selectedOption) {
            return $selectedOption.Value
        }

        Write-Host "⚠️ Enter one of: $((($Options | ForEach-Object { $_.Key }) -join ', '))" -ForegroundColor Yellow
    } while ($true)
}

function Get-Confirmation {
    param(
        [string]$Prompt,
        [bool]$Default = $true
    )

    $suffix = if ($Default) { "[Y/n]" } else { "[y/N]" }

    do {
        $response = Read-Host "$Prompt $suffix"
        if ([string]::IsNullOrWhiteSpace($response)) {
            return $Default
        }

        switch -Regex ($response.Trim()) {
            '^(y|yes)$' { return $true }
            '^(n|no)$'  { return $false }
            default     { Write-Host "⚠️ Please answer y or n." -ForegroundColor Yellow }
        }
    } while ($true)
}

function Resolve-OutputPath {
    param([string]$OutputPath)

    if ([string]::IsNullOrWhiteSpace($OutputPath)) { return "." }

    if (Test-Path $OutputPath) { return $OutputPath }

    if (Get-Confirmation "Output folder '$OutputPath' does not exist. Create it?" -Default $true) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        return $OutputPath
    }

    Write-Host "ℹ️ Keeping the current directory instead." -ForegroundColor Yellow
    return "."
}

# ---------------------------------------------------------------------------
# Interactive mode — delegated auth (device code or browser)
# ---------------------------------------------------------------------------

function Start-InteractiveMode {
    param([object]$Config)

    $script:InteractiveCancelled = $false

    Write-Host "`n🚀 Microsoft Teams Chat Exporter - Interactive Mode" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan

    # Resolve saved values (params > config)
    $savedTenantId = if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $TenantId
    } elseif ($Config -and -not [string]::IsNullOrWhiteSpace($Config.TenantId)) {
        $Config.TenantId
    } else {
        $null
    }

    $savedClientId = if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $ClientId
    } elseif ($Config -and -not [string]::IsNullOrWhiteSpace($Config.ClientId)) {
        $Config.ClientId
    } else {
        $script:DefaultDelegatedClientId
    }

    $savedTeamsUrl = if (-not [string]::IsNullOrWhiteSpace($TeamsUrl)) { $TeamsUrl } else { $null }

    Write-Host "`nThis guided mode uses delegated authentication — no client secret required." -ForegroundColor Gray
    Write-Host "You will sign in with your Microsoft 365 account." -ForegroundColor Gray

    # --- Choose sign-in method (skip prompt when -BrowserAuth was passed explicitly) ---
    if ($BrowserAuth) {
        $script:UseBrowserAuth = $true
    } else {
        Write-Host "`n🔐 Choose sign-in method" -ForegroundColor Yellow
        Write-Host "  Device code flow works on most tenants; use Browser sign-in if device code" -ForegroundColor Gray
        Write-Host "  is blocked by your organisation's Conditional Access policies." -ForegroundColor Gray
        Write-Host ""
        $authMethodChoice = Get-ChoiceInput -Prompt "Sign-in method" -DefaultKey "1" -Options @(
            @{ Key = "1"; Label = "Device code  — display a short code to enter in your browser"; Value = "DeviceCode" },
            @{ Key = "2"; Label = "Browser      — open a sign-in window directly in your browser"; Value = "Browser" }
        )
        $script:UseBrowserAuth = ($authMethodChoice -eq "Browser")
    }

    Write-Host "`n🔐 Sign-in configuration" -ForegroundColor Yellow

    # Tenant ID — optional in delegated mode.
    # Using 'common' lets Microsoft route the sign-in to the correct tenant automatically;
    # the real tenant ID is then extracted from the returned access token.
    $tenantIdDefault = if (-not [string]::IsNullOrWhiteSpace($savedTenantId)) { $savedTenantId } else { "common" }
    Write-Host "  (Leave blank or press Enter to use 'common' — tenant ID will be detected from your session)" -ForegroundColor Gray
    $script:TenantId = Get-SecureInput "Tenant ID" -DefaultValue $tenantIdDefault

    # Client ID (default = well-known MS Graph Command Line Tools app)
    $clientIdPrompt  = "Client ID [Microsoft Graph Command Line Tools (default)]"
    $enteredClientId = Get-SecureInput $clientIdPrompt -DefaultValue $savedClientId
    $script:ClientId = if ([string]::IsNullOrWhiteSpace($enteredClientId)) { $script:DefaultDelegatedClientId } else { $enteredClientId }

    # Authenticate immediately so the token is ready before we ask for more inputs
    Write-Host ""
    if ($script:UseBrowserAuth) {
        $script:AccessToken = Get-InteractiveBrowserToken -TenantId $script:TenantId -ClientId $script:ClientId
    } else {
        $script:AccessToken = Get-DelegatedAccessToken -TenantId $script:TenantId -ClientId $script:ClientId
    }

    # Auto-detect the real tenant ID from the JWT when 'common' was used or no ID was supplied
    if ([string]::IsNullOrWhiteSpace($script:TenantId) -or $script:TenantId -eq "common") {
        $detectedTenantId = Get-TenantIdFromToken -AccessToken $script:AccessToken
        if ($detectedTenantId) {
            $script:TenantId = $detectedTenantId
            Write-Host "ℹ️ Tenant ID detected from session: $detectedTenantId" -ForegroundColor Gray
        }
    }
    Write-Host ""

    Write-Host "`n💬 Chat selection" -ForegroundColor Yellow
    do {
        $script:TeamsUrl = Get-SecureInput "Teams chat URL" -DefaultValue $savedTeamsUrl -Required
        try {
            $previewChatId = Get-ChatIdFromUrl $script:TeamsUrl
            Write-Host "✅ Chat link looks valid. Chat ID: $previewChatId" -ForegroundColor Green
            break
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    } while ($true)

    Write-Host "`n📤 Export settings" -ForegroundColor Yellow
    $script:ExportFormat = Get-ChoiceInput -Prompt "Choose export format" -DefaultValue $ExportFormat -Options @(
        @{ Key = "1"; Label = "TXT  - Plain text transcript"; Value = "TXT" },
        @{ Key = "2"; Label = "JSON - Structured data";       Value = "JSON" },
        @{ Key = "3"; Label = "HTML - Readable web page";     Value = "HTML" },
        @{ Key = "4"; Label = "CSV  - Spreadsheet-friendly";  Value = "CSV" }
    )

    $script:OutputPath = Resolve-OutputPath (Get-SecureInput "Output directory" -DefaultValue $OutputPath)

    $authMethodLabel  = if ($script:UseBrowserAuth) { "Delegated — browser sign-in (PKCE)" } else { "Delegated — device code flow" }
    Write-Host "`n📝 Summary" -ForegroundColor Yellow
    Write-Host "Auth mode     : $authMethodLabel"
    Write-Host "Tenant ID     : $script:TenantId"
    Write-Host "Client ID     : $script:ClientId"
    Write-Host "Teams chat URL: $script:TeamsUrl"
    Write-Host "Export format : $script:ExportFormat"
    Write-Host "Output folder : $script:OutputPath"

    if (-not (Get-Confirmation "Start export now?" -Default $true)) {
        $script:InteractiveCancelled = $true
        Write-Host "ℹ️ Export cancelled before any API calls were made." -ForegroundColor Yellow
        return
    }
}

# ---------------------------------------------------------------------------
# Main execution logic
# ---------------------------------------------------------------------------

function Start-TeamsExport {
    Write-Host "`n🗨️ Microsoft Teams Chat Exporter" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor Cyan

    # Handle configuration file creation
    if ($ConfigFile) {
        New-ConfigFile
        return
    }

    # Load configuration from file if it exists
    $config = Get-Configuration

    # Determine whether to run interactive mode
    if ($Interactive -or (-not $TenantId -and -not $config)) {
        Start-InteractiveMode -Config $config
        if ($script:InteractiveCancelled) { return }
    }
    else {
        # Resolve credentials: parameters take precedence over config file values.
        # Use [string]::IsNullOrEmpty() so that empty-string template values ("") fall
        # back to the config correctly — the ?? operator only coalesces $null, not "".
        $script:TenantId     = if (-not [string]::IsNullOrEmpty($TenantId))     { $TenantId }     elseif ($config) { $config.TenantId }     else { $null }
        $script:ClientId     = if (-not [string]::IsNullOrEmpty($ClientId))     { $ClientId }     elseif ($config) { $config.ClientId }     else { $null }
        $script:ClientSecret = if (-not [string]::IsNullOrEmpty($ClientSecret)) { $ClientSecret } elseif ($config) { $config.ClientSecret } else { $null }
    }

    if (-not $script:ExportFormat) { $script:ExportFormat = $ExportFormat }
    if (-not $script:OutputPath)   { $script:OutputPath   = $OutputPath   }

    # Determine auth mode
    # Priority: explicit $Delegated switch > presence of ClientSecret > config AuthMode
    $useAppOnly = $false
    if ($script:ClientSecret) {
        $useAppOnly = $true
    } elseif ($Delegated) {
        $useAppOnly = $false
    } elseif ($config -and $config.AuthMode -eq "AppOnly" -and $script:ClientSecret) {
        $useAppOnly = $true
    }

    # Validate required parameters
    if ($useAppOnly) {
        if (-not $script:TenantId -or -not $script:ClientId -or -not $script:ClientSecret) {
            Write-Error "❌ App-only mode requires TenantId, ClientId, and ClientSecret."
            Write-Host "`n💡 Tips:" -ForegroundColor Yellow
            Write-Host "   - Run with -Interactive for delegated sign-in (no secret needed)"
            Write-Host "   - Run with -ConfigFile to create a configuration template"
            return
        }
    } else {
        # Delegated mode: ClientId is optional (defaults to the well-known app); TenantId defaults to 'common'
        if ([string]::IsNullOrWhiteSpace($script:ClientId)) {
            $script:ClientId = $script:DefaultDelegatedClientId
        }
        if ([string]::IsNullOrWhiteSpace($script:TenantId)) {
            $script:TenantId = "common"
        }
    }

    # Resolve Teams URL
    if (-not $TeamsUrl -and -not $script:TeamsUrl) {
        $script:TeamsUrl = Get-SecureInput "Enter the Teams chat URL"
    } elseif ($TeamsUrl) {
        $script:TeamsUrl = $TeamsUrl
    }

    try {
        # Extract chat ID
        Write-Host "`n🔍 Extracting chat ID from URL..." -ForegroundColor Cyan
        $chatId = Get-ChatIdFromUrl $script:TeamsUrl
        Write-Host "✅ Chat ID: $chatId" -ForegroundColor Green

        # Obtain access token (skip re-auth if already set by interactive mode)
        if (-not $script:AccessToken) {
            if ($useAppOnly) {
                $script:AccessToken = Get-AccessToken -TenantId $script:TenantId -ClientId $script:ClientId -ClientSecret $script:ClientSecret
            } elseif ($BrowserAuth -or $script:UseBrowserAuth) {
                $script:AccessToken = Get-InteractiveBrowserToken -TenantId $script:TenantId -ClientId $script:ClientId
            } else {
                $script:AccessToken = Get-DelegatedAccessToken -TenantId $script:TenantId -ClientId $script:ClientId
            }

            # Auto-detect the real tenant ID from the JWT when 'common' was used
            if (-not $useAppOnly -and ($script:TenantId -eq "common")) {
                $detectedTenantId = Get-TenantIdFromToken -AccessToken $script:AccessToken
                if ($detectedTenantId) {
                    $script:TenantId = $detectedTenantId
                    Write-Host "ℹ️ Tenant ID detected from session: $detectedTenantId" -ForegroundColor Gray
                }
            }
        }

        $accessToken = $script:AccessToken

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
        Write-Host "`n📤 Exporting to $script:ExportFormat format..." -ForegroundColor Cyan

        $exportedFile = switch ($script:ExportFormat.ToUpper()) {
            "TXT"  { Export-ToText  -ChatData $chatData -Messages $messages -OutputPath $script:OutputPath }
            "JSON" { Export-ToJSON  -ChatData $chatData -Messages $messages -OutputPath $script:OutputPath }
            "HTML" { Export-ToHTML  -ChatData $chatData -Messages $messages -OutputPath $script:OutputPath }
            "CSV"  { Export-ToCSV   -ChatData $chatData -Messages $messages -OutputPath $script:OutputPath }
        }

        Write-Host "`n🎉 Export completed successfully!" -ForegroundColor Green
        Write-Host "📁 File saved: $exportedFile" -ForegroundColor Green
        Write-Host "📊 Total messages exported: $($messages.Count)" -ForegroundColor Green

        # Emit the exported file path to the pipeline
        Write-Output $exportedFile

        # Open the output directory on Windows
        $onWin = if ($null -ne $IsWindows) { $IsWindows } else { $true }
        if ($onWin) {
            Write-Host "`n💡 Opening output directory..." -ForegroundColor Yellow
            Start-Process explorer.exe -ArgumentList (Split-Path $exportedFile -Parent)
        }
    }
    catch {
        Write-Error "❌ Export failed: $($_.Exception.Message)"
        Write-Host "`n🔧 Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "- Delegated mode: ensure you signed in with an account that has access to this chat"
        Write-Host "- Delegated mode: verify Chat.Read is consented for the app (ClientId)"
        Write-Host "- App-only mode: verify Chat.Read.All application permission is granted with admin consent"
        Write-Host "- App-only mode: confirm your client secret hasn't expired"
        Write-Host "- Check that the Teams URL is valid and the chat is accessible"
        Write-Host "- Ensure TenantId is correct for your organization"
    }
}

# ---------------------------------------------------------------------------
# Script entry point
# ---------------------------------------------------------------------------

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

# Display help if no parameters were supplied
if (-not $PSBoundParameters.Count -and -not $Interactive) {
    Write-Host @"
📖 USAGE EXAMPLES:

1. 🔧 Create configuration file:
   .\Export-TeamsChat.ps1 -ConfigFile

2. 🖱️ Interactive mode — guided delegated sign-in (choose device code or browser):
   .\Export-TeamsChat.ps1 -Interactive

3. 🌐 Interactive mode — force browser sign-in (skips the device code prompt):
   .\Export-TeamsChat.ps1 -Interactive -BrowserAuth

4. 🔑 Delegated (device code) — non-interactive:
   .\Export-TeamsChat.ps1 -TenantId "your-tenant-id" -TeamsUrl "https://teams.microsoft.com/l/chat/..."

5. 🌐 Delegated (browser sign-in) — non-interactive:
   .\Export-TeamsChat.ps1 -TenantId "your-tenant-id" -BrowserAuth -TeamsUrl "https://teams.microsoft.com/l/chat/..."

6. 🏢 App-only (client credentials):
   .\Export-TeamsChat.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret" -TeamsUrl "https://teams.microsoft.com/l/chat/..."

7. 📄 Using config file:
   .\Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..." -ExportFormat JSON

8. 📁 Custom output location:
   .\Export-TeamsChat.ps1 -TeamsUrl "..." -OutputPath "C:\Exports" -ExportFormat HTML

🎯 PARAMETERS:
   -TenantId       : Microsoft Entra ID tenant ID
   -ClientId       : App registration Client ID (delegated default: Microsoft Graph Command Line Tools)
   -ClientSecret   : App registration Client Secret (omit to use delegated auth)
   -TeamsUrl       : Microsoft Teams chat URL
   -ExportFormat   : TXT, JSON, HTML, or CSV (default: TXT)
   -OutputPath     : Export directory (default: current directory)
   -ConfigFile     : Create configuration template
   -Interactive    : Run in guided interactive mode (delegated sign-in)
   -BrowserAuth    : Use interactive browser sign-in (PKCE) instead of device code flow
   -Delegated      : Force delegated auth in non-interactive mode

🔐 AUTH MODES:
   Delegated  — Default for -Interactive and when no ClientSecret is given.
                Signs in as a user. Requires only TenantId (+ClientId optional).
                Uses Chat.Read delegated scope (no admin consent required in most tenants).
                Two flows available:
                  • Device code  — display a code to enter in the browser (default)
                  • Browser PKCE — opens a browser window directly (-BrowserAuth)
                    Use this if device code flow is blocked by Conditional Access.
   App-only   — Used when ClientSecret is provided.
                Uses client credentials flow. Requires Chat.Read.All with admin consent.

"@ -ForegroundColor White

    Write-Host "Choose what to do next:" -ForegroundColor Yellow
    $startupChoice = Get-ChoiceInput -Prompt "Enter choice" -DefaultKey "1" -Options @(
        @{ Key = "1"; Label = "Start guided export (delegated sign-in)"; Value = "Interactive" },
        @{ Key = "2"; Label = "Create config template";                  Value = "ConfigFile" },
        @{ Key = "3"; Label = "Exit";                                    Value = "Exit" }
    )

    switch ($startupChoice) {
        "Interactive" { $Interactive = $true }
        "ConfigFile"  { $ConfigFile  = $true }
        default {
            Write-Host "👋 Exiting. Run the script with -Interactive or provide the required parameters." -ForegroundColor Yellow
            return
        }
    }
}

# Execute main function
Start-TeamsExport
