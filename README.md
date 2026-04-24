# Export-TeamsChat.ps1 — Microsoft Teams chat exporter

Export Microsoft Teams chat conversations to TXT, JSON, HTML, or CSV using the Microsoft Graph API.

## Features
- Export formats: TXT, JSON, HTML, CSV
- Downloads inline hosted images and file attachments, then rewrites message references to local relative paths
- **Delegated auth — device code flow** — enter a short code in your browser; no admin consent required
- **Delegated auth — browser sign-in (PKCE)** — opens a real browser window for direct sign-in; useful when device code flow is blocked by Conditional Access policies
- **App-only auth (client credentials)** — tenant-wide access for admin/automation scenarios
- Guided interactive mode (`-Interactive`) with saved-value reuse, validation, and a final confirmation step
- Config template generator (`-ConfigFile`)
- Extracts chat ID from Teams deep links automatically
- Handles pagination to retrieve all messages
- Compatible with **PowerShell 5.1** (Windows PowerShell) and **PowerShell 7+**

## What this script is useful for
- AI analysis and summarization: Export JSON/CSV/HTML to feed LLMs, RAG pipelines, or analytics tooling.
- Archiving and record-keeping: Store chat history as flat files for backup or team knowledge bases.
- Legal/reference copies: Produce human-readable transcripts for internal reviews and audits.

Note/Disclaimer:
- Not legal advice.
- Not an eDiscovery solution; no cryptographic checksums, digital signatures, audit trail, or chain of custody.
- Files can be altered. For tamper-evident or evidentiary needs, consider Microsoft 365 Purview eDiscovery or implement hashing/signing and controlled storage.

## Authentication modes

| | Delegated (Interactive) | App-only (Config/Parameters) |
|---|---|---|
| **Who it's for** | Personal use; export your own chats | Admins and automation |
| **Permission required** | `Chat.Read` (delegated) | `Chat.Read.All` (application) |
| **Admin consent required** | No | Yes |
| **Chats accessible** | Only chats the signed-in user is a member of | Any chat in the tenant |
| **Credentials needed** | Tenant ID + Client ID (optional) | Tenant ID + Client ID + Client Secret |
| **How to invoke** | `-Interactive` (no `ClientSecret`) | Provide `ClientSecret` via params or config file |

### Delegated sign-in flows

Two flows are available for delegated (user) sign-in:

| Flow | Switch | How it works | When to use |
|---|---|---|---|
| **Device code** | *(default)* | Displays a short code; open `https://microsoft.com/devicelogin`, enter the code and sign in | Works on most tenants |
| **Browser sign-in (PKCE)** | `-BrowserAuth` | Opens your default browser directly; no code to copy | Use when device code is blocked by Conditional Access or legacy auth policies |

**Security note:** Delegated mode is the safer choice for personal use — it only accesses what your account can access and requires no secrets to be stored. App-only credentials grant tenant-wide access and should be protected accordingly (restrict the app registration, secure the secret, limit who can read the config file).

## Requirements

### Delegated mode (personal use)
- PowerShell 5.1+ (Windows PowerShell or PowerShell 7+)
- A Microsoft Entra tenant ID is recommended, but optional when using delegated auth. If omitted, the script signs in via the `common` endpoint and detects the tenant from the returned token.
- A Client ID with the `Chat.Read` **delegated** permission — or use the default Microsoft Graph Command Line Tools app (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), which requires no app registration
- **Browser sign-in only:** the redirect URI `http://localhost:<port>` (any port 8400–8420) must be registered on the app. The default Microsoft Graph Command Line Tools app supports loopback redirect URIs for public clients.

### App-only mode (admin/automation)
- PowerShell 5.1+ (Windows PowerShell or PowerShell 7+)
- App registration in Microsoft Entra ID with:
  - `Chat.Read.All` **application** permission + admin consent
  - `ChatMessage.Read.All` application permission (optional)
  - A client secret

## Quick start — Delegated mode (personal use)

No app registration needed. Choose between device code and browser sign-in.

### Option A — Device code flow (default)

1. Run the script in interactive mode:
   ```powershell
   pwsh ./Export-TeamsChat.ps1 -Interactive
   ```
2. When prompted, enter your **Tenant ID** (find it in Entra admin center → Overview → Tenant ID).
3. When prompted for a **Client ID**, press Enter to use the default (Microsoft Graph Command Line Tools app).
4. Choose **"Device code"** at the sign-in method prompt.
5. A device code will be displayed. Open the URL shown (e.g. `https://microsoft.com/devicelogin`), enter the code, and sign in with your Microsoft 365 account.
6. Paste your Teams chat URL when prompted.
7. Choose an export format and output location.

### Option B — Browser sign-in / PKCE (when device code is blocked)

Use this if your organisation's Conditional Access policies block device code flow.

1. Run the script with the `-BrowserAuth` flag:
   ```powershell
   pwsh ./Export-TeamsChat.ps1 -Interactive -BrowserAuth
   ```
   Or select **"Browser"** at the sign-in method prompt in standard `-Interactive` mode.
2. Enter your **Tenant ID** and optionally a **Client ID** (defaults to Microsoft Graph Command Line Tools).
3. Your default browser opens automatically — sign in with your Microsoft 365 account.
4. After signing in, return to the terminal to continue the export.

> The default Client ID is the Microsoft Graph Command Line Tools app (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). This is a well-known Microsoft first-party app. If your organisation has disabled it, you will need to register your own app with the `Chat.Read` delegated permission and add `http://localhost` as a redirect URI for browser sign-in.

## Setup — App-only mode (admin/automation)

1. Go to **Microsoft Entra admin center** → App registrations
2. Create a new app registration
3. Under **API permissions**, add **Application** permissions:
   - Microsoft Graph → `Chat.Read.All`
   - Microsoft Graph → `ChatMessage.Read.All` (optional)
4. Grant **admin consent**
5. Under **Certificates & secrets**, create a new client secret
6. Note the **Directory (tenant) ID**, **Application (client) ID**, and **Client Secret value**

## Configuration file

`TeamsExportConfig.json` can store either app-only settings or delegated defaults so you don't have to re-enter them each time.

Create a config template:
```powershell
pwsh ./Export-TeamsChat.ps1 -ConfigFile
```

The file is saved next to the script. Fields:

| Field | Description |
|---|---|
| `AuthMode` | Set to `AppOnly` for client credentials flow or `Delegated` to keep delegated defaults in the file |
| `TenantId` | Your Entra tenant ID |
| `ClientId` | Your app registration's client ID |
| `ClientSecret` | Your client secret (app-only mode only; leave blank for delegated mode) |

> Do not commit this file to source control. Add `TeamsExportConfig.json` to your `.gitignore`.

## Usage examples

**Interactive — delegated auth, let the script ask you which sign-in flow to use:**
```powershell
pwsh ./Export-TeamsChat.ps1 -Interactive
```

**Interactive — force browser sign-in (skips the device code / browser prompt):**
```powershell
pwsh ./Export-TeamsChat.ps1 -Interactive -BrowserAuth
```

**Non-interactive — delegated browser sign-in:**
```powershell
pwsh ./Export-TeamsChat.ps1 -TenantId "<tenantId>" -BrowserAuth -TeamsUrl "https://teams.microsoft.com/l/chat/..."
```

**Non-interactive — delegated device code sign-in:**
```powershell
pwsh ./Export-TeamsChat.ps1 -TenantId "<tenantId>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
```

**Direct parameters — app-only auth:**
```powershell
pwsh ./Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -ClientSecret "<secret>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
```

**Using a config file — app-only auth:**
```powershell
pwsh ./Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..."
```

**Custom output format and path:**
```powershell
pwsh ./Export-TeamsChat.ps1 -TeamsUrl "https://teams.microsoft.com/l/chat/..." -ExportFormat HTML -OutputPath "./exports"
```

## Copying a Teams chat link (Teams for work or school)
- From the chat list:
  - Right-click (or Control-click on macOS) the chat in the left sidebar.
  - Choose "Copy link" or "Copy link to chat".
- From inside a chat:
  - Click the chat name/header, then More options (…).
  - Choose "Copy link to chat".
- Copying a specific message link:
  - Hover the message → More options (… ) → Copy link.
  - The script can extract the chat ID from both chat and message links.

Tip: Links typically look like https://teams.microsoft.com/l/chat/... and contain a chat ID like 19:…@thread.v2 or …@unq, which the script can parse automatically.

## Output
- The script writes the exported file to disk and also outputs the full file path to the pipeline (stdout).
- When inline images or file attachments are found, the script downloads them into a sibling folder named `{export-filename-without-extension}-assets`.
- Example: `teams-chat-export-2026-04-24-1530.html` produces `teams-chat-export-2026-04-24-1530-assets/` in the same directory.
- HTML output rewrites `src` and `href` references to those local relative asset paths.
- JSON output preserves the rewritten message HTML and localized attachment URLs.
- TXT and CSV output include localized asset paths alongside the message text when assets were downloaded.
- If an asset download fails, the script logs a warning and continues the export.

## Notes
- Graph API base: `https://graph.microsoft.com/v1.0`
- The script parses common Teams chat URL patterns to extract the chat ID (e.g., `19:...@thread.v2` or `...@unq`).
- In delegated mode, only chats where the signed-in user is a participant are accessible — this is a Graph API constraint, not a script limitation.
- Browser sign-in (PKCE) starts a temporary HTTP listener on a loopback port (8400–8420). No data leaves your machine through this listener.
- Hosted inline images are downloaded from `GET /chats/{chatId}/messages/{messageId}/hostedContents/{hostedContentId}/$value` using the same authenticated Graph token the export already uses.

## License
- Prosperity Public License 3.0.0 (noncommercial + 30-day commercial trial)
- Full text: `LICENSE` or https://prosperitylicense.com/versions/3.0.0
- For commercial licensing, please contact the author.

## Author
- Michael Mardahl — GitHub: https://github.com/mardahl

## Support & contributions
- Issues and contributions are welcome via GitHub.
