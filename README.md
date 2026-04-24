# Export-TeamsChat.ps1 — Microsoft Teams chat exporter

Export Microsoft Teams chat conversations to TXT, JSON, HTML, or CSV using the Microsoft Graph API.

## Features
- Export formats: TXT, JSON, HTML, CSV
- **Delegated auth (device code flow)** — sign in interactively, no admin consent required
- **App-only auth (client credentials)** — tenant-wide access for admin/automation scenarios
- Guided interactive mode (`-Interactive`) with saved-value reuse, validation, and a final confirmation step
- Config template generator (`-ConfigFile`)
- Extracts chat ID from Teams deep links automatically
- Handles pagination to retrieve all messages
- Cross-platform with PowerShell 7+

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

**Security note:** Delegated mode is the safer choice for personal use — it only accesses what your account can access and requires no secrets to be stored. App-only credentials grant tenant-wide access and should be protected accordingly (restrict the app registration, secure the secret, limit who can read the config file).

## Requirements

### Delegated mode (personal use)
- PowerShell 7+
- A Microsoft Entra tenant ID
- A Client ID with the `Chat.Read` **delegated** permission — or use the default Microsoft Graph Command Line Tools app (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), which requires no app registration

### App-only mode (admin/automation)
- PowerShell 7+
- App registration in Microsoft Entra ID with:
  - `Chat.Read.All` **application** permission + admin consent
  - `ChatMessage.Read.All` application permission (optional)
  - A client secret

## Quick start — Delegated mode (personal use)

No app registration needed. Uses device code flow to sign you in interactively.

1. Run the script in interactive mode:
   ```powershell
   pwsh ./Export-TeamsChat.ps1 -Interactive
   ```
2. When prompted, enter your **Tenant ID** (find it in Entra admin center → Overview → Tenant ID).
3. When prompted for a **Client ID**, press Enter to use the default (Microsoft Graph Command Line Tools app).
4. A device code will be displayed. Open the URL shown (e.g. `https://microsoft.com/devicelogin`), enter the code, and sign in with your Microsoft 365 account.
5. Paste your Teams chat URL when prompted.
6. Choose an export format and output location.

> The default Client ID is the Microsoft Graph Command Line Tools app (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). This is a well-known Microsoft first-party app. If your organisation has disabled it, you will need to register your own app with the `Chat.Read` delegated permission.

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

`TeamsExportConfig.json` is used to store app-only credentials so you don't have to pass them as parameters each time.

Create a config template:
```powershell
pwsh ./Export-TeamsChat.ps1 -ConfigFile
```

The file is saved next to the script. Fields:

| Field | Description |
|---|---|
| `TenantId` | Your Entra tenant ID |
| `ClientId` | Your app registration's client ID |
| `ClientSecret` | Your client secret (app-only mode) |
| `AuthMode` | Set to `AppOnly` for client credentials flow |

> Do not commit this file to source control. Add `TeamsExportConfig.json` to your `.gitignore`.

## Usage examples

**Interactive — delegated auth (personal use, no secrets required):**
```powershell
pwsh ./Export-TeamsChat.ps1 -Interactive
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

## Notes
- Graph API base: `https://graph.microsoft.com/v1.0`
- The script parses common Teams chat URL patterns to extract the chat ID (e.g., `19:...@thread.v2` or `...@unq`).
- In delegated mode, only chats where the signed-in user is a participant are accessible — this is a Graph API constraint, not a script limitation.

## License
- Prosperity Public License 3.0.0 (noncommercial + 30-day commercial trial)
- Full text: `LICENSE` or https://prosperitylicense.com/versions/3.0.0
- For commercial licensing, please contact the author.

## Author
- Michael Mardahl — GitHub: https://github.com/mardahl

## Support & contributions
- Issues and contributions are welcome via GitHub.
