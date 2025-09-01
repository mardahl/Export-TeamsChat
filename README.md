# Export-TeamsChat.ps1 — Microsoft Teams chat exporter

Export Microsoft Teams chat conversations to TXT, JSON, HTML, or CSV using the Microsoft Graph API.

## Features
- Export formats: TXT, JSON, HTML, CSV
- Guided interactive mode (`-Interactive`)
- Config template generator (`-ConfigFile`)
- Extracts chat ID from Teams deep links automatically
- Handles pagination to retrieve all messages
- Cross-platform with PowerShell 7+

## Requirements
- PowerShell 7+
- Microsoft Graph application permissions with admin consent:
  - `Chat.Read.All` (required)
  - `ChatMessage.Read.All` (optional)
- App registration in Microsoft Entra ID (client credentials flow)

## Setup (Microsoft Entra admin center)
1. Go to Microsoft Entra admin center → App registrations
2. Create a new app registration
3. Under API permissions, add APPLICATION permissions:
   - Microsoft Graph → `Chat.Read.All`
   - Microsoft Graph → `ChatMessage.Read.All` (optional)
4. Grant admin consent
5. Under Certificates & secrets, create a new client secret
6. Note the Directory (tenant) ID, Application (client) ID, and Client Secret

## Configuration file
- A template can be created with:
  ```powershell
  pwsh ./Export-TeamsChat.ps1 -ConfigFile
  ```
- Location: `TeamsExportConfig.json` (next to the script)
- Fields: `TenantId`, `ClientId`, `ClientSecret`

## Usage
- Interactive guided mode:
  ```powershell
  pwsh ./Export-TeamsChat.ps1 -Interactive
  ```
- Direct parameters:
  ```powershell
  pwsh ./Export-TeamsChat.ps1 -TenantId "<tenantId>" -ClientId "<clientId>" -ClientSecret "<secret>" -TeamsUrl "https://teams.microsoft.com/l/chat/..."
  ```
- Choose output format and path:
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

## License
- Prosperity Public License 3.0.0 (noncommercial + 30-day commercial trial)
- Full text: `LICENSE` or https://prosperitylicense.com/versions/3.0.0
- For commercial licensing, please contact the author.

## Author
- Michael Mardahl — GitHub: https://github.com/mardahl

## Support & contributions
- Issues and contributions are welcome via GitHub.
