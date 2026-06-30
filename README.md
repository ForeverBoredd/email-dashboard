# School Mail Priorities

A local email priority dashboard for Outlook/Microsoft 365. It groups recent inbox messages into `Urgent`, `High`, `Normal`, and `Low`, then helps create summaries, reply drafts, and alternate reply drafts.

The app is designed for safe school-email use. It does not store your school password and it does not send emails automatically.

## Features

- Local dashboard at `http://localhost:8787`
- Demo inbox for testing without connecting an account
- Outlook Desktop connector for already-signed-in local mailboxes
- Optional Microsoft Graph sign-in path if school IT approves it
- Priority grouping with local keyword rules
- Summary, reply draft, and alternate reply actions
- OpenAI support through the local server, with user approval before selected email content is sent
- Draft-only reply creation

## Tech Stack

- Plain HTML, CSS, and JavaScript
- Node.js local server
- PowerShell scripts for Outlook Desktop automation on Windows
- Optional Microsoft Graph API
- Optional OpenAI API

## Requirements

- Windows
- Node.js 18 or newer
- Outlook Desktop installed if using the Outlook Desktop connector
- A mailbox already signed in and loaded in classic Outlook Desktop

No npm packages are required.

## Run Locally

From Command Prompt:

```cmd
cd /d "C:\Users\maxda\OneDrive - Bromsgrove School\Documents\school-email-dashboard"
start-dashboard.cmd
```

Or with Node directly:

```powershell
cd "C:\Users\maxda\OneDrive - Bromsgrove School\Documents\school-email-dashboard"
node server.js
```

Then open:

[http://localhost:8787](http://localhost:8787)

Keep the command window open while using the dashboard.

## Outlook Desktop Setup

Use this route if school Microsoft app registration is blocked.

1. Open classic Outlook Desktop.
2. Make sure the target mailbox is signed in and fully loaded.
3. Start the dashboard locally.
4. Open Settings and confirm the `Outlook Desktop mailbox` field.
5. Click `Outlook Desktop`.

The app reads recent inbox messages through Outlook Desktop and can save reply drafts. It does not send messages.

## Microsoft Graph Setup

Use this only if school IT approves a Microsoft Entra app registration.

1. Create a Microsoft Entra app registration.
2. Add this single-page application redirect URI:

```text
http://localhost:8787
```

3. Request delegated Microsoft Graph permissions:
   - `User.Read`
   - `Mail.Read`
   - `Mail.ReadWrite`
4. Paste the app/client ID into Settings.
5. Sign in from the dashboard.

If school policy blocks unapproved apps, use the Outlook Desktop route instead.

## AI Setup

The dashboard works without OpenAI by using local template summaries and replies.

To use OpenAI through the local server:

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
node server.js
```

You can also paste an OpenAI key into Settings for the current browser session. That key is not saved permanently by the dashboard.

Before AI actions, the app asks for approval unless you switch to local-only mode. Only the selected email content is sent.

## Safety Notes

- Do not put passwords or school credentials into this project.
- Do not commit `.env` files or API keys.
- Reply actions create drafts only.
- Full email bodies are not permanently stored by the app.
- The server binds to `127.0.0.1` for local-only access.

## Checks

Run JavaScript syntax checks:

```powershell
node --check server.js
node --check app.js
```

Check PowerShell scripts parse:

```powershell
$scripts = @('scripts\read-outlook.ps1','scripts\create-outlook-draft.ps1')
foreach ($script in $scripts) {
  $content = Get-Content -Raw -LiteralPath $script
  $null = [scriptblock]::Create($content)
  Write-Output "$script parses OK"
}
```

## Project Structure

```text
.
├── AGENTS.md
├── README.md
├── app.js
├── index.html
├── server.js
├── styles.css
├── start-dashboard.cmd
├── start-dashboard.ps1
└── scripts
    ├── create-outlook-draft.ps1
    └── read-outlook.ps1
```

## Deployment

Netlify is not the best fit for the full version of this app because the main functionality depends on a local Node server, Windows PowerShell, and Outlook Desktop automation.

Netlify could host a static demo-only version, but it cannot run the local Outlook Desktop connector. For the real dashboard, run it locally.
