# AGENTS.md

## Project Overview
- **Project:** School Mail Priorities — a local email dashboard that groups Outlook/Microsoft 365 messages by priority and helps create summaries and draft replies.
- **Target user:** A student using Codex/AI to manage school email and project follow-ups safely.
- **My skill level:** Beginner to intermediate.
- **Stack:** Plain HTML, CSS, JavaScript, Node.js local server, Microsoft Graph optional, Outlook Desktop PowerShell automation optional.

## Commands
- **Install:** No package install required.
- **Dev:** `node server.js`
- **Start from Windows CMD:** `start-dashboard.cmd`
- **Build:** No build step.
- **Test:** `node --check server.js && node --check app.js`
- **PowerShell parse check:** `$scripts = @('scripts\read-outlook.ps1','scripts\create-outlook-draft.ps1'); foreach ($script in $scripts) { $content = Get-Content -Raw -LiteralPath $script; $null = [scriptblock]::Create($content); Write-Output "$script parses OK" }`
- **Lint:** No linter configured.

## Do
- Read existing code before modifying anything.
- Match existing patterns, naming, and style.
- Handle errors gracefully; no silent failures.
- Keep changes small and scoped to what was asked.
- Run syntax checks after code changes.
- Preserve the safety-first email workflow: summaries and draft replies are allowed, automatic sending is not.
- Keep the local server bound to `127.0.0.1` unless the user explicitly asks otherwise.
- Use draft-only behavior for Outlook replies unless the user explicitly changes scope.

## Don't
- Install new dependencies without asking.
- Delete or overwrite files without confirming.
- Hardcode secrets, API keys, passwords, tokens, or credentials.
- Store full email bodies permanently.
- Rewrite working code unless explicitly asked.
- Push, deploy, or force-push without permission.
- Make changes outside the scope of the request.
- Add automatic email sending.
- Try to bypass school IT, Microsoft admin consent, or account security restrictions.

## When Stuck
- If a task is large, break it into steps and confirm the plan first.
- If you cannot fix an error in 2 attempts, stop and explain the issue.
- If Outlook Desktop reads the wrong mailbox, inspect visible accounts/stores before changing code.
- If Microsoft Graph access is blocked, suggest school IT approval or the Outlook Desktop route.

## Testing
- Run existing checks after any code change.
- Add at least one focused test or manual verification note for new features.
- Never skip or delete checks to make things pass.
- For UI changes, verify the localhost dashboard loads.

## Git
- Small, focused commits with descriptive messages.
- Never force push.
- Do not commit `.env`, logs, API keys, or temporary debug output.

## Response Style
- Always respond with clear and concise messages.
- Use plain English when explaining to the user.
- Avoid long sentences, complex words, or long paragraphs.
