# Gekonny Outlook Add-in — Subject Builder

In-Outlook panel that formats email subjects for Monday automation.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Add-in manifest — deployed to M365, points at this host |
| `taskpane.html/.css/.js` | The Subject Builder panel (ribbon button) |
| `commands.html/.js` | Hidden runtime — hosts the OnMessageSend Smart Alert |
| `assets/` | Add-in icons |

## Before deploying — one config edit

Open `taskpane.js`, set `CONFIG.PROJECTS_ENDPOINT` to the helper-flow GET URL
(Power Automate `GetActiveProjects`, full URL incl. `&sig=`).

## Hosting

All files are served from GitHub Pages at:
`https://zakhargalan.in/gekonny-outlook-addin/`

The manifest references that domain. If the domain changes, update every URL
in `manifest.xml`.

## Deployment to the team (M365 admin)

1. M365 admin center → Settings → Integrated apps → Upload custom apps
2. Upload `manifest.xml`
3. Assign to the team / everyone
4. Outlook shows the **Gekonny → Subject Builder** button within ~24h

## Smart Alert behaviour

On Send, the add-in checks the subject. If it is not formatted as
`[CODE] [TYPE] …` or `RFP GC/SUB - …`, it shows a soft warning
("no task will be created"). The user can still send.
