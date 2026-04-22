# MS Teams Bot Setup Notes and Working Rules

## Goal
Build and run a Microsoft Teams bot on Azure App Service and Azure Bot, then package it as a Teams app.

## Architecture
There are 3 separate pieces and they must not be confused:

1. **Azure App Service**
   Hosts the actual bot application code.

2. **Azure Bot**
   Exposes the bot through Bot Framework and points to the bot messaging endpoint.

3. **Teams app manifest**
   Used only to package/install the bot into Microsoft Teams.
   It does **not** host code and does **not** make the bot run.

---

## Working deployment flow

### App Service
Use **Linux App Service**.

### Bot code deployment
Deploy actual bot code, not manifest folder.

Deployment zip should contain at root:
- `package.json`
- `package-lock.json` if present
- `dist/`
- `node_modules/`

Do **not** zip the parent folder.

### Startup
`package.json` uses:

```json
"start": "node dist/index.js"
```
