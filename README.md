# Inactivity Bot

A Microsoft Teams bot that tracks chat inactivity and sends reminders to the same conversation until someone marks the thread as done.

## What It Does

- Starts tracking when someone sends a `/track ...` command.
- Parses natural-language reminder instructions through a Power Automate endpoint.
- Stores tracked chats in memory for the running process.
- Sends periodic reminders after inactivity windows.
- Stops reminders when someone sends `done`.

## Project Structure

```text
appPackage/         Teams app manifest assets
src/index.ts        Bot runtime and command handling
src/storage.ts      In-memory chat record helpers
```

## Prerequisites

- Node.js 18+
- npm
- A valid `PA_PARSE_URL` endpoint (Power Automate flow)
- Teams/Bot Framework setup for your target environment

## Environment Variables

Create a `.env` file in the project root.

```env
PA_PARSE_URL=<your-power-automate-http-trigger-url>
PORT=3978
```

Notes:
- `PA_PARSE_URL` is required for `/track` parsing.
- `PORT` is optional and defaults to `3978`.
- Do not commit real secrets or endpoint signatures.

## Install

```bash
npm install
```

## Run Locally

Development mode (TypeScript watch):

```bash
npm run dev
```

Production build and start:

```bash
npm run build
npm start
```

## Available Scripts

- `npm run clean` - remove build output
- `npm run build` - build TypeScript with tsup into `dist/`
- `npm start` - run compiled bot from `dist/index.js`
- `npm run dev` - run source with `tsx watch` and dotenv
- `npm run start:dev` - run package entry directly with Node

## Bot Commands

Start tracking in a chat:

```text
/track in 2 hours send PR update
/track tomorrow morning follow up with support
/track in 15 minutes ping the team
```

Stop tracking in a chat:

```text
done
```

## Behavior Notes

- Tracking is per chat/conversation ID.
- Mentions from `/track` messages are preserved in reminders.
- Reminder interval currently allows 5 seconds to 30 days.
- Check loop currently runs every second (`CHECK_EVERY_MS = 1000`) for test-friendly behavior.

## Build Output

Compiled artifacts are generated in `dist/` and include type declarations.

## Deployment Notes

For Azure deployment, deploy the bot application code (not only Teams manifest files):

- Include at zip root: `package.json`, lockfile (if used), `dist/`, `node_modules/` (if your deployment strategy expects it).
- Ensure startup command resolves to `node dist/index.js`.
- Keep Azure App Service and Azure Bot configuration separate from Teams app manifest packaging.

## Current Limitations

- Tracking state is in-memory and resets on restart/redeploy.
- There is no persistent storage or retry queue yet.

