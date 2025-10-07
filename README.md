# Annoyabot

Slack bot that reminds users about pending Outlook calendar invites and lets them respond directly from Slack DMs.

## Development

Node.js >=20 and Azurite required to run the project.

```bash
npm install
npm run build
```

Start Azurite (in a separate terminal):

```bash
azurite #(or npx azurite)
```

Run the function app:

```bash
func start
```

## Deployment

This bot is deployed as a **Timer Trigger & HTTP Trigger Function** using @azure/functions.

Set up environment:

```bash
nvm install 20
nvm use 20
npm install
```

Build project:

```bash
git clone https://github.com/dump-hr/annoyabot
cp .env.example .env
npm run build
```
