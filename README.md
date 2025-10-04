Annoyabot
=========

Slack bot that reminds users about pending Outlook calendar invites and lets them respond directly from Slack DMs.

Development
-----------

Node.js >=20 and Azurite required to run the project.
`npm install  npm run build`

Start Azurite (in a separate terminal):
`azurite`

Run the function app:
`func start`

Deployment
----------

This bot is deployed as a **Timer Trigger Function** using @azure/functions.

Set up environment:
`nvm install 20  nvm use 20  npm install --global`

Build project:
`git clone https://github.com/dump-hr/annoyabot  cp .env.example .env  yarn build`
