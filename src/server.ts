import { App } from "@slack/bolt";
import { setupSlackInteractions } from "./slackInteractions";

export function startServer() {
  const slackApp = new App({
    token: process.env.SLACK_BOT_TOKEN,
    signingSecret: process.env.SLACK_SIGNING_SECRET,
    socketMode: true,
    appToken: process.env.SLACK_APP_TOKEN,
  });

  setupSlackInteractions(slackApp);

  (async () => {
    await slackApp.start();
    console.log("App pokrenut");
  })();

  return slackApp;
}
