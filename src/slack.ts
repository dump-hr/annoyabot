import { App } from "@slack/bolt";
import { setupSlackInteractions } from "./slackInteractions";

let slackApp: App | null = null;

export function getSlackApp(): App {
  if (!slackApp) {
    slackApp = new App({
      token: process.env.SLACK_BOT_TOKEN,
      signingSecret: process.env.SLACK_SIGNING_SECRET,
    });

    setupSlackInteractions(slackApp);
  }

  return slackApp;
}
