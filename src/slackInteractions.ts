import { App } from "@slack/bolt";
import { postThreadMessage } from "./slackMessaging";

export function setupSlackInteractions(slackApp: App) {
  slackApp.action("event_accept", async ({ body, ack, client }) => {
    await ack();
    await postThreadMessage(client, body, "Hvala na dolasku! :party-blob:");
  });

  slackApp.action("event_tentative", async ({ ack, body, client }) => {
    await ack();
    await postThreadMessage(
      client,
      body,
      'Zabilježen je tvoj "Možda" :thinking_face:'
    );
  });

  slackApp.action("event_decline", async ({ ack, body, client }) => {
    await ack();
    await postThreadMessage(client, body, "Razumijem! :sadge:");
  });
}
