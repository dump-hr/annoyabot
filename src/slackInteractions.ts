import { App } from "@slack/bolt";

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

  async function postThreadMessage(client: any, body: any, text: string) {
    try {
      await client.chat.postMessage({
        channel: body.container.channel_id,
        thread_ts: body.container.message_ts,
        text,
      });
    } catch (error) {
      console.error("Greška pri slanju poruke:", error);
    }
  }
}
