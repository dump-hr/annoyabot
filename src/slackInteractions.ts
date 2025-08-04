import { App } from "@slack/bolt";
import { updateEventResponse } from "./outlook";

const RESPONSE_MAP = {
  event_accept: {
    message: "Hvala na dolasku, vidimo se! :party-blob:",
    status: "accepted" as const,
  },
  event_tentative: {
    message:
      "All good, samo mi javi kad budeš siguran/na!:care:\nDržim ti mjesto u međuvremenu! :chair:",
    status: "tentative" as const,
  },
  event_decline: {
    message:
      "Žao mi je što ne možeš doći :sadge:\nVidimo se drugi put! :cvjetic:",
    status: "declined" as const,
  },
};

const handleResponse = async (actionId: string, { ack, body, client }) => {
  await ack();
  const { message, status } = RESPONSE_MAP[actionId];
  const { value: eventId } = body.actions[0];
  const { email: userEmail } = body.user;

  const threadTs =
    body.container?.thread_ts || body.message?.ts || body.container?.message_ts;

  try {
    await updateEventResponse(eventId, userEmail, status);
    await client.chat.postMessage({
      channel: body.channel.id,
      text: message,
      thread_ts: threadTs,
    });
  } catch (error) {
    console.error(`Error processing ${actionId}:`, error);
    await client.chat.postMessage({
      channel: body.channel.id,
      text: `Došlo je do greške u obradi odgovora!`,
      thread_ts: threadTs,
    });
  }
};

export const setupSlackInteractions = (app: App) => {
  Object.keys(RESPONSE_MAP).forEach((actionId) => {
    app.action(actionId, handleResponse.bind(null, actionId));
  });
};
