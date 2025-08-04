import { App } from "@slack/bolt";
import { updateEventResponse } from "./outlook";
import { postThreadMessage } from "./slackMessaging";
import { getSlackUserEmail } from "./slackHelper";

const RESPONSE_MAP = {
  event_accept: {
    message: "Hvala na dolasku, vidimo se! :party-blob:",
    status: "accepted" as const,
  },
  event_tentative: {
    message:
      "All good, samo mi javi kad budeš siguran/na!:care:\nDržim ti mjesto u međuvremenu! :chair:",
    status: "tentativelyAccepted" as const,
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

  const userEmail = await getSlackUserEmail(client, body.user.id);

  if (!eventId) throw new Error("Missing event ID");
  if (!userEmail) throw new Error("Missing user email");

  try {
    await updateEventResponse(eventId, userEmail, status);
    await postThreadMessage(client, body, message);
  } catch (error) {
    console.error(`Error processing ${actionId}:`, error);
    await postThreadMessage(
      client,
      body,
      `Došlo je do greške u obradi odgovora!`
    );
  }
};

export const setupSlackInteractions = (app: App) => {
  Object.keys(RESPONSE_MAP).forEach((actionId) => {
    app.action(actionId, handleResponse.bind(null, actionId));
  });
};
