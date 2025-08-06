import { App } from "@slack/bolt";
import { updateEventResponse } from "./outlook";
import { ResponseStatus } from "./types";
import { getSlackUserEmail } from "./slackHelper";
import { postThreadMessage } from "./slackMessaging";

const RESPONSE_MAP = {
  event_accept: {
    message: "Hvala na dolasku, vidimo se! :party-blob:",
    status: ResponseStatus.ACCEPTED,
  },
  event_tentative: {
    message:
      "Okej, samo javi kad budeš siguran/na!:care:\nDržim ti mjesto u međuvremenu! :chair:",
    status: ResponseStatus.TENTATIVE,
  },
  event_decline: {
    message:
      "Žao mi je što ne možeš doći :sadge:\nVidimo se drugi put! :cvjetic:",
    status: ResponseStatus.DECLINED,
  },
};

const handleResponse = async (actionId: string, { ack, body, client }) => {
  await ack();
  const { message, status } = RESPONSE_MAP[actionId];
  const { value: iCalUID } = body.actions[0];

  try {
    const userEmail = await getSlackUserEmail(client, body.user.id);

    if (!iCalUID) throw new Error("Nedostaje ID dogadjaja");
    if (!userEmail) throw new Error("Nije moguce dohvatiti email korisnika");

    await updateEventResponse(iCalUID, userEmail, status);
    await postThreadMessage(client, body, message);
  } catch (error) {
    console.error(`Error processing ${actionId}:`, error);
    await postThreadMessage(
      client,
      body,
      `Došlo je do greške u obradi odgovora!\nMolim te pokušaj ponovno kasnije`
    );
  }
};

export const setupSlackInteractions = (app: App) => {
  Object.keys(RESPONSE_MAP).forEach((actionId) => {
    app.action(actionId, handleResponse.bind(null, actionId));
  });
};
