import { updateEventResponse } from "./outlook";
import { getSlackUserEmail } from "./slackHelper";
import { postThreadMessage } from "./slackMessaging";
import { ResponseStatus } from "./types";
import { App } from "@slack/bolt";

const RESPONSE_MAP = {
  event_accept: {
    message: "Hvala na dolasku, vidimo se! :party-blob:",
    status: ResponseStatus.ACCEPTED,
  },
  event_tentative: {
    message: "Okej, držim ti mjesto u međuvremenu! :chair:",
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
  const { value: iCalUId } = body.actions[0];

  try {
    const userEmail = await getSlackUserEmail(client, body.user.id);
    if (!userEmail) throw new Error("Nije moguce dohvatiti email korisnika");

    const updateResult = await updateEventResponse(iCalUId, userEmail, status);

    if (!updateResult.success && updateResult.errorType === "EVENT_NOT_FOUND") {
      await postThreadMessage(
        client,
        body,
        "Nisam našao event. Za vraćanje eventa slijedi korake: :face_with_monocle:\n" +
          "1. Otvori <https://outlook.office.com/mail/deleteditems/|Deleted Items> :incoming_envelope:\n" +
          "2. Pronađi email s eventom :satellite_antenna:\n" +
          '3. Vrati ga u Inbox i klikni "Prihvati":rocket:"\n' +
          "_Važno: Na mobitelu se možda neće prikazati opcija “Prihvati”. Otvori desktop verziju._"
      );
      return;
    }

    await postThreadMessage(client, body, message);
  } catch (error) {
    console.error(`Error processing ${actionId}:`, error);
    await postThreadMessage(
      client,
      body,
      "Došlo je do greške u obradi odgovora!\nPokušaj ponovno kasnije. :politecat:"
    );
  }
};

export const setupSlackInteractions = (app: App) => {
  Object.keys(RESPONSE_MAP).forEach((actionId) => {
    app.action(actionId, handleResponse.bind(null, actionId));
  });
};
