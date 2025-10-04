import { OutlookEvent, SlackMessage } from "./types";
import { getSlackIdByEmail, slack } from "./slackHelper";
import dayjs from "dayjs";

export async function sendReminder(userEmail: string, message: SlackMessage) {
  if (!userEmail) {
    console.error("No email provided");
    return;
  }

  try {
    const userId = await getSlackIdByEmail(userEmail);
    await slack.chat.postMessage({
      channel: userId,
      ...message,
    });
  } catch (error) {
    console.error(`Failed to send to ${userEmail}:`, error);
  }
}

export async function postThreadMessage(client: any, body: any, text: string) {
  try {
    const threadTs =
      body.container?.thread_ts ||
      body.message?.ts ||
      body.container?.message_ts;

    await client.chat.postMessage({
      channel: body.container.channel_id,
      thread_ts: threadTs,
      text,
    });
  } catch (error) {
    console.error("Error sending message:", error);
  }
}

export function composeMessage(event: OutlookEvent): SlackMessage {
  const startDate = dayjs(event.start);
  const endDate = dayjs(event.end);
  const daysUntil = Math.ceil(dayjs(event.start).diff(dayjs(), "day", true));

  let notificationText: string;
  let headerText: string;
  let bodyIntro: string;
  let contextText: string;

  if (daysUntil === 3) {
    notificationText = `Event za 3 dana, a od tebe ni glasa :pensive:`;
    headerText = `Još tri dana i tri noći!`;
    bodyIntro = "Nisi odgovorio/la na sljedeći event :upside_down_face:";
    contextText = ":hourglass: Rok za odgovor: još 2 (i po) dana!";
  } else if (daysUntil === 1) {
    notificationText = `Dolaziš li na event? :thinking_face:`;
    headerText = `Do sutra imaš vremena… ili nemaš...`;
    bodyIntro = "Ako nisi siguran/na, stisni na možda :face_with_rolling_eyes:";
    contextText = ":hourglass: Rok za odgovor: do sutra!";
  } else {
    return null;
  }

  const message: SlackMessage = {
    text: notificationText,
    blocks: [
      {
        type: "header",
        text: { type: "plain_text", text: headerText, emoji: true },
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `${bodyIntro}\n\n*${event.subject}*\n:date: ${startDate.format(
            "DD.MM.YYYY"
          )} | :clock3: ${startDate.format("HH:mm")} - ${endDate.format(
            "HH:mm"
          )}`,
        },
      },
      { type: "divider" },
      {
        type: "actions",
        elements: [
          {
            type: "button",
            text: {
              type: "plain_text",
              text: ":blob-yes: Dolazim",
              emoji: true,
            },
            style: "primary",
            action_id: "event_accept",
            value: event.iCalUId,
          },
          {
            type: "button",
            text: {
              type: "plain_text",
              text: "🤷 Možda",
              emoji: true,
            },
            action_id: "event_tentative",
            value: event.iCalUId,
          },
          {
            type: "button",
            text: {
              type: "plain_text",
              text: ":blob-no: Ne mogu",
              emoji: true,
            },
            style: "danger",
            action_id: "event_decline",
            value: event.iCalUId,
          },
        ],
      },
      {
        type: "context",
        elements: [
          {
            type: "mrkdwn",
            text: contextText,
          },
        ],
      },
    ],
  };

  return message;
}
