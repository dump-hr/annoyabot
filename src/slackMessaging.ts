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
    console.log(`Sending to ${userEmail}`);
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

  const message: SlackMessage = {
    text: `Hey, dolazis li na: ${event.subject}?`,
    blocks: [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: "Kralju odgovori na event :neutral_face::exclamation:",
          emoji: true,
        },
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `Nisi odgovorio/la na sljedeći event:\n\n*${
            event.subject
          }*\n:date: ${startDate.format(
            "DD.MM.YYYY"
          )} | :clock3: ${startDate.format("HH:mm")} - ${endDate.format(
            "HH:mm"
          )}`,
        },
      },
      {
        type: "divider",
      },
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
            text: "⌛ Rok za odgovor: do sutra!",
          },
        ],
      },
    ],
  };

  return message;
}
