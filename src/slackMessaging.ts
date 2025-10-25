import {
  ComposedReminder,
  OutlookEvent,
  ReminderType,
  SlackMessage,
} from "./types";
import { containerClient, getSlackIdByEmail, slack } from "./slackHelper";
import dayjs from "dayjs";

export async function sendReminder(
  userEmail: string,
  message: SlackMessage,
  event: OutlookEvent,
  reminderType: "3days" | "1day"
) {
  if (!userEmail) return;

  const blobName = `${userEmail}-${event.iCalUId}-reminder-${reminderType}.lock`;

  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  try {
    await blockBlobClient.upload("", 0, { conditions: { ifNoneMatch: "*" } });
  } catch (err: any) {
    if (err.code === "BlobAlreadyExists") {
      console.log(
        `Message for ${userEmail}, event: ${event.subject} was already sent.`
      );
      return;
    }
    throw err;
  }

  const userId = await getSlackIdByEmail(userEmail);
  const result = await slack.chat.postMessage({
    channel: userId,
    ...message,
  });

  console.log(
    `Message sent successfully to ${userEmail} for event ${event.subject}. TS: ${result.ts}`
  );
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

export function composeMessage(event: OutlookEvent): ComposedReminder | null {
  const now = dayjs();
  const startDate = dayjs(event.start);
  const endDate = dayjs(event.end);
  const daysUntil = startDate.startOf("day").diff(now.startOf("day"), "days");
  let reminderType: ReminderType;

  let notificationText: string;
  let headerText: string;
  let bodyIntro: string;
  let contextText: string;

  if (daysUntil === 3) {
    reminderType = "3days";
    notificationText = `Event za 3 dana, a od tebe ni glasa :pensive:`;
    headerText = `Još tri dana i tri noći!`;
    bodyIntro = "Nisi odgovorio/la na sljedeći event :upside_down_face:";
    contextText = ":hourglass: Rok za odgovor: još 2 (i po) dana!";
  } else if (daysUntil === 1) {
    reminderType = "1day";
    notificationText = `Dolaziš li na event? :thinking_face:`;
    headerText = `Do sutra imaš vremena… ili nemaš...`;
    bodyIntro = "Ako nisi siguran/na, stisni na možda :face_with_rolling_eyes:";
    contextText = ":hourglass: Rok za odgovor: do sutra!";
  } else {
    console.log(
      `ComposeMessage: daysUntil ${daysUntil}, skipping message, for event ${event.subject}`
    );
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

  return { message, reminderType };
}
