import { composeMessage, sendReminder } from "./slackMessaging";
import { getNonResponders } from "./slackHelper";
import { OutlookEvent } from "./types";

export async function fetchEvents(): Promise<OutlookEvent[]> {
  const today = new Date().toISOString().split("T")[0]; // Daje današnji datum u formatu "YYYY-MM-DD"

  return [
    {
      subject: "[MOCK] Sastanak o projektu",
      start: `${today}T16:00:00`,
      end: `${today}T17:00:00`,
      organizer: process.env.OUTLOOK_USER_ID,
      attendees: [{ email: process.env.OUTLOOK_USER_ID, status: "none" }],
    },
  ];
}

export async function processEvent(event: OutlookEvent) {
  const nonResponders = await getNonResponders(event);
  const message = composeMessage(event);

  for (const user of nonResponders) {
    await sendReminder(user.email, message);
  }
}
