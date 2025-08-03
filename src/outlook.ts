import { composeMessage, sendReminder } from "./slackMessaging";
import { getNonResponders } from "./slackHelper";
import { OutlookEvent } from "./types";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";
import timezone from "dayjs/plugin/timezone";

dayjs.extend(utc);
dayjs.extend(timezone);

export async function getAccessToken() {
  const params = new URLSearchParams({
    client_id: process.env.OUTLOOK_CLIENT_ID!,
    client_secret: process.env.OUTLOOK_CLIENT_SECRET!,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    }
  );

  const { access_token } = await response.json();
  if (!access_token) {
    throw new Error("Failed to retrieve access token");
  }

  return access_token;
}

export async function fetchEvents(): Promise<OutlookEvent[]> {
  const tomorrow = dayjs().add(1, "day").format("YYYY-MM-DD");
  const dayAfterTomorrow = dayjs().add(2, "days").format("YYYY-MM-DD");
  const filter = `start/dateTime ge '${tomorrow}' and end/dateTime lt '${dayAfterTomorrow}'`;
  const token = await getAccessToken();

  console.log("Filter:", filter);

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${process.env.OUTLOOK_USER_EMAIL}/calendar/events?$filter=${filter}&$select=id,subject,start,end,organizer,attendees`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  const { value }: { value: OutlookEvent[] } = await response.json();
  if (!Array.isArray(value)) {
    console.error("No events found or invalid response format: ", value);
    return [];
  }

  const events = value.map(
    ({ id, subject, start, end, organizer, attendees }) => ({
      id,
      subject,
      start: dayjs.utc(start).tz("Europe/Zagreb").format("HH:mm"),
      end: dayjs.utc(end).tz("Europe/Zagreb").format("HH:mm"),
      organizer,
      attendees,
    })
  );

  return events;
}

export async function processEvent(event: OutlookEvent) {
  const nonResponders = await getNonResponders(event);
  const message = composeMessage(event);

  for (const user of nonResponders) {
    await sendReminder(user.email, message);
  }
}
