import { composeMessage, sendReminder } from "./slackMessaging";
import { parseOutlookEvents } from "./utils";
import { getNonResponders } from "./slackHelper";
import { OutlookEvent } from "./types";
import timezone from "dayjs/plugin/timezone";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";

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
  const filter = `start/dateTime ge '${tomorrow}T00:00:00' and end/dateTime lt '${dayAfterTomorrow}T00:00:00'`;
  const token = await getAccessToken();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${process.env.OUTLOOK_USER_EMAIL}/calendar/events?filter=${filter}&select=id,subject,start,end,organizer,attendees`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  const responseData = await response.json();

  if (!response.ok || !Array.isArray(responseData.value)) {
    console.error("Error fetching events:", responseData);
    return [];
  }

  const events = parseOutlookEvents(responseData);

  const testEvents = events.filter((event) =>
    event.subject.includes("annoyabot-test")
  );

  return testEvents;
}

export async function processEvent(event: OutlookEvent) {
  const nonResponders = await getNonResponders(event);
  const message = composeMessage(event);

  for (const user of nonResponders) {
    await sendReminder(user.email, message);
  }
}
