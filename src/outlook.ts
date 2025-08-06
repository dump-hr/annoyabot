import { composeMessage, sendReminder } from "./slackMessaging";
import { parseOutlookEvents } from "./utils";
import { getNonResponders } from "./slackHelper";
import { OutlookEvent, ResponseStatus } from "./types";
import timezone from "dayjs/plugin/timezone";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";
import { decodeOutlookGlobalId } from "./decodeID";

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
  const tomorrow = dayjs().add(18, "day").format("YYYY-MM-DD");
  const dayAfterTomorrow = dayjs().add(19, "days").format("YYYY-MM-DD");
  const filter = `start/dateTime ge '${tomorrow}T00:00:00' and end/dateTime lt '${dayAfterTomorrow}T00:00:00'`;
  const token = await getAccessToken();

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${process.env.OUTLOOK_USER_EMAIL}/calendar/events?filter=${filter}&select=id,subject,start,end,organizer,attendees,iCalUId`,
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
    event.subject.includes("Annoyabot-test")
  );

  return testEvents;
}

export async function processEvent(event: OutlookEvent) {
  const nonResponders = await getNonResponders(event);
  const message = composeMessage(event);

  await sendReminder(process.env.MY_EMAIL, message);

  // for (const user of nonResponders) {
  //   await sendReminder(user.email, message);
  // }
}

export async function updateEventResponse(
  iCalUId: string,
  userEmail: string,
  statusResponse: ResponseStatus
) {
  const token = await getAccessToken();
  const eventId = await getEventIdByICalUid(userEmail, iCalUId, token);
  if (!eventId) {
    throw new Error("Event not found for the given iCalUID");
  }

  const eventUrl = `https://graph.microsoft.com/v1.0/users/${userEmail}/events/${eventId}/${statusResponse}`;

  try {
    const response = await fetch(eventUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(
        `Failed to update status: ${error.error?.message || "Unknown error"}`
      );
    }

    console.log("Status updated successfully:");

    return true;
  } catch (error) {
    console.error("Update failed:", {
      eventId: eventId.substring(0, 8) + "...",
      userEmail,
      error: error.message,
      stack: error.stack,
    });
    throw error;
  }
}

export async function getEventIdByICalUid(
  userEmail: string,
  iCalUId: string,
  token: string
): Promise<string | null> {
  const decoded = decodeOutlookGlobalId(iCalUId);
  const queryId = decoded?.isOutlookId
    ? `040000008200E00074C5B7101A82E008${iCalUId.substring(32)}`
    : iCalUId;

  const url = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${encodeURIComponent(
    new Date().toISOString()
  )}&endDateTime=${encodeURIComponent(
    new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString()
  )}&$filter=iCalUId eq '${queryId}'&$select=id`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    console.error("Error fetching event by iCalUId:", await response.text());
    return null;
  }

  const data = await response.json();
  return data.value?.[0]?.id || null;
}
