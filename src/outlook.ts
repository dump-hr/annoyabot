import { composeMessage, sendReminder } from "./slackMessaging";
import { OutlookEvent, ResponseStatus } from "./types";
import { decodeOutlookGlobalId } from "./icalUidDecoder";
import { parseOutlookEvents } from "./utils";
import { getNonResponders } from "./slackHelper";
import timezone from "dayjs/plugin/timezone";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";

dayjs.extend(utc);
dayjs.extend(timezone);

export async function getAccessToken() {
  const params = new URLSearchParams({
    client_id: process.env.OUTLOOK_CLIENT_ID,
    client_secret: process.env.OUTLOOK_CLIENT_SECRET,
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
  const start = dayjs().startOf("day").format("YYYY-MM-DD");
  const end = dayjs().add(11, "day").endOf("day").format("YYYY-MM-DD");
  const filter = `start/dateTime ge '${start}T00:00:00' and end/dateTime le '${end}T23:59:59'`;
  const token = await getAccessToken();

  console.log("filter", filter);

  const response = await fetch(
    `${process.env.GRAPH_API_BASE}/users/${process.env.OUTLOOK_USER_EMAIL}/calendar/events?filter=${filter}&select=id,subject,start,end,organizer,attendees,iCalUId`,
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

  return events;
}

export async function processEvent(event: OutlookEvent) {
  const nonResponders = await getNonResponders(event);
  const message = composeMessage(event);

  if (!message) {
    console.log(`Skipping ${event.subject}, not 3 or 1 day before.`);
    return;
  }

  try {
    await sendReminder(process.env.MY_EMAIL, message);
    console.log(`Sent reminder for ${event.subject} me`);
  } catch (err) {
    console.error(`Failed to send reminder for ${event.subject}`, err);
  }

  // for (const user of nonResponders) {

  // if (user.email === event.organizer) {
  //   console.log(`Skipping organizer: ${user.email}`);
  //   continue;
  // }

  // try {
  //   await sendReminder(user.email, message);
  //   console.log(`Sent to ${user.email}`);
  // } catch (err) {
  //   console.error(`Failed to send to ${user.email}`, err);
  // }
}

export async function updateEventResponse(
  iCalUId: string,
  userEmail: string,
  statusResponse: ResponseStatus
) {
  const token = await getAccessToken();
  const eventId = await getEventByICalUid(userEmail, iCalUId, token);
  if (!eventId) {
    return { success: false, errorType: "EVENT_NOT_FOUND" };
  }

  const eventUrl = `${process.env.GRAPH_API_BASE}/users/${userEmail}/events/${eventId}`;

  try {
    const response = await fetch(`${eventUrl}/${statusResponse}`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        sendResponse: false,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(
        `Failed to update status: ${error.error?.message || "Unknown error"}`
      );
    }

    return { success: true };
  } catch (error) {
    console.error("Update failed:", error.message);
    return { success: false, errorType: "UNKNOWN_ERROR" };
  }
}

export async function getEventByICalUid(
  userEmail: string,
  iCalUId: string,
  token: string
): Promise<string | null> {
  const decoded = decodeOutlookGlobalId(iCalUId);
  const queryId = decoded?.isOutlookId
    ? `${process.env.ICAL_UID_PREFIX}${iCalUId.substring(32)}`
    : iCalUId;

  const baseUrl = `${process.env.GRAPH_API_BASE}/users/${userEmail}/calendarView`;

  const params = new URLSearchParams({
    startDateTime: new Date().toISOString(),
    endDateTime: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
    $filter: `iCalUId eq '${queryId}'`,
    $select: "id",
  });

  const url = `${baseUrl}?${params.toString()}`;

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
  const event = data.value?.[0];

  if (!event) {
    return null;
  }

  return event.id;
}
