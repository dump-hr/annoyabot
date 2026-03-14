import { InvocationContext } from "@azure/functions";
import dayjs from "dayjs";
import timezone from "dayjs/plugin/timezone";
import utc from "dayjs/plugin/utc";
import { getEventByICalUid, getUserResponseStatusByICalUid } from "./icalUid";
import { getAccessToken } from "./outlookAuth";
import { composeMessage, sendReminder } from "./slackMessaging";
import { Attendee, OutlookEvent, ResponseStatus } from "./types";
import { parseOutlookEvents } from "./utils";

dayjs.extend(utc);
dayjs.extend(timezone);

export async function fetchEvents(): Promise<OutlookEvent[]> {
  const start = dayjs().startOf("day").format("YYYY-MM-DD");
  const end = dayjs().add(4, "day").endOf("day").format("YYYY-MM-DD");
  const filter = `start/dateTime ge '${start}T00:00:00' and end/dateTime le '${end}T23:59:59'`;
  const token = await getAccessToken();

  const params = new URLSearchParams({
    $filter: filter,
    $select: "id,subject,start,end,organizer,attendees,iCalUId",
  });

  const response = await fetch(
    `${process.env.GRAPH_API_BASE}/users/${process.env.OUTLOOK_USER_EMAIL}/calendar/events?${params.toString()}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  const responseData = await response.json();

  if (!response.ok || !Array.isArray(responseData.value)) {
    console.error("Error fetching events:", responseData);
    return [];
  }

  const events = parseOutlookEvents(responseData);
  console.log(`Fetched ${events.length} events from Outlook.`);

  return events;
}

export async function processEvent(
  event: OutlookEvent,
  context?: InvocationContext,
): Promise<void> {
  const nonResponders = await getNonResponders(event);
  if (nonResponders.length === 0) return;

  console.log(
    `Event "${event.subject}" - Non-responders:`,
    nonResponders.map((u) => u.email),
  );

  const composed = composeMessage(event);
  if (!composed) return;

  const { message, reminderType } = composed;

  for (const user of nonResponders) {
    if (user.email.toLowerCase() === event.organizer.toLowerCase()) continue;

    try {
      await sendReminder(user.email, message, event, reminderType);
    } catch (err) {
      context?.error(`Failed to send to ${user.email}:`, err);
    }
  }
}

export async function updateEventResponse(
  iCalUId: string,
  userEmail: string,
  statusResponse: ResponseStatus,
): Promise<{ success: boolean; errorType?: string }> {
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
      throw new Error(`Failed to update status: ${error.error?.message}`);
    }

    return { success: true };
  } catch (error) {
    console.error("Update failed:", (error as Error).message);
    return { success: false, errorType: "UNKNOWN_ERROR" };
  }
}

export async function expandAttendees(
  attendees: Attendee[],
  token: string,
): Promise<Attendee[]> {
  const expanded: Attendee[] = [];
  const groupCache = new Map<string, Attendee[]>();
  const seen = new Set<string>();

  const statusByEmail = new Map<string, ResponseStatus>();
  for (const a of attendees) {
    if (a.email) {
      statusByEmail.set(a.email.toLowerCase(), a.status);
    }
  }

  for (const attendee of attendees) {
    const email = attendee.email?.toLowerCase();
    if (!email) continue;

    if (!groupCache.has(email)) {
      const members = await getDistributionListMembers(email, token);
      groupCache.set(email, members);
    }

    const members = groupCache.get(email) ?? [];

    if (members.length > 0) {
      for (const m of members) {
        const mEmail = m.email?.toLowerCase() ?? "";
        if (!mEmail || seen.has(mEmail)) continue;
        seen.add(mEmail);
        expanded.push({
          ...m,
          type: "user",
          status: statusByEmail.get(mEmail) ?? ResponseStatus.NONE,
        });
      }
    } else {
      if (seen.has(email)) continue;
      seen.add(email);
      expanded.push({
        ...attendee,
        type: "user",
        status: attendee.status ?? ResponseStatus.NONE,
      });
    }
  }

  return expanded;
}

export async function getDistributionListMembers(
  groupEmail: string,
  token: string,
): Promise<Attendee[]> {
  const members: Attendee[] = [];

  const groupRes = await fetch(
    `https://graph.microsoft.com/v1.0/groups?$filter=mail eq '${groupEmail}'`,
    {
      headers: { Authorization: `Bearer ${token}` },
    },
  );

  if (!groupRes.ok) {
    console.error(
      `Failed to resolve group '${groupEmail}':`,
      groupRes.status,
      groupRes.statusText,
      await groupRes.text(),
    );
    return members;
  }

  const groupData = await groupRes.json();
  const group = groupData.value?.[0];

  console.log("Resolving group:", groupEmail);
  console.log("Group found:", group);

  if (!group) {
    console.warn(`Group not found for email: ${groupEmail}`);
    return members;
  }

  let url = `https://graph.microsoft.com/v1.0/groups/${group.id}/transitiveMembers?$select=mail,userPrincipalName`;

  while (url) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!res.ok) {
      console.error(
        `Failed to list members for group '${groupEmail}' (${group.id}):`,
        res.status,
        res.statusText,
        await res.text(),
      );
      break;
    }

    const data = await res.json();

    if (!data.value) break;

    for (const m of data.value) {
      const email = m.mail || m.userPrincipalName;
      if (email) {
        members.push({
          email,
          type: "user",
          status: ResponseStatus.NONE,
        });
      }
    }

    url = data["@odata.nextLink"];
  }

  return members;
}

export async function getNonResponders(event: OutlookEvent) {
  const token = await getAccessToken();
  const expandedAttendees = await expandAttendees(event.attendees, token);

  console.log("Expanded attendees count:", expandedAttendees.length);

  const maybeNonResponders = expandedAttendees.filter(
    (attendee) =>
      attendee.email.toLowerCase() !== event.organizer.toLowerCase() &&
      attendee.status === ResponseStatus.NONE,
  );

  const verified = await Promise.all(
    maybeNonResponders.map(async (attendee) => {
      const latestStatus = await getUserResponseStatusByICalUid(
        attendee.email,
        event.iCalUId,
        token,
      );

      if (latestStatus === null) {
        console.warn(
          `Skipping ${attendee.email}: unable to verify current response status.`,
        );
        return null;
      }

      if (latestStatus !== ResponseStatus.NONE) {
        console.log(
          `Skipping ${attendee.email}: latest status is ${latestStatus}.`,
        );
        return null;
      }

      return attendee;
    }),
  );

  const nonResponders = verified.filter(
    (attendee): attendee is NonNullable<typeof attendee> => attendee !== null,
  );

  console.log(
    `Event '${event.subject}' has ${nonResponders.length} non-responders`,
  );

  return nonResponders;
}
