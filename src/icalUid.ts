import { ResponseStatus } from "./types";
import { normalizeGraphResponseStatus } from "./utils";

export function decodeOutlookGlobalId(iCalUId: string): {
  isOutlookId: boolean;
  prefix: string;
  guid?: string;
  timestamp?: Date;
} | null {
  if (!iCalUId || iCalUId.length !== 96 || !/^[0-9A-Fa-f]+$/.test(iCalUId)) {
    return null;
  }

  const outlookPrefix = process.env.OUTLOOK_ID_PREFIX;
  const prefix = iCalUId.toLowerCase().substring(0, 32);

  const isOutlookId = prefix === outlookPrefix;

  const guid = iCalUId.substring(32, 64);

  const last16Bytes = iCalUId.substring(80);
  let timestamp: Date | undefined;

  try {
    const filetime = parseInt(last16Bytes, 16);
    if (!isNaN(filetime)) {
      const microseconds = filetime / 10;
      const epochStart = new Date("1601-01-01T00:00:00Z").getTime();
      timestamp = new Date(epochStart + microseconds);
    }
  } catch (e) {
    console.warn("Could not extract timestamp from iCalUID:", e);
  }

  return {
    isOutlookId,
    prefix,
    guid,
    timestamp,
  };
}

export async function getEventByICalUid(
  userEmail: string,
  iCalUId: string,
  token: string,
): Promise<string | null> {
  const queryId = buildGraphQueryId(iCalUId);
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

export async function getUserResponseStatusByICalUid(
  userEmail: string,
  iCalUId: string,
  token: string,
): Promise<ResponseStatus | null> {
  const queryId = buildGraphQueryId(iCalUId);
  const baseUrl = `${process.env.GRAPH_API_BASE}/users/${userEmail}/calendarView`;

  const params = new URLSearchParams({
    startDateTime: new Date().toISOString(),
    endDateTime: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
    $filter: `iCalUId eq '${queryId}'`,
    $select: "id,responseStatus",
  });

  const response = await fetch(`${baseUrl}?${params.toString()}`, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    console.error(
      `Error fetching response status for ${userEmail}:`,
      await response.text(),
    );
    return null;
  }

  const data = await response.json();
  const event = data.value?.[0];

  return normalizeGraphResponseStatus(event?.responseStatus?.response);
}

function buildGraphQueryId(iCalUId: string): string {
  const decoded = decodeOutlookGlobalId(iCalUId);
  return decoded?.isOutlookId
    ? `${process.env.ICAL_UID_PREFIX}${iCalUId.substring(32)}`
    : iCalUId;
}
