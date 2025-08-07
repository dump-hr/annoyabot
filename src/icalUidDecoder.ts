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
