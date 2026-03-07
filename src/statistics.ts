import {
  containerClient,
  streamToBuffer,
  slack,
  getSlackIdByEmail,
  getSlackUserName,
} from "./slackHelper";
import dayjs from "dayjs";
import { StatsRecord, NonResponderStats } from "./types";

const STATS_BLOB_NAME = "non-responder-statistics.json";
const ARCHIVE_BLOB_PREFIX = "statistics-archive/";

export async function loadStats(): Promise<StatsRecord> {
  try {
    const blobClient = containerClient.getBlobClient(STATS_BLOB_NAME);
    const downloadResponse = await blobClient.download();
    const downloaded = await streamToBuffer(
      downloadResponse.readableStreamBody!
    );
    const data = JSON.parse(downloaded.toString());
    return data;
  } catch (error: any) {
    if (error?.statusCode === 404 || error?.code === "BlobNotFound") {
      console.log("Stats file not found, creating new one");
      return {};
    }
    throw error;
  }
}

export async function saveStats(stats: StatsRecord): Promise<void> {
  const blobClient = containerClient.getBlockBlobClient(STATS_BLOB_NAME);
  const content = JSON.stringify(stats, null, 2);
  await blobClient.upload(content, Buffer.byteLength(content));
}

export async function recordReminderSent(
  userEmail: string,
  eventSubject: string,
  reminderType: "3days" | "1day",
  userName?: string
): Promise<void> {
  const stats = await loadStats();

  let displayName: string | undefined;

  try {
    const slackId = await getSlackIdByEmail(userEmail);
    displayName = (await getSlackUserName(slackId)) ?? undefined;
  } catch (e) {
    console.warn("Failed to fetch Slack display name for", userEmail);
  }

  if (!stats[userEmail]) {
    stats[userEmail] = {
      email: userEmail,
      name: userName,
      displayName: displayName,
      remindersSent3Days: 0,
      remindersSent1Day: 0,
      totalReminders: 0,
      eventsMissed: [],
    };
  }

  const userStats = stats[userEmail];

  if (reminderType === "3days") {
    userStats.remindersSent3Days++;
  } else {
    userStats.remindersSent1Day++;
  }
  userStats.totalReminders++;
  userStats.lastReminderDate = dayjs().format("YYYY-MM-DD HH:mm");

  if (!userStats.eventsMissed.includes(eventSubject)) {
    userStats.eventsMissed.push(eventSubject);
  }

  if (displayName) {
    userStats.displayName = displayName;
  }

  if (userName) {
    userStats.name = userName;
  }

  await saveStats(stats);
  console.log(
    `Recorded reminder for ${userEmail}: ${reminderType} (Total: ${userStats.totalReminders})`
  );
}

export async function getTopNonResponders(
  limit: number = 10
): Promise<NonResponderStats[]> {
  const stats = await loadStats();
  const sorted = Object.values(stats).sort(
    (a, b) => b.totalReminders - a.totalReminders
  );
  return sorted.slice(0, limit);
}

export async function getUserStats(
  email: string
): Promise<NonResponderStats | null> {
  const stats = await loadStats();
  return stats[email] || null;
}

export async function archiveAndResetMonthlyStats(): Promise<void> {
  const stats = await loadStats();

  if (Object.keys(stats).length === 0) {
    console.log("No statistics to archive");
    return;
  }

  const archiveDate = dayjs().subtract(1, "month").format("YYYY-MM");
  const archiveName = `${ARCHIVE_BLOB_PREFIX}${archiveDate}.json`;

  const archiveBlob = containerClient.getBlockBlobClient(archiveName);
  const archiveContent = JSON.stringify(stats, null, 2);
  await archiveBlob.upload(archiveContent, Buffer.byteLength(archiveContent));

  console.log(`Archived statistics to ${archiveName}`);

  await saveStats({});
  console.log("Statistics reset for new month");
}

export async function generateMonthlySlackReport(): Promise<any> {
  const stats = await loadStats();
  const sorted = Object.values(stats).sort(
    (a, b) => b.totalReminders - a.totalReminders
  );

  const previousMonth = dayjs().subtract(1, "month").format("MMMM YYYY");
  const totalReminders = sorted.reduce((sum, s) => sum + s.totalReminders, 0);

  const top10 = sorted.slice(0, 10);

  const blocks: any[] = [
    {
      type: "header",
      text: {
        type: "plain_text",
        text: `📊 Mjesečni izvještaj - ${previousMonth}`,
        emoji: true,
      },
    },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*Statistika podsjetnika za evente*\n\n• Ukupno praćenih korisnika: *${sorted.length}*\n• Ukupno poslanih podsjetnika: *${totalReminders}*`,
      },
    },
    { type: "divider" },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*🏆 Top 10 ne-odgovarača:*",
      },
    },
  ];

  top10.forEach((stat, index) => {
    const medal =
      index === 0
        ? "🥇"
        : index === 1
        ? "🥈"
        : index === 2
        ? "🥉"
        : `${index + 1}.`;
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: `${medal} *${stat.displayName ?? stat.name}*\n   3-dnevni: ${
          stat.remindersSent3Days
        } | 1-dnevni: ${stat.remindersSent1Day} | Ukupno: *${
          stat.totalReminders
        }*`,
      },
    });
  });

  blocks.push({ type: "divider" });
  blocks.push({
    type: "context",
    elements: [
      {
        type: "mrkdwn",
        text: `Izvještaj generiran: ${dayjs().format("DD.MM.YYYY HH:mm")}`,
      },
    ],
  });

  return { blocks, text: `Mjesečni izvještaj - ${previousMonth}` };
}

export async function sendMonthlyReportToSlack(
  channelId: string
): Promise<void> {
  try {
    const message = await generateMonthlySlackReport();

    await slack.chat.postMessage({
      channel: channelId,
      ...message,
    });

    console.log(`Monthly report sent to channel ${channelId}`);
  } catch (error) {
    console.error("Failed to send monthly report:", error);
    throw error;
  }
}
