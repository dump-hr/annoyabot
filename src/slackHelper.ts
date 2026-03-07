import { OutlookEvent, ResponseStatus } from "./types";
import { BlobServiceClient, ContainerClient } from "@azure/storage-blob";
import { WebClient } from "@slack/web-api";
import dayjs from "dayjs";
import * as dotenv from "dotenv";
import { expandAttendees, getAccessToken } from "./outlook";

dotenv.config();

export const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

export async function getNonResponders(event: OutlookEvent) {
  const token = await getAccessToken();
  const expandedAttendees = await expandAttendees(event.attendees, token);

  console.log("Expanded attendees count:", expandedAttendees.length);

  const nonResponders = expandedAttendees.filter(
    (attendee) =>
      attendee.email.toLowerCase() !== event.organizer.toLowerCase() &&
      attendee.status === ResponseStatus.NONE
  );
  console.log(
    `Event '${event.subject}' has ${nonResponders.length} non-responders`
  );

  return nonResponders;
}

export async function getSlackIdByEmail(email: string): Promise<string> {
  try {
    const response = await slack.users.lookupByEmail({ email });
    if (!response.user?.id) {
      throw new Error(`Slack user not found for email: ${email}`);
    }
    return response.user.id;
  } catch (error) {
    console.error(`Error finding slack ID for ${email}:`, error);
    throw error;
  }
}

export const getSlackUserEmail = async (client: WebClient, userId: string) => {
  const userInfo = await client.users.info({ user: userId });
  return userInfo.user?.profile?.email;
};

export async function getSlackUserName(userId: string): Promise<string> {
  const res = await slack.users.info({ user: userId });

  if (!res.ok || !res.user) {
    throw new Error("Failed to fetch Slack user info");
  }

  const profile = res.user.profile;

  return (
    profile.display_name ||
    profile.real_name ||
    res.user.name ||
    "Nepoznati korisnik"
  );
}

export function getDaysUntilEvent(event: OutlookEvent): number {
  const now = dayjs();
  const startDate = dayjs(event.start);
  const daysUntil = startDate.startOf("day").diff(now.startOf("day"), "days");
  return daysUntil;
}

const blobConnectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
const containerName = "slack-message-locks";

const blobServiceClient =
  BlobServiceClient.fromConnectionString(blobConnectionString);
export const containerClient: ContainerClient =
  blobServiceClient.getContainerClient(containerName);

blobServiceClient.getAccountInfo &&
  blobServiceClient
    .getAccountInfo()
    .then(() => console.log("Connected to Blob service"))
    .catch(() => console.log("Could not connect to Blob service"));

export async function ensureContainer() {
  const exists = await containerClient.exists();
  if (!exists) {
    await containerClient.create();
  }
}

export async function cleanupOldLocks(daysToKeep = 1) {
  const cutoffDate = dayjs().subtract(daysToKeep, "day").startOf("day");

  for await (const blob of containerClient.listBlobsFlat()) {
    const match = blob.name.match(/-(\d{4}-\d{2}-\d{2})-reminder-/);

    if (match) {
      const eventDate = dayjs(match[1]);

      if (eventDate.isBefore(cutoffDate)) {
        await containerClient.deleteBlob(blob.name).catch((err) => {
          console.error(`Failed to delete ${blob.name}:`, err);
        });
        console.log(`Deleted old lock: ${blob.name} (event: ${match[1]})`);
      }
    } else {
      if (blob.properties.createdOn) {
        const ageInDays =
          (Date.now() - blob.properties.createdOn.getTime()) /
          (1000 * 60 * 60 * 24);

        if (ageInDays > 7) {
          await containerClient.deleteBlob(blob.name).catch(() => {});
          console.log(`Deleted old format lock: ${blob.name}`);
        }
      }
    }
  }

  await cleanupStatsEventReferences();
}

async function cleanupStatsEventReferences() {
  try {
    const statsBlob = containerClient.getBlobClient(
      "non-responder-statistics.json"
    );
    const downloadResponse = await statsBlob.download();
    const downloaded = await streamToBuffer(
      downloadResponse.readableStreamBody!
    );
    const stats = JSON.parse(downloaded.toString());

    Object.values(stats).forEach((userStat: any) => {
      if (userStat.eventsMissed && userStat.eventsMissed.length > 50) {
        userStat.eventsMissed = userStat.eventsMissed.slice(-50);
      }
    });

    const blobClient = containerClient.getBlockBlobClient(
      "non-responder-statistics.json"
    );
    const content = JSON.stringify(stats, null, 2);
    await blobClient.upload(content, Buffer.byteLength(content));

    console.log("Cleaned up statistics event references");
  } catch (error: any) {
    if (error?.statusCode !== 404 && error?.code !== "BlobNotFound") {
      console.error("Failed to cleanup stats:", error);
    }
  }
}

export async function streamToBuffer(
  readableStream: NodeJS.ReadableStream
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    readableStream.on("data", (data) => {
      chunks.push(data instanceof Buffer ? data : Buffer.from(data));
    });
    readableStream.on("end", () => {
      resolve(Buffer.concat(chunks));
    });
    readableStream.on("error", reject);
  });
}
