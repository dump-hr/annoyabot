import { OutlookEvent, ResponseStatus } from "./types";
import { BlobServiceClient, ContainerClient } from "@azure/storage-blob";
import { WebClient } from "@slack/web-api";
import dayjs from "dayjs";
import * as dotenv from "dotenv";

dotenv.config();

export const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

export async function getNonResponders(event: OutlookEvent) {
  const nonResponders = event.attendees.filter(
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

export async function cleanupOldLocks(daysToKeep = 7) {
  for await (const blob of containerClient.listBlobsFlat()) {
    if (!blob.properties.createdOn) continue;

    const ageInDays =
      (Date.now() - blob.properties.createdOn.getTime()) /
      (1000 * 60 * 60 * 24);

    if (ageInDays > daysToKeep) {
      await containerClient.deleteBlob(blob.name).catch(() => {});
      console.log(`Deleted: ${blob.name}`);
    }
  }
}
