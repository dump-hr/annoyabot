import { WebClient } from "@slack/web-api";
import { OutlookEvent, ResponseStatus } from "./types";
import * as dotenv from "dotenv";

dotenv.config();

export const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

export async function getNonResponders(event: OutlookEvent) {
  return event.attendees.filter(
    (attendee) =>
      attendee.status === ResponseStatus.NONE &&
      attendee.email !== event.organizer
  );
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
