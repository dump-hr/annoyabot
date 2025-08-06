import { Dayjs } from "dayjs";

export enum ResponseStatus {
  ACCEPTED = "accept",
  TENTATIVE = "tentativelyAccept",
  DECLINED = "decline",
  NONE = "none",
}

export interface Attendee {
  email: string;
  status: ResponseStatus;
}

export type OutlookEvent = {
  id: string;
  iCalUId: string;
  subject: string;
  start: Dayjs;
  end: Dayjs;
  organizer: string;
  attendees: Attendee[];
};

export type RawOutlookEvent = {
  id: string;
  iCalUId: string;
  subject?: string;
  start?: { dateTime?: string };
  end?: { dateTime?: string };
  organizer?: {
    emailAddress?: {
      address?: string;
    };
  };
  attendees?: Array<{
    emailAddress?: {
      address?: string;
    };
    status?: {
      response?: ResponseStatus;
    };
  }>;
};

export type SlackBlock =
  | {
      type: "header" | "section";
      text: {
        type: "plain_text" | "mrkdwn";
        text: string;
        emoji?: boolean;
      };
    }
  | {
      type: "divider";
    }
  | {
      type: "actions";
      elements: Array<{
        type: "button";
        text: {
          type: "plain_text";
          text: string;
          emoji?: boolean;
        };
        action_id: string;
        style?: "primary" | "danger";
        value: string;
      }>;
    }
  | {
      type: "context";
      elements: Array<{
        type: "mrkdwn";
        text: string;
      }>;
    };

export type SlackMessage = {
  text: string;
  blocks: SlackBlock[];
};
