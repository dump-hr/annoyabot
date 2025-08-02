interface Attendee {
  email: string;
  status: "none" | "accepted" | "tentative" | "declined";
}

export type OutlookEvent = {
  subject: string;
  start: string;
  end: string;
  organizer: string;
  attendees: Attendee[];
};
type SlackTextObject = {
  type: "plain_text" | "mrkdwn";
  text: string;
  emoji?: boolean;
};

type SlackBlock =
  | {
      type: "header";
      text: SlackTextObject;
    }
  | {
      type: "section";
      text: SlackTextObject;
    }
  | {
      type: "divider";
    }
  | {
      type: "actions";
      elements: SlackButton[];
    }
  | {
      type: "context";
      elements: Array<{
        type: "mrkdwn";
        text: string;
      }>;
    };

type SlackButton = {
  type: "button";
  text: {
    type: "plain_text";
    text: string;
    emoji?: boolean;
  };
  action_id: string;
  style?: "primary" | "danger";
};

export type SlackMessage = {
  text: string;
  blocks: SlackBlock[];
};
