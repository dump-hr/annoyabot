import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import dayjs from "dayjs";
import { fetchEvents } from "../outlook";
import { getNonResponders } from "../outlook";
import { composeMessage } from "../slackMessaging";

export async function PreviewReminders(
  _request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const events = await fetchEvents();
    context.log(
      `Fetched ${events.length} events. Computing who would get reminders (no messages sent)...`,
    );

    const preview: Array<{
      subject: string;
      start: string;
      wouldSendCount: number;
      wouldSendTo: string[];
      skipReason?: string;
    }> = [];

    for (const event of events) {
      const nonResponders = await getNonResponders(event);
      const composed = composeMessage(event);

      const startStr = dayjs(event.start).format("YYYY-MM-DD HH:mm");

      if (!composed) {
        preview.push({
          subject: event.subject,
          start: startStr,
          wouldSendCount: 0,
          wouldSendTo: [],
          skipReason:
            "No reminder scheduled for this event (not 3 or 1 day before)",
        });
        continue;
      }

      const wouldSendTo = nonResponders
        .filter((u) => u.email.toLowerCase() !== event.organizer.toLowerCase())
        .map((u) => u.email);

      preview.push({
        subject: event.subject,
        start: startStr,
        wouldSendCount: wouldSendTo.length,
        wouldSendTo,
      });

      context.log(
        `"${event.subject}" (${startStr}): would send to ${wouldSendTo.length} people:`,
        wouldSendTo,
      );
    }

    return {
      status: 200,
      jsonBody: {
        message: "Dry run — no Slack messages sent.",
        eventsCount: events.length,
        preview,
      },
    };
  } catch (error) {
    context.error("PreviewReminders failed:", error);
    return {
      status: 500,
      jsonBody: { error: String((error as Error).message) },
    };
  }
}

app.http("PreviewReminders", {
  methods: ["GET"],
  authLevel: "function",
  route: "PreviewReminders",
  handler: PreviewReminders,
});
