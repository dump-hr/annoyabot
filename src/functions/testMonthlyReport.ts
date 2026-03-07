import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import {
  saveStats,
  sendMonthlyReportToSlack,
  archiveAndResetMonthlyStats,
  generateMonthlySlackReport,
} from "../statistics";
import { StatsRecord } from "../types";
import { getSlackIdByEmail, slack } from "../slackHelper";

export async function TestMonthlyReport(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const mockStats: StatsRecord = {
      "marko.horvat@company.com": {
        email: "marko.horvat@company.com",
        name: "Marko Horvat",
        remindersSent3Days: 12,
        remindersSent1Day: 8,
        totalReminders: 20,
        lastReminderDate: "2024-12-28 14:23",
        eventsMissed: ["Team Meeting", "Sprint Planning", "Client Demo"],
      },
      "ana.kovac@company.com": {
        email: "ana.kovac@company.com",
        name: "Ana Kovač",
        remindersSent3Days: 9,
        remindersSent1Day: 7,
        totalReminders: 16,
        lastReminderDate: "2024-12-27 09:15",
        eventsMissed: ["Design Review", "Product Demo"],
      },
      "ivan.novak@company.com": {
        email: "ivan.novak@company.com",
        name: "Ivan Novak",
        remindersSent3Days: 8,
        remindersSent1Day: 5,
        totalReminders: 13,
        lastReminderDate: "2024-12-26 16:42",
        eventsMissed: ["Strategy Meeting", "All Hands"],
      },
      "petra.babic@company.com": {
        email: "petra.babic@company.com",
        name: "Petra Babić",
        remindersSent3Days: 6,
        remindersSent1Day: 4,
        totalReminders: 10,
        lastReminderDate: "2024-12-25 11:30",
        eventsMissed: ["Quarterly Review"],
      },
      "luka.jovic@company.com": {
        email: "luka.jovic@company.com",
        name: "Luka Jović",
        remindersSent3Days: 5,
        remindersSent1Day: 3,
        totalReminders: 8,
        lastReminderDate: "2024-12-24 13:12",
        eventsMissed: ["Workshop", "Training"],
      },
      "maja.pavlic@company.com": {
        email: "maja.pavlic@company.com",
        name: "Maja Pavlić",
        remindersSent3Days: 4,
        remindersSent1Day: 2,
        totalReminders: 6,
        lastReminderDate: "2024-12-23 10:05",
        eventsMissed: ["Standup"],
      },
      "tomislav.kralj@company.com": {
        email: "tomislav.kralj@company.com",
        name: "Tomislav Kralj",
        remindersSent3Days: 3,
        remindersSent1Day: 1,
        totalReminders: 4,
        lastReminderDate: "2024-12-22 08:47",
        eventsMissed: ["Retrospective"],
      },
    };


    const message = await generateMonthlySlackReport();

    const yourEmail = request.query.get("email") || process.env.YOUR_EMAIL;

    if (!yourEmail) {
      return {
        status: 400,
        body: "Provide your email: ?email=tvoj.email@company.com or set YOUR_EMAIL env var",
      };
    }

    const yourUserId = await getSlackIdByEmail(yourEmail);

    await slack.chat.postMessage({
      channel: yourUserId,
      ...message,
    });

    return {
      status: 200,
      body: "Test monthly report sent successfully! Check your Slack channel.",
    };
  } catch (error) {
    context.error("Failed to send test report:", error);
    return {
      status: 500,
      body: `Error: ${error.message}`,
    };
  }
}

app.http("TestMonthlyReport", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: TestMonthlyReport,
});
