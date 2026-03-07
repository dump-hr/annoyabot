import { app, InvocationContext, Timer } from "@azure/functions";
import {
  archiveAndResetMonthlyStats,
  sendMonthlyReportToSlack,
} from "../statistics";

export async function MonthlyReport(
  _myTimer: Timer,
  context: InvocationContext
) {
  try {
    const reportChannelId = process.env.SLACK_REPORT_CHANNEL_ID;
    if (!reportChannelId) {
      context.error("SLACK_REPORT_CHANNEL_ID is not set");
      return;
    }

    await sendMonthlyReportToSlack(reportChannelId);
    context.log("Monthly report sent to Slack");

    await archiveAndResetMonthlyStats();
    context.log("Statistics archived and reset");

    context.log("Monthly report completed successfully");
  } catch (error) {
    context.error("Failed to generate monthly report:", error);
    throw error;
  }
}

app.timer("MonthlyReport", {
  schedule: "0 0 9 1 * *",
  handler: MonthlyReport,
});
