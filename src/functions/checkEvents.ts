import { app, InvocationContext, Timer } from "@azure/functions";
import { fetchEvents, processEvent } from "../outlook";

export async function CheckEvents(_myTimer: Timer, context: InvocationContext) {
  try {
    const events = await fetchEvents();

    const results = await Promise.allSettled(
      events.map((event) => processEvent(event, context))
    );

    const failed = results.filter((r) => r.status === "rejected");
    if (failed.length > 0) {
      context.warn(`${failed.length} events failed to process.`);
    }
  } catch (error) {
    context.error("Failed to fetch events:", error);
    throw error;
  }
}

app.timer("CheckEvents", {
  schedule: "0 0 7 * * *",
  handler: CheckEvents,
});
