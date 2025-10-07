import { app, InvocationContext, Timer } from "@azure/functions";
import { fetchEvents, processEvent } from "../outlook";

export async function CheckEvents(_myTimer: Timer, context: InvocationContext) {
  try {
    const events = await fetchEvents();

    const results = await Promise.allSettled(
      events.map((event) => processEventWithRetry(event, context))
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

async function processEventWithRetry(
  event: any,
  context: InvocationContext,
  maxRetries = 3
) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      await processEvent(event, context);
      return;
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      context.warn(`Retry ${i + 1}/${maxRetries} for event ${event.id}`);
      await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
    }
  }
}

app.timer("CheckEvents", {
  schedule: "0 0 7 * * *",
  runOnStartup: true,
  handler: CheckEvents,
});
