import { app, InvocationContext, Timer } from "@azure/functions";
import { fetchEvents, processEvent } from "../outlook";

export async function CheckEvents(
  _myTimer: Timer,
  context: InvocationContext
): Promise<void> {
  context.log("Timer function running...fetching events...");

  try {
    const events = await fetchEvents();
    context.log(`Found ${events.length} events to process`);

    for (const event of events) {
      await processEventWithRetry(event, context);
    }

    context.log("Event processing completed");
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
      await processEvent(event);
      return;
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      context.warn(`Retry ${i + 1}/${maxRetries} for event ${event.id}`);
      await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
    }
  }
}

// app.timer("CheckEvents", {
//   schedule: "0 0 9 * * *",
//   handler: CheckEvents,
// });

app.timer("CheckEvents", {
  schedule: "0 0 9 * * *",
  runOnStartup: true,
  handler: CheckEvents,
});
