import { fetchEvents, processEvent } from "./outlook";
import { startServer } from "./server";

startServer();

async function sendDailyReminders() {
  const events = await fetchEvents();
  for (const event of events) {
    await processEvent(event);
  }
}

sendDailyReminders().catch(console.error);
