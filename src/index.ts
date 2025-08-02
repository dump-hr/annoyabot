import { fetchEvents, processEvent } from "./outlook";

async function main() {
  const events = await fetchEvents();

  for (const event of events) {
    await processEvent(event);
  }
}

main().catch(console.error);
