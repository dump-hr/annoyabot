import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import { getSlackApp } from "../slack";
import { handleResponse } from "../slackInteractions";

export async function SlackResponse(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const body = await request.text();

    let payload;
    if (body.startsWith("payload=")) {
      const params = new URLSearchParams(body);
      payload = JSON.parse(params.get("payload") || "{}");
    } else {
      payload = JSON.parse(body);
    }

    if (payload.type === "url_verification") {
      return {
        status: 200,
        jsonBody: { challenge: payload.challenge },
      };
    }

    if (payload.type === "block_actions") {
      const actionId = payload.actions[0].action_id;
      const slackApp = getSlackApp();

      await handleResponse(actionId, {
        ack: async () => {},
        body: payload,
        client: slackApp.client,
      });

      return { status: 200, body: "" };
    }

    context.warn("Unknown payload type:", payload.type);
    return { status: 400, body: "Unknown payload type" };
  } catch (error) {
    context.error("SlackHandler error:", error);
    return {
      status: 500,
      body: "Internal Server Error",
    };
  }
}

app.http("SlackResponse", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "slack/events",
  handler: SlackResponse,
});
