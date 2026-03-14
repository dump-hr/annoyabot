import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import { expandAttendees } from "../outlook";
import { getAccessToken } from "../outlookAuth";
import { Attendee, ResponseStatus } from "../types";

export async function testDl(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  const token = await getAccessToken();

  const attendees: Attendee[] = [
    {
      email: "sara.ticinovic@dump.hr",
      type: "user",
      status: ResponseStatus.NONE,
    },
    {
      email: "clanovi@dump.hr",
      type: "group",
      status: ResponseStatus.NONE,
    },
  ];

  const expanded = await expandAttendees(attendees, token);

  context.log("Expanded attendees:");
  expanded.forEach((a) => context.log(a.email));

  return {
    status: 200,
    jsonBody: expanded.map((a) => a.email),
  };
}

app.http("TestDl", {
  methods: ["GET"],
  authLevel: "function",
  route: "testDl",
  handler: testDl,
});
