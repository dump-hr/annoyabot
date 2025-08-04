import { Attendee, OutlookEvent, RawOutlookEvent } from "./types";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";
import timezone from "dayjs/plugin/timezone";

dayjs.extend(utc);
dayjs.extend(timezone);

export const parseOutlookEvents = (responseData: {
  value: RawOutlookEvent[];
}): OutlookEvent[] => {
  return responseData.value.map(
    (event): OutlookEvent => ({
      id: event.id,
      subject: event.subject,
      start: dayjs(event.start?.dateTime),
      end: dayjs(event.end?.dateTime),
      organizer: event.organizer.emailAddress.address,
      attendees:
        event.attendees?.map(
          (attendee): Attendee => ({
            email: attendee?.emailAddress.address,
            status: attendee?.status.response,
          })
        ) || [],
    })
  );
};
