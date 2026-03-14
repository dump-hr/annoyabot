import dayjs from "dayjs";
import utc from "dayjs/plugin/utc";
import timezone from "dayjs/plugin/timezone";
import {
  Attendee,
  OutlookEvent,
  RawOutlookEvent,
  ResponseStatus,
} from "./types";

dayjs.extend(utc);
dayjs.extend(timezone);

export const parseOutlookEvents = (responseData: {
  value: RawOutlookEvent[];
}): OutlookEvent[] => {
  return responseData.value.map(
    (event): OutlookEvent => ({
      id: event.id,
      iCalUId: event.iCalUId,
      subject: event.subject,
      start: convertToLocalTime(event.start?.dateTime),
      end: convertToLocalTime(event.end?.dateTime),
      organizer: event.organizer.emailAddress.address,
      attendees:
        event.attendees?.map(
          (attendee): Attendee => ({
            email: attendee?.emailAddress.address,
            status: normalizeGraphResponseStatus(attendee?.status?.response),
          }),
        ) || [],
    }),
  );
};

export function normalizeGraphResponseStatus(
  response?: string,
): ResponseStatus {
  switch (response?.toLowerCase()) {
    case "accept":
    case "accepted":
      return ResponseStatus.ACCEPTED;
    case "tentativelyaccept":
    case "tentativelyaccepted":
      return ResponseStatus.TENTATIVE;
    case "decline":
    case "declined":
      return ResponseStatus.DECLINED;
    case "none":
    case "notresponded":
    default:
      return ResponseStatus.NONE;
  }
}

const convertToLocalTime = (dateTimeString?: string): dayjs.Dayjs => {
  if (!dateTimeString) return dayjs();

  return dayjs.utc(dateTimeString).tz("Europe/Zagreb");
};
