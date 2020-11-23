import { Moment } from "moment";
import { ICalendarEvent } from "../services/CalendarService";

/**
 * Interface to store cached events with an expiry date
 */
export interface IFeedCache {
    events: ICalendarEvent[];
    expiry: Moment;
    feedType: string;
    feedUrl: string;
  }
  