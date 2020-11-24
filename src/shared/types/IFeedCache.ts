import { Moment } from "moment";
import { IFeedEvent } from "../services/CalendarService";

/**
 * Interface to store cached events with an expiry date
 */
export interface IFeedCache {
    events: IFeedEvent[];
    expiry: Moment;
    feedType: string;
    feedUrl: string;
  }
  