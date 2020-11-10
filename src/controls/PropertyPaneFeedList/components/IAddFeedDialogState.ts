import { CalendarServiceProviderType, DateRange } from "../../../shared/services/CalendarService";

export interface IAddFeedDialogState {
  feedKey?: string;
  /**
   * The URL where to get the feed from
   */
  feedUrl: string;

  /**
   * The type of feed provider
   */
  feedType: CalendarServiceProviderType;

  /**
   * Maximum total number of events to load
   */
  maxTotal: number;

  /**
   * Date range to retrieve events
   */
  dateRange: DateRange;

  /**
   * use CORS proxy when retrieving events
   */
  useCORS: boolean;

  /**
   * how long to cache events for
   */
  cacheDuration: number;

  /**
   * Indicates the dates received from feeds do not specify a timezone
   */
  convertFromUTC: boolean;

  dialogHidden: boolean;
}