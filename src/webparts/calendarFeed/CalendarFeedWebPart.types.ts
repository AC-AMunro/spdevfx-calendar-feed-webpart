import { DateRange, CalendarServiceProviderType } from "../../shared/services/CalendarService";

/**
 * Web part properties stored in web part configuration
 */
export interface ICalendarFeedWebPartProps {
  /**
   * The title of the web part
   */
  title: string;

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
}
