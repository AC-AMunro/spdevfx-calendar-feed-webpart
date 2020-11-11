import { DateRange } from "./CalendarEventRange";
import { CalendarServiceProviderType } from "./CalendarServiceProviderList";

export interface ICalendarServiceSettings {
    FeedType: CalendarServiceProviderType;
    FeedUrl: string;
    DateRange: DateRange;
    UseCORS: boolean;
    CacheDuration: number;
    MaxTotal: number;
    ConvertFromUTC: boolean;
}
