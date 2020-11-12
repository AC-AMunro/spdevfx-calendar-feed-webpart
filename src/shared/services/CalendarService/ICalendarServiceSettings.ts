import { IColor } from "office-ui-fabric-react";
import { DateRange } from "./CalendarEventRange";
import { CalendarServiceProviderType } from "./CalendarServiceProviderList";

export interface ICalendarServiceSettings {
    FeedColor?: string;
    FeedType: CalendarServiceProviderType;
    FeedUrl: string;
    DateRange: DateRange;
    UseCORS: boolean;
    CacheDuration: number;
    MaxTotal: number;
    ConvertFromUTC: boolean;
    DisplayName?: string;
}
