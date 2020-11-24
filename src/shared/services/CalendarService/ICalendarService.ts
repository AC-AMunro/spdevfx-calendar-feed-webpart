import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CalendarEventRange, IFeedEvent } from ".";

export interface ICalendarService {
    Context: WebPartContext;
    FeedUrl: string;
    EventRange: CalendarEventRange;
    UseCORS: boolean;
    CacheDuration: number;
    MaxTotal: number;
    ConvertFromUTC: boolean;
    Name: string;
    DisplayName?: string;
    Color?: string;
    getEvents: () => Promise<IFeedEvent[]>;
}
