import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IColor } from "office-ui-fabric-react";
import { CalendarEventRange, ICalendarEvent } from ".";

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
    getEvents: () => Promise<ICalendarEvent[]>;
}
