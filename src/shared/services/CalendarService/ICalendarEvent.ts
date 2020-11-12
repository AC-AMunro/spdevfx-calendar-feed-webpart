import { IColor } from "office-ui-fabric-react";

export interface ICalendarEvent {
    color?: string;
    title: string;
    start: Date;
    end: Date;
    url: string|undefined;
    allDay: boolean;
    category: string|undefined;
    description: string|undefined;
    location: string|undefined;
}
