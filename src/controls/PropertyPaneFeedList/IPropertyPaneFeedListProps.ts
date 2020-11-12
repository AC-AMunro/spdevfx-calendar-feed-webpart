import { ICalendarServiceSettings } from "../../shared/services/CalendarService/ICalendarServiceSettings";

export interface IPropertyPaneFeedListProps {
    label: string;
    providers: ICalendarServiceSettings[];
}