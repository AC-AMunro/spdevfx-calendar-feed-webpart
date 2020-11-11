import { ICalendarServiceSettings } from "../../../shared/services/CalendarService/ICalendarServiceSettings";

export interface IFeedListState {
    feedPropertiesDialogIsOpen: boolean;
    selectedFeed?: ICalendarServiceSettings;
    items: ICalendarServiceSettings[];
}