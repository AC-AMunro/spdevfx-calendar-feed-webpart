import { CalendarServiceProviderType, DateRange } from "../../../shared/services/CalendarService";
import { ICalendarServiceSettings } from "../../../shared/services/CalendarService/ICalendarServiceSettings";

export interface IAddFeedDialogState extends ICalendarServiceSettings {
  showColorPicker: boolean;
}