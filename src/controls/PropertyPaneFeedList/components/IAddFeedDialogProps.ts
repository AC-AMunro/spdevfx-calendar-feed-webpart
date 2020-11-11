import { ICalendarServiceSettings } from "../../../shared/services/CalendarService/ICalendarServiceSettings";
import { IAddFeedDialogState } from "./IAddFeedDialogState";

export interface IAddFeedDialogProps {
    SelectedFeed?: ICalendarServiceSettings;
    OnSave: (item?: IAddFeedDialogState) => void;
    OnDelete: (key: any) => void;
    OnDismiss: () => void;
}