import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";
import { ICalendarServiceSettings } from "../../shared/services/CalendarService/ICalendarServiceSettings";

export interface IPropertyPaneFeedListInternalProps extends IPropertyPaneCustomFieldProps {
    label: string;
    providers: ICalendarServiceSettings[];
}