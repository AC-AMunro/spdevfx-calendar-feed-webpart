import { IFeedEvent } from "../../services/CalendarService";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IEventCardProps {
    isEditMode: boolean;
    event: IFeedEvent;
    isNarrow: boolean;
    themeVariant?: IReadonlyTheme;
}
