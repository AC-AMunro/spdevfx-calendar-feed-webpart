import { IFeedListItem } from "./IFeedListItem";

export interface IFeedListState {
    feedPropertiesDialogIsOpen: boolean;
    items: IFeedListItem[];
}