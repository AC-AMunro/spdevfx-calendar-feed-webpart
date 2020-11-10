import { IAddFeedDialogState } from "./IAddFeedDialogState";
import { IFeedListItem } from "./IFeedListItem";

export interface IAddFeedDialogProps {
    Feed?: IFeedListItem;
    OnSave: (item?: IAddFeedDialogState) => void;
}