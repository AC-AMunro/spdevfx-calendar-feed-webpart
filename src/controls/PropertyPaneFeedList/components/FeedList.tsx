import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, IconButton, IIconProps, PrimaryButton, SelectionMode } from 'office-ui-fabric-react';

import AddFeedDialog from './AddFeedDialog';
import { IFeedListProps } from './IFeedListProps';
import { IFeedListState } from './IFeedListState';

import styles from './FeedList.module.scss';
import * as strings from "CalendarFeedWebPartStrings";

import { ICalendarServiceSettings } from '../../../shared/services/CalendarService/ICalendarServiceSettings';

export default class FeedList extends React.Component<IFeedListProps, IFeedListState> {
    private _columns: IColumn[];

    private editIcon: IIconProps = { iconName: 'Edit' };
    private deleteIcon: IIconProps = { iconName: 'Delete' };

    constructor(props: IFeedListProps, state: IFeedListState) {
        super(props);

        this._columns = [
            { key: 'feedColor', name: '', fieldName: 'FeedColor', minWidth: 16, maxWidth: 16, isResizable: false, onRender: (item:ICalendarServiceSettings, index, column) => {
               return <div className={styles.colorPickerRect} style={{backgroundColor: item.FeedColor}}></div>;
            } },
            { key: 'feedType', name: strings.FeedTypeFieldLabel, fieldName: 'FeedType', minWidth: 50, maxWidth: 100, isResizable: true },
            { key: 'feedUrl', name: strings.FeedUrlFieldLabel, fieldName: 'FeedUrl', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'edit', name: '', fieldName: 'edit', minWidth: 16, maxWidth: 16, isResizable: false },
            { key: 'delete', name: '', fieldName: 'delete', minWidth: 16, maxWidth: 16, isResizable: false }
        ];

        this.state = {
            feedPropertiesDialogIsOpen: false,
            selectedFeed: null,
            items: props.items ? props.items : []
        };
    }

    public render(): JSX.Element {
        
        return (
            <div>
                <DetailsList
                    columns={this._columns}
                    items={this.state.items}
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.none}
                    onRenderItemColumn={this.handleRenderItemColumn}
                />

                <PrimaryButton
                    text={strings.AddFeedLabel}
                    onClick={() => this.setState({ feedPropertiesDialogIsOpen: true, selectedFeed: null })}
                />

                {this.state.feedPropertiesDialogIsOpen ? (
                    <AddFeedDialog SelectedFeed={this.state.selectedFeed} OnSave={this.handleSave} OnDelete={this.handleDelete} OnDismiss={() => this.setState({ feedPropertiesDialogIsOpen: false, selectedFeed: null })} />
                ) : null}
            </div>
        );
    }

    private handleRenderItemColumn = (item: ICalendarServiceSettings, index: number, column: IColumn) : JSX.Element => {
        if(column.fieldName === 'edit') {
            return <IconButton iconProps={this.editIcon} title={strings.EditFeedLabel} ariaLabel={strings.EditFeedLabel} onClick={() => this.setState({ feedPropertiesDialogIsOpen: true, selectedFeed: item })} />;
        }
        else if(column.fieldName === 'delete') {
            return <IconButton iconProps={this.deleteIcon} title={strings.DeleteLabel} ariaLabel={strings.DeleteLabel} onClick={() => this.handleDelete(item)} />;
        }

        return item[column.fieldName];
    }

    private handleDelete = (item: ICalendarServiceSettings) => {
        const items = [...this.state.items];
        if(items.length > 0 && items.indexOf(item) != -1) {
            items.splice(items.indexOf(item), 1);
        }
        this.setState({ items: items, feedPropertiesDialogIsOpen: false, selectedFeed: null }, () => {
            this.props.onChange(this.state.items);
        });
    }

    private handleSave = (item: ICalendarServiceSettings) => {
        const items = [...this.state.items];

        if(!this.state.selectedFeed) {
            items.push(item);
        }
        else {
            items[items.indexOf(this.state.selectedFeed)] = item;
        }
        
        this.setState({ items: items, feedPropertiesDialogIsOpen: false, selectedFeed: null }, () => {
            this.props.onChange(this.state.items);
        });
    }
}