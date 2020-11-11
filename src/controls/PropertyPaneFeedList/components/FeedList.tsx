import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, IconButton, IIconProps, PrimaryButton, SelectionMode } from 'office-ui-fabric-react';

import AddFeedDialog from './AddFeedDialog';
import { IFeedListProps } from './IFeedListProps';
import { IFeedListState } from './IFeedListState';

import * as settingsStrings from "CalendarServiceSettingsStrings";
import { ICalendarServiceSettings } from '../../../shared/services/CalendarService/ICalendarServiceSettings';

export default class FeedList extends React.Component<IFeedListProps, IFeedListState> {
    private _columns: IColumn[];

    private editIcon: IIconProps = { iconName: 'Edit' };
    private deleteIcon: IIconProps = { iconName: 'Delete' };

    constructor(props: IFeedListProps, state: IFeedListState) {
        super(props);

        this._columns = [
            { key: 'feedType', name: settingsStrings.FeedTypeFieldLabel, fieldName: 'FeedType', minWidth: 50, maxWidth: 100, isResizable: true },
            { key: 'feedUrl', name: settingsStrings.FeedUrlFieldLabel, fieldName: 'FeedUrl', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'edit', name: '', fieldName: 'edit', minWidth:25, isResizable: false },
            { key: 'delete', name: '', fieldName: 'delete', minWidth: 25, isResizable: false }
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
                    text={settingsStrings.AddFeed}
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
            return <IconButton iconProps={this.editIcon} title="Edit" ariaLabel="Edit" onClick={() => this.setState({ feedPropertiesDialogIsOpen: true, selectedFeed: item })} />;
        }
        else if(column.fieldName === 'delete') {
            return <IconButton iconProps={this.deleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this.handleDelete(item)} />;
        }

        return item[column.fieldName];
    }

    private handleDelete = (item: ICalendarServiceSettings) => {
        const items = [...this.state.items];
        if(items.length > 0) {
            items.splice(items.indexOf(item), 1);
        }
        this.setState({ items: items, feedPropertiesDialogIsOpen: false, selectedFeed: null });
    }

    private handleSave = (item: ICalendarServiceSettings) => {
        const items = [...this.state.items];

        if(!this.state.selectedFeed) {
            items.push({
                FeedType: item.FeedType,
                FeedUrl: item.FeedUrl,
                MaxTotal: item.MaxTotal,
                DateRange: item.DateRange,
                UseCORS: item.UseCORS,
                CacheDuration: item.CacheDuration,
                ConvertFromUTC: item.ConvertFromUTC
            });
        }
        else {
            items[items.indexOf(this.state.selectedFeed)] = {
                FeedType: item.FeedType,
                FeedUrl: item.FeedUrl,
                MaxTotal: item.MaxTotal,
                DateRange: item.DateRange,
                UseCORS: item.UseCORS,
                CacheDuration: item.CacheDuration,
                ConvertFromUTC: item.ConvertFromUTC
            };
        }
        
        this.setState({ items: items, feedPropertiesDialogIsOpen: false, selectedFeed: null });
    }
}