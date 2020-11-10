import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogFooter, DialogType, Dropdown, IColumn, IDropdownOption, MaskedTextField, PrimaryButton, SelectionMode, Slider, TextField, Toggle } from 'office-ui-fabric-react';
import { CalendarServiceProviderList, CalendarServiceProviderType, DateRange } from "../../../shared/services/CalendarService";

import AddFeedDialog from './AddFeedDialog';
import { IFeedListItem } from './IFeedListItem';
import { IFeedListProps } from './IFeedListProps';
import { IFeedListState } from './IFeedListState';

import * as strings from "CalendarFeedWebPartStrings";
import { IAddFeedDialogState } from './IAddFeedDialogState';

export default class FeedList extends React.Component<IFeedListProps, IFeedListState> {
    private _providerList: any[];

    private _columns: IColumn[];

    constructor(props: IFeedListProps, state: IFeedListState) {
        super(props);

        this._providerList = CalendarServiceProviderList.getProviders();

        this._columns = [
            { key: 'feedType', name: 'Type', fieldName: 'feedType', minWidth: 50, maxWidth: 100, isResizable: true },
            { key: 'feedUrl', name: 'Url', fieldName: 'feedUrl', minWidth: 100, maxWidth: 200, isResizable: true }
        ];

        this.state = {
            feedPropertiesDialogIsOpen: false,
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
                />

                <PrimaryButton
                    text='Add Feed'
                    onClick={() => this.setState({ feedPropertiesDialogIsOpen: true })}
                />

                {this.state.feedPropertiesDialogIsOpen ? (
                    <AddFeedDialog OnSave={this.handleSave} />
                ) : null}
            </div>
        );
    }

    private handleSave = (item: IAddFeedDialogState) => {
        if(item.feedKey === undefined) {
            this.setState({ items: [...this.state.items, {
                key: item.feedType + item.feedUrl,
                feedType: item.feedType,
                feedUrl: item.feedUrl,
                maxTotal: item.maxTotal,
                dateRange: item.dateRange,
                useCORS: item.useCORS,
                cacheDuration: item.cacheDuration,
                convertFromUTC: item.convertFromUTC
            }] });
        }
        else {
            var items = this.state.items;
            items.map((i) => {
                if(i.key == item.feedKey)
                    return {
                        key: item.feedType + item.feedUrl,
                        feedType: item.feedType,
                        feedUrl: item.feedUrl,
                        maxTotal: item.maxTotal,
                        dateRange: item.dateRange,
                        useCORS: item.useCORS,
                        cacheDuration: item.cacheDuration,
                        convertFromUTC: item.convertFromUTC
                    }
                else
                    return i;
            });
            this.setState({ items: items });
        }
    }
}