import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IColumn, IDropdownOption, MaskedTextField, PrimaryButton, SelectionMode, Slider, TextField, Toggle } from 'office-ui-fabric-react';

import * as strings from "CalendarFeedWebPartStrings";

import { IAddFeedDialogProps } from './IAddFeedDialogProps';
import { IAddFeedDialogState } from './IAddFeedDialogState';
import { CalendarServiceProviderList, CalendarServiceProviderType, DateRange } from '../../../shared/services/CalendarService';

export default class AddFeedDialog extends React.Component<IAddFeedDialogProps, IAddFeedDialogState> {
    private _providerList: any[];

    constructor(props: IAddFeedDialogProps, state: IAddFeedDialogState) {
        super(props);

        this._providerList = CalendarServiceProviderList.getProviders();

        this.state = (this.props.SelectedFeed) ? {
            ...this.props.SelectedFeed,
        } : {
            FeedType: null,
            FeedUrl: '',
            MaxTotal: 0,
            DateRange: DateRange.Year,
            UseCORS: false,
            CacheDuration: 60,
            ConvertFromUTC: false
        };
    }

    public render(): JSX.Element {
        const feedTypeOptions = this._providerList.map(provider => {
            return { key: provider.key, text: provider.label };
        });
        
        const dateRangeOptions = [
            { key: DateRange.OneWeek, text: strings.DateRangeOptionWeek },
            { key: DateRange.TwoWeeks, text: strings.DateRangeOptionTwoWeeks },
            { key: DateRange.Month, text: strings.DateRangeOptionMonth },
            { key: DateRange.Quarter, text: strings.DateRangeOptionQuarter },
            { key: DateRange.Year, text: strings.DateRangeOptionUpcoming },
        ];

        return (
            <Dialog hidden={false} title="Add Feed" type={DialogType.normal} onDismiss={this.props.OnDismiss}>
                <Dropdown label={strings.FeedTypeFieldLabel}
                    options={feedTypeOptions}
                    onChange={(e, newValue?) => this.setState({ FeedType: CalendarServiceProviderType[newValue.key] })}
                    selectedKey={this.state.FeedType}
                />
                <TextField label={strings.FeedUrlFieldLabel} placeholder="https://"
                    onChange={(e, newValue?) => this.setState({ FeedUrl: newValue }) }
                    defaultValue={this.state.FeedUrl}
                />
                <Dropdown label={strings.DateRangeFieldLabel}
                    options={dateRangeOptions}
                    onChange={(e, newValue?) => { this.setState({ DateRange: DateRange[newValue.key] }); } }
                    selectedKey={DateRange[this.state.DateRange]}
                />
                <Toggle label={strings.ConvertFromUTCLabel}
                    onText={strings.ConvertFromUTCOptionYes}
                    offText={strings.ConvertFromUTCOptionNo}
                    onChange={(e, newValue?) => this.setState({ ConvertFromUTC: newValue }) }
                    defaultChecked={this.state.ConvertFromUTC}
                />
                <Toggle label={strings.UseCORSFieldLabel}
                    onText={strings.CORSOn}
                    offText={strings.CORSOff}
                    onChange={(e, newValue?) => this.setState({ UseCORS: newValue }) }
                    defaultChecked={this.state.UseCORS}
                />
                <Slider label={strings.CacheDurationFieldLabel} max={1440} min={0} step={15} showValue
                    onChange={(newValue) => this.setState({ CacheDuration: newValue }) }
                    defaultValue={this.state.CacheDuration}
                />
                <TextField label={strings.MaxTotalFieldLabel}
                    onChange={(e, newValue?) => this.setState({ MaxTotal: parseInt(newValue) }) }
                    defaultValue={this.state.MaxTotal.toString()}
                />
                <DialogFooter>
                    <PrimaryButton onClick={() => { this.props.OnSave(this.state); this.props.OnDismiss(); }} text="Save" />
                    <DefaultButton onClick={() => { this.props.OnDelete(this.state); this.props.OnDismiss(); }} text="Delete" hidden={this.props.SelectedFeed==null} />
                </DialogFooter>
            </Dialog>
        );
    }
}