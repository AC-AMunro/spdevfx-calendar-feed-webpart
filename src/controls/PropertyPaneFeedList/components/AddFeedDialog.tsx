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

        this.state = (this.props.Feed) ? {
            feedKey: this.props.Feed.key,
            feedType: this.props.Feed.feedType,
            feedUrl: this.props.Feed.feedUrl,
            maxTotal: this.props.Feed.maxTotal,
            dateRange: this.props.Feed.dateRange,
            useCORS: this.props.Feed.useCORS,
            cacheDuration: this.props.Feed.cacheDuration,
            convertFromUTC: this.props.Feed.convertFromUTC,
            dialogHidden: false
        } : {
            feedKey: undefined,
            feedType: undefined,
            feedUrl: '',
            maxTotal: 0,
            dateRange: undefined,
            useCORS: false,
            cacheDuration: 60,
            convertFromUTC: false,
            dialogHidden: false
        };
    }

    public hideDialog(): void {
        this.setState({ dialogHidden: !this.state.dialogHidden });
    }

    public render(): JSX.Element {
        const feedTypeOptions: IDropdownOption[] = this._providerList.map(provider => {
            return { key: provider.key, text: provider.label };
        });
        
        return (
            <Dialog hidden={this.state.dialogHidden} title="Add Feed" type={DialogType.normal}>
                <Dropdown label={strings.FeedTypeFieldLabel}
                    options={feedTypeOptions}
                    onChange={(e, newValue?) => this.setState({ feedType: CalendarServiceProviderType[newValue.key] })}
                />
                <TextField label={strings.FeedUrlFieldLabel} placeholder="https://"
                    onChange={(e, newValue?) => this.setState({ feedUrl: newValue }) }
                />
                <Dropdown label={strings.DateRangeFieldLabel}
                    options={[
                        { key: DateRange.OneWeek, text: strings.DateRangeOptionWeek },
                        { key: DateRange.TwoWeeks, text: strings.DateRangeOptionTwoWeeks },
                        { key: DateRange.Month, text: strings.DateRangeOptionMonth },
                        { key: DateRange.Quarter, text: strings.DateRangeOptionQuarter },
                        { key: DateRange.Year, text: strings.DateRangeOptionUpcoming },
                    ]}
                    onChange={(e, newValue?) => this.setState({ dateRange: DateRange[newValue.key] }) }
                />
                <Toggle label={strings.ConvertFromUTCLabel}
                    onText={strings.ConvertFromUTCOptionYes}
                    offText={strings.ConvertFromUTCOptionNo}
                    onChange={(e, newValue?) => this.setState({ convertFromUTC: newValue }) }
                />
                <Toggle label={strings.UseCORSFieldLabel}
                    onText={strings.CORSOn}
                    offText={strings.CORSOff}
                    onChange={(e, newValue?) => this.setState({ useCORS: newValue }) }
                />
                <Slider label={strings.CacheDurationFieldLabel} max={1440} min={0} step={15} showValue
                    onChange={(newValue) => this.setState({ cacheDuration: newValue }) }
                />
                <TextField label={strings.MaxTotalFieldLabel}
                    onChange={(e, newValue?) => this.setState({ maxTotal: parseInt(newValue) }) }
                />
                <DialogFooter>
                    <PrimaryButton onClick={() => { this.props.OnSave(this.state); this.hideDialog(); }} text="Save" />
                </DialogFooter>
            </Dialog>
        );
    }
}