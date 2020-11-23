import * as React from 'react';
import { Callout, ColorPicker, DefaultButton, Dropdown, Icon, Label, Panel, PanelType, PrimaryButton, Slider, TextField, Toggle, TooltipHost } from '@fluentui/react';

import * as strings from "CalendarFeedWebPartStrings";

import { IAddFeedDialogProps } from './IAddFeedDialogProps';
import { IAddFeedDialogState } from './IAddFeedDialogState';
import { CalendarServiceProviderList, CalendarServiceProviderType, DateRange } from '../../../shared/services/CalendarService';
import styles from './AddFeedDialog.module.scss';

export default class AddFeedDialog extends React.Component<IAddFeedDialogProps, IAddFeedDialogState> {
    private _providerList: any[];

    constructor(props: IAddFeedDialogProps, state: IAddFeedDialogState) {
        super(props);

        this._providerList = CalendarServiceProviderList.getProviders();

        this.state = (this.props.SelectedFeed) ? {
            ...this.props.SelectedFeed,
            MaxTotal: (this.props.SelectedFeed && this.props.SelectedFeed.MaxTotal) ? this.props.SelectedFeed.MaxTotal : 0,
            showColorPicker: false
        } : {
            DisplayName: '',
            FeedType: null,
            FeedUrl: '',
            MaxTotal: 0,
            DateRange: DateRange.Year,
            UseCORS: false,
            CacheDuration: 60,
            ConvertFromUTC: false,
            FeedColor: '#000000',
            showColorPicker: false
        };
    }

    /**
   * Validates a URL when users type them in the configuration pane.
   * @param feedUrl The URL to validate
   */
    private _validateFeedUrl(feedUrl: string): string {
        if (this.state.FeedType === CalendarServiceProviderType.Mock) {
            // we don't need a URL for mock feeds
            return '';
        }

        // Make sure the feed isn't empty or null
        if (feedUrl === null ||
            feedUrl.trim().length === 0) {
            return strings.FeedUrlValidationNoUrl;
        }

        if (!feedUrl.match(/(http|https):\/\/(\w+:{0,1}\w*)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%!\-\/]))?/)) {
            return strings.FeedUrlValidationInvalidFormat;
        }

        // No errors
        return '';
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

        const isMock: boolean = this.state.FeedType === CalendarServiceProviderType.Mock;

        return (
            <Panel headerText={(this.props.SelectedFeed) ? strings.EditFeedLabel : strings.AddFeedLabel}
                type={PanelType.medium} isOpen={true} onDismiss={this.props.OnDismiss}
                onRenderFooterContent={this.onRenderFooterContent} isFooterAtBottom={true}>
                <Dropdown id="feedTypeField" label={strings.FeedTypeFieldLabel}
                    options={feedTypeOptions}
                    onChange={(e, newValue?) => this.setState({ FeedType: CalendarServiceProviderType[newValue.key] })}
                    selectedKey={this.state.FeedType} disabled={this.props.SelectedFeed != null}
                    required={true}
                />
                {!isMock ? <TextField id="feedUrlField" label={strings.FeedUrlFieldLabel} placeholder="https://"
                    onChange={(e, newValue?) => this.setState({ FeedUrl: newValue }) }
                    defaultValue={this.state.FeedUrl} disabled={this.props.SelectedFeed != null}
                    onGetErrorMessage={this._validateFeedUrl.bind(this)}
                    required={true}
                /> : null }
                <TextField id="feedDisplayNameField" label={strings.FeedDisplayNameFieldLabel}
                    onChange={(e, newValue?) => this.setState({ DisplayName: newValue }) }
                    defaultValue={this.state.DisplayName}
                    required={true}
                />
                <Dropdown label={strings.DateRangeFieldLabel}
                    options={dateRangeOptions}
                    onChange={(e, newValue?) => { console.log(newValue.key); this.setState({ DateRange: newValue.key as DateRange }); } }
                    selectedKey={this.state.DateRange}
                />
                <Toggle label={
                    <div>
                        {strings.ConvertFromUTCLabel}{' '}
                        <TooltipHost content={strings.ConvertFromUTCFieldDescription}>
                            <Icon iconName="Info" aria-label="Info tooltip" />
                        </TooltipHost>
                    </div>
                    }
                    onText={strings.ConvertFromUTCOptionYes}
                    offText={strings.ConvertFromUTCOptionNo}
                    onChange={(e, newValue?) => this.setState({ ConvertFromUTC: newValue }) }
                    defaultChecked={this.state.ConvertFromUTC}
                />
                {!isMock ? <Toggle label={
                    <div>
                        {strings.UseCORSFieldLabel}{' '}
                        <TooltipHost content={strings.UseCorsFieldDescription}>
                            <Icon iconName="Info" aria-label="Info tooltip" />
                        </TooltipHost>
                    </div>
                    }
                    onText={strings.CORSOn}
                    offText={strings.CORSOff}
                    onChange={(e, newValue?) => this.setState({ UseCORS: newValue }) }
                    defaultChecked={this.state.UseCORS}
                /> : null}
                <Slider label={strings.CacheDurationFieldLabel} max={1440} min={0} step={15} showValue
                    onChange={(newValue) => this.setState({ CacheDuration: newValue }) }
                    defaultValue={this.state.CacheDuration}
                />
                <TextField label={strings.MaxTotalFieldLabel}
                    onChange={(e, newValue?) => this.setState({ MaxTotal: parseInt(newValue) }) }
                    defaultValue={this.state.MaxTotal.toString()}
                    description={strings.MaxTotalFieldDescription}
                />
                <Label>Color</Label>
                <div className={styles.colorPickerRect} style={{ backgroundColor: this.state.FeedColor }} onClick={() => this.setState({showColorPicker: !this.state.showColorPicker})}></div>
                {this.state.showColorPicker ?
                    <Callout target={`.${styles.colorPickerRect}`} onDismiss={this.hideColorPicker}>
                        <ColorPicker alphaSliderHidden={true} onChange={(ev, newValue) => this.setState({ FeedColor: '#'+newValue.hex }) } color={this.state.FeedColor} />
                    </Callout>
                : null }
            </Panel>
        );
    }

    private hideColorPicker = () => {
        this.setState({ showColorPicker: false });
    }

    private onRenderFooterContent = () => {
        return (<>
            <PrimaryButton disabled={
                (this.state.DisplayName && this.state.DisplayName.length == 0) ||
                (this.state.FeedUrl && this.state.FeedUrl.length == 0 || this._validateFeedUrl(this.state.FeedUrl) != '') ||
                this.state.FeedType == null} onClick={() => { this.props.OnSave(this.state); this.props.OnDismiss(); }} text="Save" />
            {this.props.SelectedFeed ? <DefaultButton onClick={() => { this.props.OnDelete(this.state); this.props.OnDismiss(); }} text="Delete" /> : null }
        </>);
    }
}