import * as React from "react";
import * as ReactDom from "react-dom";

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";

// Needed for data versions
import { Version } from '@microsoft/sp-core-library';

// Localization
import * as strings from "CalendarFeedWebPartStrings";

// Calendar services
import { CalendarEventRange, ICalendarService } from "../../shared/services/CalendarService";
import { CalendarServiceProviderList } from "../../shared/services/CalendarService/CalendarServiceProviderList";

// Web part properties
import { ICalendarFeedWebPartProps } from "./CalendarFeedWebPart.types";

// Calendar Feed component
import CalendarFeed from "./components/CalendarFeed";
import { ICalendarFeedProps } from "./components/CalendarFeed.types";

// Support for theme variants
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyPaneFeedList } from "../../controls/PropertyPaneFeedList/PropertyPaneFeedList";

/**
 * Calendar Feed Web Part
 * This web part shows a calendar of events, in a film-strip (for normal views) or list view (for narrow views)
 */
export default class CalendarFeedWebPart extends BaseClientSideWebPart<ICalendarFeedWebPartProps> {
  // the list of proviers available
  private _providerList: any[];

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  constructor() {
    super();

    // get the list of providers so that we can offer it to users
    this._providerList = CalendarServiceProviderList.getProviders();
  }

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, _reject) => {
      // Consume the new ThemeProvider service
      this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();

      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

      if(this.properties.feedType !== undefined && (this.properties.providers === undefined || this.properties.providers.length == 0)) {
        if(this.properties.providers == undefined) this.properties.providers = [];

        this.properties.providers.push({
          FeedType: this.properties.feedType,
          FeedUrl: this.properties.feedUrl,
          DateRange: this.properties.dateRange,
          ConvertFromUTC: this.properties.convertFromUTC,
          UseCORS: this.properties.useCORS,
          MaxTotal: this.properties.maxTotal,
          CacheDuration: this.properties.cacheDuration
        });
      }

      resolve(undefined);
    });
  }

  /**
   * Renders the web part
   */
  public render(): void {
    // We pass the width so that the components can resize
    const { clientWidth } = this.domElement;

    // display the calendar (or the configuration screen)
    const element: React.ReactElement<ICalendarFeedProps> = React.createElement(
      CalendarFeed,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        context: this.context,
        isConfigured: this._isConfigured(),
        providers: this._getDataProviders(),
        themeVariant: this._themeVariant,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        clientWidth: clientWidth
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * We're disabling reactive property panes here because we don't want the web part to try to update events as
   * people are typing in the feed URL.
   */
  // @ts-ignore
  protected get disableReactivePropertyChanges(): boolean {
    // require an apply button on the property pane
    return true;
  }

  /**
   * Show the configuration pane
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // migrate from old version to new
    if(this.properties.providers && this.properties.providers.length > 0 && this.properties.feedType != undefined) {
      this.properties.feedType = undefined;
      this.properties.feedUrl = undefined;
      this.properties.dateRange = undefined;
      this.properties.convertFromUTC = undefined;
      this.properties.useCORS = undefined;
      this.properties.maxTotal = undefined;
      this.properties.cacheDuration = undefined;
    }

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.ProvidersLabel,
              groupFields: [
                new PropertyPaneFeedList("providers", {
                  label: strings.ProvidersLabel,
                  providers: this.properties.providers
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * If we get resized, call the Render method so that we can switch between the narrow view and the regular view
   */
  protected onAfterResize(newWidth: number): void {
    // redraw the web part
    this.render();
  }

  /**
     * Returns the data version
     */
  // @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }


  /**
   * Returns true if the web part is configured and ready to show events. If it returns false, we'll show the configuration placeholder.
   */
  private _isConfigured(): boolean {
    const { providers } = this.properties;

    return providers && providers.length > 0;
  }

  /**
   * Initialize a feed data provider from the list of existing providers
   */
  private _getDataProviders(): ICalendarService[] {
    const { providers } = this.properties;
    let dataProviders: ICalendarService[] = [];

    if(providers) {
      for(var i = 0; i < providers.length; ++i) {
        const {
          FeedType,
          FeedUrl,
          UseCORS,
          CacheDuration,
          DateRange,
          ConvertFromUTC,
          MaxTotal,
          FeedColor,
          DisplayName
        } = providers[i];

        let providerItem: any = this._providerList.filter(p => p.key === FeedType)[0];

        // make sure we got a valid provider
        if (!providerItem) {
          // return nothing. This should only happen if we removed a provider that we used to support or changed our provider keys
          continue;
        }

        let provider: ICalendarService = providerItem.initialize();
        // pass props
        provider.Context = this.context;
        provider.FeedUrl = FeedUrl;
        provider.UseCORS = UseCORS;
        provider.CacheDuration = CacheDuration;
        provider.EventRange = new CalendarEventRange(DateRange);
        provider.ConvertFromUTC = ConvertFromUTC;
        provider.MaxTotal = MaxTotal;
        provider.Color = FeedColor;
        provider.DisplayName = DisplayName;
        dataProviders.push(provider);
      }
    }

    return dataProviders;
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }
}
