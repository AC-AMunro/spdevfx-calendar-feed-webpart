import { DisplayMode } from "@microsoft/sp-core-library";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import * as strings from "CalendarFeedWebPartStrings";
import { Calendar, momentLocalizer } from 'react-big-calendar';
import * as moment from "moment";
import {
  FocusZone, FocusZoneDirection, List, Spinner, css,
  Icon,
  DocumentCard, DocumentCardTitle, IDocumentCardPreviewProps, DocumentCardPreview, DocumentCardDetails, DocumentCardActivity,
  IPersonaSharedProps, Persona, PersonaSize, PersonaPresence,
  HoverCard, HoverCardType, IColor, ImageFit
} from "office-ui-fabric-react";
import * as React from "react";
import { CalendarServiceProviderType, ICalendarEvent, ICalendarService } from "../../../shared/services/CalendarService";
import styles from "./CalendarFeed.module.scss";
import { ICalendarFeedProps, ICalendarFeedState, IFeedCache } from "./CalendarFeed.types";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import 'react-big-calendar/lib/css/react-big-calendar.css';
require('./calendar.css');

// the key used when caching events
const CacheKey: string = "calendarFeed";

// this is the same width that the SharePoint events web parts use to render as narrow
const MaxMobileWidth: number = 480;

const localizer = momentLocalizer(moment);

/**
 * Displays a feed from a given calendar feed provider. Renders a different view for mobile/narrow web parts.
 */
export default class CalendarFeed extends React.Component<ICalendarFeedProps, ICalendarFeedState> {
  constructor(props: ICalendarFeedProps) {
    super(props);
    this.state = {
      isLoading: false,
      events: [],
      error: undefined
    };
  }

  /**
   * When components are mounted, get the events
   */
  public componentDidMount(): void {
    if (this.props.isConfigured) {
      this._loadEvents(true);
    }
  }

  /**
   * When someone changes the property pane, it triggers this event. Use it to determine if we need to refresh the events or not
   * @param prevProps The previous props before changes are applied
   * @param prevState The previous state before changes are applied
   */
  public componentDidUpdate(prevProps: ICalendarFeedProps, prevState: ICalendarFeedState): void {
    // only reload if the provider info has changed
    const prevProviders: ICalendarService[] = prevProps.providers;
    const currProviders: ICalendarService[] = this.props.providers;

    // if there isn't a current provider, do nothing
    if (currProviders === undefined || currProviders.length == 0) {
      return;
    }

    // if we didn't have a provider and now we do, we definitely need to update
    if (prevProviders === undefined || prevProviders.length == 0) {
      if (currProviders !== undefined && currProviders.length > 0) {
        this._loadEvents(false);
      }
      // there's nothing to do because there isn't a provider
      return;
    }

    let settingsHaveChanged: boolean = prevProviders.length !== currProviders.length;

    if(!settingsHaveChanged) {
      for(let prevProvider of prevProviders) {
        for(let currProvider of currProviders) {
          if(prevProvider.FeedUrl == currProvider.FeedUrl) {
            if(prevProvider.CacheDuration !== currProvider.CacheDuration ||
              prevProvider.Name !== currProvider.Name ||
              prevProvider.FeedUrl !== currProvider.FeedUrl ||
              prevProvider.EventRange.DateRange !== currProvider.EventRange.DateRange ||
              prevProvider.UseCORS !== currProvider.UseCORS ||
              prevProvider.MaxTotal !== currProvider.MaxTotal ||
              prevProvider.ConvertFromUTC !== currProvider.ConvertFromUTC ||
              prevProvider.Color !== currProvider.Color ||
              prevProvider.DisplayName !== currProvider.DisplayName) settingsHaveChanged = true;
          }
        }
      }
    }

    if (settingsHaveChanged) {
      // only load from cache if the providers haven't changed, otherwise reload.
      this._loadEvents(false);
    }
  }

  /**
   * Renders the view. There can be three different outcomes:
   * 1. Web part isn't configured and we show the placeholders
   * 2. Web part is configured and we're loading events, or
   * 3. Web part is configured and events are loaded
   */
  public render(): React.ReactElement<ICalendarFeedProps> {
    const {
      isConfigured,
    } = this.props;

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    // if we're not configured, show the placeholder
    if (!isConfigured) {
      return <Placeholder
        iconName="Calendar"
        iconText={strings.PlaceholderTitle}
        description={strings.PlaceholderDescription}
        buttonLabel={strings.ConfigureButton}
        onConfigure={this._onConfigure} />;
    }

    // we're configured, let's show stuff

    // put everything together in a nice little calendar view
    return (
      <div className={css(styles.calendar, styles.webPartChrome)} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <div className={css(styles.webPartHeader, styles.headerSmMargin)}>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
            themeVariant={this.props.themeVariant}
          />
        </div>
        {this._renderContent()}
      </div>
    );
  }

  /**
   *
   * @param {*} date
   * @memberof Calendar
   */
  public dayPropGetter(date: Date) {
    return {
        className: styles.dayPropGetter
    };
  }

  /**
   *
   * @param {*} event
   * @param {*} start
   * @param {*} end
   * @param {*} isSelected
   * @returns {*}
   * @memberof Calendar
   */
  public eventStyleGetter(event, start, end, isSelected): any {
    let style: any = {
      backgroundColor: 'white',
      borderRadius: '0px',
      opacity: 1,
      color: '#000000',//event.color,
      borderWidth: '1.1px',
      borderStyle: 'solid',
      borderColor: event.color,
      borderLeftWidth: '6px',
      display: 'block'
    };

    return {
      style: style
    };
  }

  /**
   * @private
   * @param {*} { event }
   * @returns
   * @memberof Calendar
   */
  private renderEvent({ event }) {

    const previewEventIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          // previewImageSrc: event.ownerPhoto,
          //previewIconProps: { iconName: event.fRecurrence === '0' ? 'Calendar': 'RecurringEvent', styles: { root: { color: event.color } }, className: styles.previewEventIcon },
          previewIconProps: { iconName: 'Calendar', styles: { root: { color: event.color } }, className: styles.previewEventIcon },
          height: 43
        }
      ]
    };

    /**
     * @returns {JSX.Element}
     */
    const onRenderPlainCard = (): JSX.Element => {
      return (
        <div className={styles.plainCard}>
          <DocumentCard className={styles.Documentcard}   >
            <div>
              <DocumentCardPreview {...previewEventIcon} />
            </div>
            <DocumentCardDetails>
              <div className={styles.DocumentCardDetails}>
                <DocumentCardTitle title={event.title} shouldTruncate={true} className={styles.DocumentCardTitle} styles={{ root: { color: event.color} }} />
              </div>
              {
                moment(event.start).format('YYYY/MM/DD') !== moment(event.end).format('YYYY/MM/DD') ?
                  <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('dddd')} - {moment(event.end).format('dddd')} </span>
                  :
                  <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('dddd')} </span>
              }
              {
                (event.allDay) ?
                  <span className={styles.DocumentCardTitleTime}>{strings.AllDayLabel}</span> :
                  <span className={styles.DocumentCardTitleTime}>{moment(event.start).format('h:mm A')} - {moment(event.end).format('h:mm A')}</span>
              }
              { (event.location != undefined && event.location != null && event.location != '') && <span className={styles.locationContainer}> 
                <Icon iconName='MapPin' className={styles.locationIcon} style={{ color: event.color }} />
                <span className={styles.location}>{event.location}</span>
              </span> }
              { (event.url != undefined && event.url != null && event.url != '') && <span className={styles.websiteContainer}>
                <Icon iconName='Globe' className={styles.websiteIcon} style={{ color: event.color }} />
                <a href={event.url} className={styles.website}>Visit URL</a>
              </span> }
            </DocumentCardDetails>
          </DocumentCard>
        </div>
      );
    };

    return (
      <div style={{ height: 22 }}>
        <HoverCard
          cardDismissDelay={250}
          cardOpenDelay={100}
          type={HoverCardType.plain}
          plainCardProps={{ onRenderPlainCard: onRenderPlainCard }}
          instantOpenOnClick={true}
          onCardHide={(): void => {
          }}
        >
          {event.title}
        </HoverCard>
      </div>
    );
  }

  /**
   * Render your web part content
   */
  private _renderContent(): JSX.Element {
    const isNarrow: boolean = this.props.clientWidth < MaxMobileWidth;

    const {
      displayMode
    } = this.props;
    const {
      events,
      isLoading,
      error
    } = this.state;

    const isEditMode: boolean = displayMode === DisplayMode.Edit;
    const hasErrors: boolean = error !== undefined;
    const hasEvents: boolean = events.length > 0;

    /*if (isLoading) {
      // we're currently loading
      return (<div className={styles.spinner}><Spinner label={strings.Loading} /></div>);
    }*/

    if (hasErrors) {
      // we're done loading but got some errors
      if (!isEditMode) {
        // otherwise, just show a friendly message
        return (<div className={styles.errorMessage}>{strings.ErrorMessage}</div>);
      } else {
        // render a more advanced diagnostic of what went wrong
        return this._renderError();
      }
    }

    return (
      <>
        <div className={styles.container}>
          {(isLoading) ? <div className={styles.spinnerContainer}><Spinner label={strings.Loading} className={styles.spinner} /></div> : null}
          {(!isLoading && !hasEvents) ? <div className={styles.emptyMessage}>{strings.NoEventsMessage}</div> : null }
          <Calendar
            dayPropGetter={this.dayPropGetter}
            localizer={localizer}
            selectable
            events={this.state.events}
            startAccessor="start"
            endAccessor="end"
            eventPropGetter={this.eventStyleGetter}
            components={{
              event: this.renderEvent
            }}
            defaultDate={moment().startOf('day').toDate()}
            messages={
              {
                'today': strings.TodayLabel,
                'previous': strings.PreviousLabel,
                'next': strings.NextLabel,
                'month': strings.MonthLabel,
                'week': strings.WeekLabel,
                'day': strings.DayLabel,
                'showMore': total => `+${total} ${strings.ShowMore}`
              }
            }
          />
        </div>
        <ul className={styles.legend}>
          {this.props.providers.map((provider:ICalendarService, idx) => {
            if (provider.DisplayName) return <li key={idx} style={{ borderColor: provider.Color }}>{provider.DisplayName}</li>;
          })}
        </ul>
      </>
    );
  }

  /**
   * Tries to make sense of the returned error messages and provides
   * (hopefully) helpful guidance on how to fix the issue.
   * It isn't the best piece of coding I've seen. I'm open to suggested improvements
   */
  private _renderError(): JSX.Element {
    const { error } = this.state;
    const { providers } = this.props;

    let errorMsg: string = strings.ErrorMessage;

    providers.forEach(provider => {
      
      switch (error) {
        case "Not Found":
          errorMsg = strings.ErrorNotFound;
          break;
        case "Failed to fetch":
          if (!provider.UseCORS) {
            // maybe it is because of mixed content?
            if (provider.FeedUrl.toLowerCase().substr(0, 7) === "http://") {
              errorMsg = strings.ErrorMixedContent;
            } else {
              errorMsg = strings.ErrorFailedToFetchNoProxy;
            }
          } else {
            errorMsg = strings.ErrorFailedToFetch;
          }
          break;
        default:
          // specific provider messages
          if (provider.Name === CalendarServiceProviderType.RSS) {
            switch (error) {
              case "No result":
                errorMsg = strings.ErrorRssNoResult;
                break;
              case "No root":
                errorMsg = strings.ErrorRssNoRoot;
                break;
              case "No channel":
                errorMsg = strings.ErrorRssNoChannel;
                break;
            }
          } else if (provider.Name === CalendarServiceProviderType.iCal &&
            error.indexOf("Unable to get property 'property' of undefined or null reference") !== -1) {
            errorMsg = strings.ErrorInvalidiCalFeed;
          } else if (provider.Name === CalendarServiceProviderType.WordPress && error.indexOf("Failed to read") !== -1) {
            errorMsg = strings.ErrorInvalidWordPressFeed;
          }
      }
    });

    return (<div className={styles.errorMessage} >
      <div className={styles.moreDetails}>
        {errorMsg}
      </div>
    </div>);
  }

  /**
   * When users click on the Configure button, we display the property pane
   */
  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }

  /**
   * Load events from the cache or, if expired, load from the event provider
   */
  private async _loadEvents(useCacheIfPossible: boolean): Promise<void> {
    const { providers } = this.props;

    let events:ICalendarEvent[] = [];
    let error:string = undefined;

    this.setState({
      isLoading: true,
      error: undefined,
      events: []
    });

    for(const provider of providers) {
      const { Name, FeedUrl } = provider;
      let FullCacheKey = CacheKey + ":" + FeedUrl;

      // before we do anything with the data provider, let's make sure that we don't have stuff stored in the cache
      // load from cache if: 1) we said to use cache, and b) if we have something in cache
      if ((provider.Name !== CalendarServiceProviderType.Mock || provider.CacheDuration != 0) && useCacheIfPossible && localStorage.getItem(FullCacheKey)) {

        // RegEx for matching dates
        var reISO = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*))(?:Z|(\+|-)([\d|:]*))?$/;
        var reMsAjax = /^\/Date\((d|-|.*)\)[\/|\\]$/;
        
        // Parser for field data, turn string dates into Date objects
        let cacheParser = (key, value) => {
            if ((key === 'start' || key === 'end') && typeof value === 'string') {
                var a = reISO.exec(value);
                if (a)
                    return new Date(value);
                a = reMsAjax.exec(value);
                if (a) {
                    var b = a[1].split(/[-+,.]/);
                    return new Date(b[0] ? +b[0] : 0 - +b[1]);
                }
            }
            return value;
        };

        // parse the stored JSON with our cacheParser
        let feedCache: IFeedCache = JSON.parse(localStorage.getItem(FullCacheKey), cacheParser);

        if (provider.MaxTotal > 0) {
          feedCache.events = feedCache.events.slice(0, provider.MaxTotal);
        }

        //const { Name, FeedUrl } = this.props.provider;
        let cacheStillValid: boolean = moment().isBefore(feedCache.expiry);

        // make sure the cache hasn't expired or that the settings haven't changed
        if (cacheStillValid && feedCache.feedType == Name && feedCache.feedUrl == FeedUrl) {
          events.push(...feedCache.events);
          error = undefined;
        }
      } else {
        // nothing in cache, load fresh
        if (provider) {
          try {
            let providerEvents = await provider.getEvents();

            providerEvents.map((event) => {
              if(provider.Color) event.color = provider.Color;
              
              return event;
            });

            localStorage.removeItem(FullCacheKey);

            if(provider.CacheDuration > 0) {
              const cache: IFeedCache = {
                expiry: moment().add(provider.CacheDuration, "minutes"),
                feedType: Name,
                feedUrl: FeedUrl,
                events: providerEvents
              };

              localStorage.setItem(FullCacheKey, JSON.stringify(cache));
            }

            if (provider.MaxTotal > 0) {
              providerEvents = providerEvents.slice(0, provider.MaxTotal);
            }

            events.push(...providerEvents);
          }
          catch (error) {
            console.log("Exception returned by getEvents", error.message);
            localStorage.removeItem(FullCacheKey);
            this.setState({
              isLoading: false,
              error: error.message,
              events: []
            });
          }
        }
      }
    }

    console.log(events);

    this.setState({
      isLoading: false,
      error: error,
      events: [...events]
    });
  }
}
