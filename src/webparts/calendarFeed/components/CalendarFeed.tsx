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
  HoverCard, HoverCardType
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
    const prevProvider: ICalendarService = prevProps.provider;
    const currProvider: ICalendarService = this.props.provider;

    // if there isn't a current provider, do nothing
    if (currProvider === undefined) {
      return;
    }

    // if we didn't have a provider and now we do, we definitely need to update
    if (prevProvider === undefined) {
      if (currProvider !== undefined) {
        this._loadEvents(false);
      }

      // there's nothing to do because there isn't a provider
      return;
    }

    const settingsHaveChanged: boolean = prevProvider.CacheDuration !== currProvider.CacheDuration ||
      prevProvider.Name !== currProvider.Name ||
      prevProvider.FeedUrl !== currProvider.FeedUrl ||
      prevProvider.Name !== currProvider.Name ||
      prevProvider.EventRange.DateRange !== currProvider.EventRange.DateRange ||
      prevProvider.UseCORS !== currProvider.UseCORS ||
      prevProvider.MaxTotal !== currProvider.MaxTotal ||
      prevProvider.ConvertFromUTC !== currProvider.ConvertFromUTC;

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
        <div className={styles.container}>
          {this._renderContent()}
        </div>
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
      color: event.color,
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
          height: 43,
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
              { (event.location != undefined && event.location != null && event.location != '') && <span> 
                <Icon iconName='MapPin' className={styles.locationIcon} style={{ color: event.color }} />
                <span className={styles.location}>{event.location}</span>
              </span> }
              { (event.url != undefined && event.url != null && event.url != '') && <span>
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
          cardDismissDelay={1000}
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

    if (isLoading) {
      // we're currently loading
      return (<div className={styles.spinner}><Spinner label={strings.Loading} /></div>);
    }

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

    if (!hasEvents) {
      // we're done loading, no errors, but have no events
      return (<div className={styles.emptyMessage}>{strings.NoEventsMessage}</div>);
    }

    return (<Calendar
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
    />);
  }

  /**
   * Tries to make sense of the returned error messages and provides
   * (hopefully) helpful guidance on how to fix the issue.
   * It isn't the best piece of coding I've seen. I'm open to suggested improvements
   */
  private _renderError(): JSX.Element {
    const { error } = this.state;
    const { provider } = this.props;
    let errorMsg: string = strings.ErrorMessage;
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
    const { Name, FeedUrl } = this.props.provider;
    const FullCacheKey = CacheKey + ":" + FeedUrl;

    if(this.props.provider.Name == 'Mock' || this.props.provider.CacheDuration == 0) {
      useCacheIfPossible = false;
    }

    // before we do anything with the data provider, let's make sure that we don't have stuff stored in the cache
    // load from cache if: 1) we said to use cache, and b) if we have something in cache
    if (useCacheIfPossible && localStorage.getItem(FullCacheKey)) {

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

      if (this.props.provider.MaxTotal > 0) {
        feedCache.events = feedCache.events.slice(0, this.props.provider.MaxTotal);
      }

      //const { Name, FeedUrl } = this.props.provider;
      const cacheStillValid: boolean = moment().isBefore(feedCache.expiry);

      // make sure the cache hasn't expired or that the settings haven't changed
      if (cacheStillValid && feedCache.feedType === Name && feedCache.feedUrl === FeedUrl) {
        this.setState({
          isLoading: false,
          error: undefined,
          events: feedCache.events
        });
        return;
      }
    }

    // nothing in cache, load fresh
    if (this.props.provider) {
      this.setState({
        isLoading: true
      });

      try {
        let events = await this.props.provider.getEvents();

        if(useCacheIfPossible) {
          const cache: IFeedCache = {
            expiry: moment().add(this.props.provider.CacheDuration, "minutes"),
            feedType: Name,
            feedUrl: FeedUrl,
            events: events
          };

          localStorage.setItem(FullCacheKey, JSON.stringify(cache));
        }

        if (this.props.provider.MaxTotal > 0) {
          events = events.slice(0, this.props.provider.MaxTotal);
        }

        // don't cache in the case of errors
        this.setState({
          isLoading: false,
          error: undefined,
          events: events
        });

        return;
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
