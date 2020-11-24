/**
 * CalendarFeed Types
 * Contains the various types used by the component.
 * (I like to  keep my props and state in a separate ".types"
 * file because that's what the Office UI Fabric team does and
 * I kinda liked it.
 */
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFeedEvent, ICalendarService } from "../../../shared/services/CalendarService";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IFeedCache } from "../../../shared/types";

/**
 * The props for the calendar feed component
 */
export interface ICalendarFeedProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
  isConfigured: boolean;
  providers: ICalendarService[];
  themeVariant: IReadonlyTheme;
  clientWidth: number;
}

/**
 * The state for the calendar feed component
 */
export interface ICalendarFeedState {
  providers: ICalendarProvider[];
  events: ICalendarEvent[];
  error: any|undefined;
  isLoading: boolean;
}

export interface ICalendarProvider extends ICalendarService {
  visible: boolean;
}

export interface ICalendarEvent extends IFeedEvent {
  provider: string;
  visible: boolean;
}

export interface ICalendarFeedCache extends IFeedCache {
  events: ICalendarEvent[];
}