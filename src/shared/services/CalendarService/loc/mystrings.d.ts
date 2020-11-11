declare interface ICalendarServicesStrings {
  SharePointProviderName: string;
  WordPressProviderName: string;
  ExchangeProviderName: string;
  iCalProviderName: string;
  RSSProviderName: string;
  MockProviderName: string;
}

declare module 'CalendarServicesStrings' {
  const strings: ICalendarServicesStrings;
  export = strings;
}

declare interface ICalendarServiceSettingsStrings {
  AddFeed: string;
  EditFeed: string;
  UseCorsFieldDescription: string;
  ConvertFromUTCFieldDescription: string;
  ConvertFromUTCOptionNo: string;
  ConvertFromUTCOptionYes: string;
  ConvertFromUTCLabel: string;
  MaxTotalFieldDescription: string;
  MaxTotalFieldLabel: string;
  FilmStripAriaLabel: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AllItemsUrlFieldLabel: string;
  FeedUrlFieldLabel: string;
  FeedTypeFieldLabel: string;
  PlaceholderTitle: string;
  PlaceholderDescription: string;
  ConfigureButton: string;
  FeedTypeOptionGoogle: string;
  FeedTypeOptioniCal: string;
  FeedTypeOptionRSS: string;
  FeedTypeOptionWordPress: string;
  FeedUrlCallout: string;
  DateRangeFieldLabel: string;
  DateRangeOptionUpcoming: string;
  DateRangeOptionWeek: string;
  DateRangeOptionTwoWeeks: string;
  DateRangeOptionMonth: string;
  DateRangeOptionQuarter: string;
  UseCORSFieldLabel: string;
  UseCORSFieldCallout: string;
  UseCORSFieldCalloutDisabled: string;
  CORSOn: string;
  CORSOff: string;
  AdvancedGroupName: string;
  FocusZoneAriaLabelReadMode: string;
  FocusZoneAriaLabelEditMode: string;
  CacheDurationFieldLabel: string;
  CacheDurationFieldCallout: string;
  FeedUrlValidationNoUrl: string;
  FeedUrlValidationInvalidFormat: string;
  ErrorMessage: string;
  ErrorNotFound: string;
  ErrorMixedContent: string;
  ErrorFailedToFetch: string;
  ErrorFailedToFetchNoProxy: string;
  ErrorRssNoResult: string;
  ErrorRssNoRoot: string;
  ErrorRssNoChannel: string;
  ErrorInvalidiCalFeed: string;
  ErrorInvalidWordPressFeed: string;
  AddToCalendarAriaLabel: string;
  AddToCalendarButtonLabel: string;
  AllDayDateFormat: string;
  LocalizedTimeFormat: string;
  FeedSettingsGroupName: string;
}

declare module 'CalendarServiceSettingsStrings' {
  const strings: ICalendarServiceSettingsStrings;
  export = strings;
}