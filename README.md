# React Event Feed Web Parts

## Summary

This solution contains 2 web parts which use event feeds from various sources.

- Event Feed Summary
- Event Feed Calendar

It supports the following types of feeds:

- iCal
- WordPress
- RSS
- Exchange Public Calendar
- SharePoint

The solution was designed to allow other calendar feed types (or any other type of data you'd like to show as events). If you have additional feeds that you'd like to support, please contact the author or submit a pull request.

To improve performance, the web part caches the events to the user's local storage (so that it doesn't retrieve the events every time the user visits the page). You can turn off the cache by setting the cache duration to 0 minutes.

For more information about how this solution was built, including some design decisions and information on how you can extend this example to allow additional event feed provider, visit the original calendar-feed-summary author's blog: https://tahoeninjas.blog/creating-a-calendar-feed-web-part.

## Origins

This web part is a combined version of the two sample web parts on the pnp/sp-dev-fx-webparts repository

- https://github.com/pnp/sp-dev-fx-webparts/tree/master/samples/react-calendar-feed
- https://github.com/pnp/sp-dev-fx-webparts/tree/master/samples/react-calendar

## Used SharePoint Framework Version

![SPFx v1.10.0](https://img.shields.io/badge/SPFx-1.10.0-green.svg)

## Applies to

- [SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Office 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

## Prerequisites

Before you can use this web part example, you will need one of the following:

- A publicly-accessible iCal feed (i.e.: .ics)
- A publicly-accessible RSS feed of events (e.g.: Google calendar)
- A WordPress WP-FullCalendar feed
- An Exchange Public Calendar

This web part only supports anonymous external feeds. Also, make sure that your calendar includes upcoming events, as the web part will filter out evens that are earlier than today's date.

If your feed supports filtering by dates, you can specify `{s}` in the URL where the start date should be inserted, and the web part will automatically replace the `{s}` placeholder with today's date. Similarly, you can specify `{e}` in the URL where you wish the end date to be inserted, and the web part will automatically replace the placeholder for the end date, as determined by the date range you select.

## Solution

Solution|Author(s)
--------|---------
react-calendar-feed | Hugo Bernier ([Tahoe Ninjas](http://tahoeninjas.blog), @bernierh)
react-calendar-feed | Peter Paul Kirschner ([@petkir_at](https://twitter.com/petkir_at))
spdevfx-calendar-feed | Anthony Munro

## Version history

Version|Date|Comments
-------|----|--------
1.0|October 13, 2020|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- in the command line run:
  - `npm install`
  - `gulp serve`
- Insert the web part on a page
- When prompted to configure the web part, select **Configure** to launch the web part property pane.
- Select a feed type (RSS, iCal, WordPress, or Mock if using the debug solution)
- Provide the feed's URL. If using _Mock_, provide any valid URL (value will be ignored). If you wish to use a SharePoint calendar feed, provide the URL to the list (e.g.: https://yourtenant.sharepoint.com/sites/sitename/lists/eventlistname)
- Specify a date range (one week, two weeks, one month, one quarter, one year)
- Specify a maximum number of events to retrieve
- If necessary, specify to use a proxy. Use this option if you encounter issues where your feed provider does not accept your tenant URL as a CORS origin.
- If desired, specify how long (in minutes) you want to expire your users' local storage and refresh the events.
- Exclude IE11 support with gulp parameter ```--NoIE11``` this is in-case-sensitive 

## Features

This Web Part illustrates the following concepts on top of the SharePoint Framework:

- Rendering different views based on size
- Loading third-party CSS from a CDN
- Excluding mock data from production build
- Using @pnp/spfx-property-controls
- Using @pnp/spfx-controls-react
- Using localStorage to cache results locally
- Creating shared components and services
- Creating extensible services
- Using a proxy to resolve CORS issues
- Retrieving SharePoint events from a list with a filter


## Fixing ical.js issue

find build\ical.js in node_modules\ical.js and add the below to the start

if (typeof ICAL !== "undefined") { module.exports = ICAL; return; }