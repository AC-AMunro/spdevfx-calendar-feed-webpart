export interface IFeedEvent {
    color?: string;
    title: string;
    start: Date;
    end: Date;
    url: string|undefined;
    allDay: boolean;
    category: string|undefined;
    description: string|undefined;
    location: string|undefined;
}
