import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
    IPropertyPaneField,
    PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IPropertyPaneFeedListProps } from './IPropertyPaneFeedListProps';
import { IPropertyPaneFeedListInternalProps } from './IPropertyPaneFeedListInternalProps';
import { IFeedListProps } from './components/IFeedListProps';
import FeedList from './components/FeedList';

export class PropertyPaneFeedList implements IPropertyPaneField<IPropertyPaneFeedListProps> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyPaneFeedListInternalProps;
    private elem: HTMLElement;

    constructor(targetProperty: string, properties: IPropertyPaneFeedListProps) {
        this.targetProperty = targetProperty;
        this.properties = {
            key: properties.label,
            label: properties.label,
            onRender: this.onRender.bind(this),
            onDispose: this.onDispose.bind(this)
        };
    }

    public render(): void {
        if(!this.elem) {
            return;
        }

        this.onRender(this.elem);
    }

    private onDispose(element: HTMLElement): void {
        ReactDom.unmountComponentAtNode(element);
    }

    private onRender(elem: HTMLElement): void {
        if(!this.elem) {
            this.elem = elem;
        }

        const element: React.ReactElement<IFeedListProps> = React.createElement(FeedList);
        ReactDom.render(element, elem);
    }
}