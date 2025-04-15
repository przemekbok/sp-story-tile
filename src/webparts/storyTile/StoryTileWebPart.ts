import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'StoryTileWebPartStrings';
import StoryTile from './components/StoryTile';
import { IStoryTileProps } from './components/IStoryTileProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IStoryTileWebPartProps {
  title: string;
  listName: string;
  itemsPerPage: number;
  imageFieldName: string;
  titleFieldName: string;
  descriptionFieldName: string;
  linkFieldName: string;
}

export interface IStoryItem {
  id: number;
  title: string;
  description: string;
  imageUrl: string;
  linkUrl: string;
}

export default class StoryTileWebPart extends BaseClientSideWebPart<IStoryTileWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _storyItems: IStoryItem[] = [];
  private _isLoading: boolean = true;

  public render(): void {
    this._isLoading = true;
    this._fetchStoriesFromList().then(() => {
      this._isLoading = false;
      this._renderWebPart();
    }).catch(error => {
      console.error("Error fetching stories:", error);
      this._isLoading = false;
      this._renderWebPart();
    });
  }

  private _renderWebPart(): void {
    const element: React.ReactElement<IStoryTileProps> = React.createElement(
      StoryTile,
      {
        webPartTitle: this.properties.title || 'Stories',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        storyItems: this._storyItems,
        isLoading: this._isLoading,
        itemsPerPage: this.properties.itemsPerPage || 4
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async _fetchStoriesFromList(): Promise<void> {
    if (!this.properties.listName) {
      this._storyItems = []; // No list selected, empty the array
      return;
    }

    // Prepare URL to fetch items from SP list
    const listUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$select=Id,${this.properties.titleFieldName || 'Title'},${this.properties.descriptionFieldName || 'Description'},${this.properties.imageFieldName || 'ImageURL'},${this.properties.linkFieldName || 'LinkURL'}`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        listUrl, 
        SPHttpClient.configurations.v1
      );
      
      const listItems: any = await response.json();
      
      if (listItems && listItems.value) {
        // Map the SharePoint items to our IStoryItem format
        this._storyItems = listItems.value.map((item: any) => {
          return {
            id: item.Id,
            title: item[this.properties.titleFieldName || 'Title'] || 'No Title',
            description: item[this.properties.descriptionFieldName || 'Description'] || '',
            imageUrl: item[this.properties.imageFieldName || 'ImageURL'] || require('./assets/welcome-light.png'),
            linkUrl: item[this.properties.linkFieldName || 'LinkURL'] || '#'
          };
        });
      }
    } catch (error) {
      console.error("Error fetching items from SharePoint list:", error);
      this._storyItems = [];
    }
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name'
                }),
                PropertyPaneDropdown('itemsPerPage', {
                  label: 'Tiles Per View',
                  options: [
                    { key: 1, text: '1' },
                    { key: 2, text: '2' },
                    { key: 3, text: '3' },
                    { key: 4, text: '4' }
                  ],
                  selectedKey: 4
                }),
                PropertyPaneTextField('titleFieldName', {
                  label: 'Title Field Name',
                  description: 'Default: Title'
                }),
                PropertyPaneTextField('descriptionFieldName', {
                  label: 'Description Field Name',
                  description: 'Default: Description'
                }),
                PropertyPaneTextField('imageFieldName', {
                  label: 'Image URL Field Name',
                  description: 'Default: ImageURL'
                }),
                PropertyPaneTextField('linkFieldName', {
                  label: 'Link URL Field Name',
                  description: 'Default: LinkURL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}