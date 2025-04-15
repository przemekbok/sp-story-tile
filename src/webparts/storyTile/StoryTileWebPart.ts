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
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";

import * as strings from 'StoryTileWebPartStrings';
import StoryTile from './components/StoryTile';
import { IStoryTileProps } from './components/IStoryTileProps';

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
  private _sp: SPFI = null;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _storyItems: IStoryItem[] = [];
  private _isLoading: boolean = true;

  public onInit(): Promise<void> {
    // Initialize PnP JS
    this._sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

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

    try {
      // Set field names - with defaults if not provided
      const titleFieldName = this.properties.titleFieldName || 'Title';
      const descFieldName = this.properties.descriptionFieldName || 'Description';
      const imageFieldName = this.properties.imageFieldName || 'Image';
      const linkFieldName = this.properties.linkFieldName || 'LinkURL';

      // Get list fields to understand the structure
      const fields = await this._sp.web.lists.getByTitle(this.properties.listName).fields();
      
      // Find the Image field to get its internal name
      const imageField = fields.find(field => 
        field.Title === imageFieldName && 
        (field.TypeAsString === 'Image' || field.TypeAsString === 'Thumbnail')
      );
      
      let imageFieldInternalName = '';
      if (imageField) {
        imageFieldInternalName = imageField.InternalName;
        console.log('Found image field internal name:', imageFieldInternalName);
      } else {
        console.warn(`Could not find image field with title '${imageFieldName}'`);
        // Try a common pattern for internal names
        imageFieldInternalName = imageFieldName.replace(/ /g, '_x0020_');
      }

      // Build a dynamic select query that includes all necessary fields
      let selectQuery = `ID,${titleFieldName},${descFieldName},${linkFieldName}`;
      
      // Try multiple approaches to get the image data
      const items = await this._sp.web.lists.getByTitle(this.properties.listName).items
        .select(selectQuery)
        .expand(imageFieldInternalName)
        .get();

      // Process items
      const processedItems: IStoryItem[] = await Promise.all(
        items.map(async (item) => {
          let imageUrl = '';
          
          // Approach 1: Try to get image via fieldValuesAsHtml
          try {
            const itemWithHtml = await this._sp.web.lists.getByTitle(this.properties.listName)
              .items.getById(item.ID)
              .fieldValuesAsHtml.get();
            
            if (itemWithHtml[imageFieldInternalName]) {
              // Extract URL from HTML
              const htmlValue = itemWithHtml[imageFieldInternalName];
              const imgSrcMatch = htmlValue.match(/src="([^"]+)"/);
              if (imgSrcMatch && imgSrcMatch[1]) {
                imageUrl = imgSrcMatch[1];
              }
            }
          } catch (error) {
            console.warn(`Approach 1 failed for item ${item.ID}:`, error);
          }

          // Approach 2: Try to access image JSON data directly
          if (!imageUrl && item[imageFieldInternalName]) {
            try {
              // Handle different potential formats of the image data
              if (typeof item[imageFieldInternalName] === 'string') {
                imageUrl = item[imageFieldInternalName];
              } else if (item[imageFieldInternalName].Url) {
                imageUrl = item[imageFieldInternalName].Url;
              } else if (item[imageFieldInternalName].serverRelativeUrl) {
                imageUrl = item[imageFieldInternalName].serverRelativeUrl;
              } else if (item[imageFieldInternalName].ServerRelativeUrl) {
                imageUrl = item[imageFieldInternalName].ServerRelativeUrl;
              }
            } catch (error) {
              console.warn(`Approach 2 failed for item ${item.ID}:`, error);
            }
          }

          // Approach 3: Try with RenderListDataAsStream
          if (!imageUrl) {
            try {
              const list = this._sp.web.lists.getByTitle(this.properties.listName);
              const renderData = await list.renderListDataAsStream({
                ViewXml: `<View><Query><Where><Eq><FieldRef Name='ID'/><Value Type='Number'>${item.ID}</Value></Eq></Where></Query></View>`
              });
              
              if (renderData && renderData.Row && renderData.Row.length > 0) {
                const row = renderData.Row[0];
                // Check for various possible field names
                const possibleFields = [
                  imageFieldInternalName,
                  `${imageFieldInternalName}.serverUrl`,
                  `${imageFieldInternalName}.urlEncoded`,
                  `${imageFieldInternalName}Url`
                ];
                
                for (const field of possibleFields) {
                  if (row[field] && !imageUrl) {
                    imageUrl = row[field];
                    break;
                  }
                }
              }
            } catch (error) {
              console.warn(`Approach 3 failed for item ${item.ID}:`, error);
            }
          }

          // Convert relative URLs to absolute URLs
          if (imageUrl && imageUrl.startsWith('/')) {
            imageUrl = `${this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')}${imageUrl}`;
          }
          
          // Use default image if no image was found
          if (!imageUrl) {
            imageUrl = require('./assets/welcome-light.png');
          }
          
          return {
            id: item.ID,
            title: item[titleFieldName] || 'No Title',
            description: item[descFieldName] || '',
            imageUrl: imageUrl,
            linkUrl: item[linkFieldName] || '#'
          };
        })
      );
      
      this._storyItems = processedItems;
    } catch (error) {
      console.error("Error fetching items from SharePoint list:", error);
      this._storyItems = [];
    }
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
                  label: 'Image Field Display Name',
                  description: 'Default: Image'
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