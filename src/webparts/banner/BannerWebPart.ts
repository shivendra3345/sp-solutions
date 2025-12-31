import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BannerWebPartStrings';
import Banner from './components/Banner';
import { IBannerProps } from './components/IBannerProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBannerWebPartProps {
  description: string;
  /** Show the welcome panel for new joiners */
  showWelcome?: boolean;
  /** Comma-separated list of site URLs to exclude from events */
  excludedSites?: string;
  /** Show events carousel */
  showEvents?: boolean;
  /** Show news/announcements */
  showNews?: boolean;
  /** Titles configurable from property pane */
  welcomeTitle?: string;
  eventsTitle?: string;
  newsTitle?: string;
  /** Toggle whether the header extension is enabled on the page */
  enableHeaderExtension?: boolean;
  context: WebPartContext;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBannerProps> = React.createElement(
      Banner,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        showWelcome: this.properties.showWelcome !== undefined ? this.properties.showWelcome : true,
        excludedSites: this.properties.excludedSites || '',
        showEvents: this.properties.showEvents !== undefined ? this.properties.showEvents : true,
        showNews: this.properties.showNews !== undefined ? this.properties.showNews : true,
        enableHeaderExtension: this.properties.enableHeaderExtension !== undefined ? this.properties.enableHeaderExtension : true,
        welcomeTitle: this.properties.welcomeTitle || 'Welcome New Joiners',
        eventsTitle: this.properties.eventsTitle || 'Events',
        newsTitle: this.properties.newsTitle || 'News & Announcements'
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
                ,
                PropertyPaneToggle('showNews', {
                  label: 'Show News Section',
                  onText: 'Shown',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showEvents', {
                  label: 'Show Events Carousel',
                  onText: 'Shown',
                  offText: 'Hidden'
                }),
                PropertyPaneTextField('newsTitle', {
                  label: 'News Section Title'
                }),
                PropertyPaneTextField('welcomeTitle', {
                  label: 'Welcome Panel Title'
                }),
                PropertyPaneTextField('eventsTitle', {
                  label: 'Events Carousel Title'
                }),
                PropertyPaneToggle('enableHeaderExtension', {
                  label: 'Enable Header Extension',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('showWelcome', {
                  label: 'Show Welcome Panel',
                  onText: 'Shown',
                  offText: 'Hidden'
                })
                ,
                PropertyPaneTextField('excludedSites', {
                  label: 'Exclude Sites (comma-separated URLs or serverRelativePaths)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
