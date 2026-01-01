import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'BannerWebPartStrings';
import Banner from './components/Banner';
import { IBannerProps } from './components/IBannerProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import CalendarPicker from './components/PropertyPane/CalendarPicker';

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
  /** Comma-separated list of selected calendar lists encoded as encodeURIComponent(webUrl) + '::' + listId */
  selectedListIds?: string;
  /** Comma-separated list of selected calendar display titles for UI */
  selectedListTitles?: string;
  /** Comma-separated list of internal field names to try for event description/body (priority order) */
  descriptionFieldNames?: string;
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
        selectedListIds: this.properties.selectedListIds || '',
        selectedListTitles: this.properties.selectedListTitles || '',
        descriptionFieldNames: this.properties.descriptionFieldNames || '',
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
                ,
                PropertyPaneTextField('selectedListIds', {
                  label: 'Selected Calendars (webUrl::listId CSV)'
                }),
                PropertyPaneTextField('selectedListTitles', {
                  label: 'Selected Calendars (display names)'
                }),
                PropertyPaneTextField('descriptionFieldNames', {
                  label: 'Description field internal names (CSV)'
                }),
                PropertyPaneButton('discoverCalendars', {
                  text: 'Discover Calendars Across Sites',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._discoverCalendars.bind(this)
                }),
                PropertyPaneButton('selectCalendars', {
                  text: 'Select Calendars...',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._openCalendarPicker.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async _discoverCalendars(): Promise<void> {
    try {
      const context = this.context as any;
      const siteRoot = context.pageContext.site && context.pageContext.site.absoluteUrl ? context.pageContext.site.absoluteUrl : context.pageContext.web.absoluteUrl;
      const webUrls: string[] = [siteRoot];
      for (let i = 0; i < webUrls.length; i++) {
        try {
          const wurl = `${webUrls[i]}/_api/web/webs?$select=Url`;
          const r = await context.spHttpClient.get(wurl, SPHttpClient.configurations.v1);
          if (!r || !r.ok) continue;
          const j = await r.json();
          const subs = j.value || [];
          subs.forEach((s: any) => {
            if (s && s.Url) webUrls.push(s.Url);
          });
        } catch (err) {
          console.debug('BannerWebPart: failed to enumerate subwebs for', webUrls[i], err);
        }
      }

      const discovered: string[] = [];
      // look for lists with name containing 'event' or 'calendar' or BaseTemplate 106
      for (const webUrl of webUrls) {
        try {
          const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate`;
          const lr = await context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1);
          if (!lr || !lr.ok) continue;
          const lj = await lr.json();
          const allLists = lj.value || [];
          for (const l of allLists) {
            try {
              const id = (l.Id || l.ID || l.id || '').toString();
              const title = (l.Title || '').toString().toLowerCase();
              const base = l.BaseTemplate || 0;
              if (base === 106 || title.indexOf('event') !== -1 || title.indexOf('calendar') !== -1) {
                discovered.push(`${encodeURIComponent(webUrl)}::${id}`);
              }
            } catch (inner) { }
          }
        } catch (e) {
          console.debug('BannerWebPart: failed to fetch lists for', webUrl, e);
        }
      }

      this.properties.selectedListIds = discovered.join(',');
      this.context.propertyPane.refresh();
      this.render();
      console.debug('BannerWebPart: discovered calendars', discovered.length);
    } catch (err) {
      console.error('BannerWebPart: discover calendars failed', err);
    }
  }

  private _openCalendarPicker(): void {
    try {
      const container = document.createElement('div');
      container.className = 'sp-solutions-calendar-picker-host';
      document.body.appendChild(container);

      const onSave = (csv: string, titlesCsv: string) => {
        this.properties.selectedListIds = csv || '';
        this.properties.selectedListTitles = titlesCsv || '';
        this.context.propertyPane.refresh();
        this.render();
        try { ReactDom.unmountComponentAtNode(container); } catch (e) { /* ignore */ }
        if (container.parentNode) container.parentNode.removeChild(container);
      };

      const onDismiss = () => {
        try { ReactDom.unmountComponentAtNode(container); } catch (e) { /* ignore */ }
        if (container.parentNode) container.parentNode.removeChild(container);
      };

      ReactDom.render(React.createElement(CalendarPicker, { context: this.context, currentSelectedCsv: this.properties.selectedListIds || '', onSave, onDismiss }), container);
      console.debug('BannerWebPart: calendar picker mounted to DOM', container);
    } catch (e) {
      console.error('BannerWebPart: failed to open calendar picker', e);
    }
  }
}
