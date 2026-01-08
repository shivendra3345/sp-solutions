import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DptPageWebPartStrings';
import DptPage from './components/DptPage';
import { IDptPageProps } from './components/IDptPageProps';

export interface IDptPageWebPartProps {
  description: string;
  bannerLayout: 'carousel' | 'tile';
  bannerListTitle: string;
  calendarLayout: 'grid' | 'carousel';
  calendarSource: 'outlook' | 'sharepoint' | 'both';
  groupId?: string;
  spSiteUrl?: string;
  spListTitle?: string;
  teamMembersListTitle?: string;
  teamMembersTitle?: string;
  teamMembersItemsPerRow?: number;
  showPageBanner: boolean;
  showCalendar: boolean;
  showTeamMembers: boolean;
  quickLinksListTitle?: string;
  quickLinksTitle?: string;
  showQuickLinks: boolean;
}

const bannerLayoutOptions: IPropertyPaneDropdownOption[] = [
  { key: 'carousel', text: 'Carousel (1-Slide)' },
  { key: 'tile', text: 'Tile View (1 Large + 2 Small)' }
];

const calendarLayoutOptions: IPropertyPaneDropdownOption[] = [
  { key: 'grid', text: 'Calendar Grid View' },
  { key: 'carousel', text: 'Carousel View' }
];

const calendarSourceOptions: IPropertyPaneDropdownOption[] = [
  { key: 'outlook', text: 'Outlook Group Calendar' },
  { key: 'sharepoint', text: 'SharePoint Calendar' },
  { key: 'both', text: 'Both Outlook & SharePoint' }
];

export default class DptPageWebPart extends BaseClientSideWebPart<IDptPageWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDptPageProps> = React.createElement(
      DptPage,
      {
        description: this.properties.description,
        context: this.context,
        bannerLayout: this.properties.bannerLayout || 'carousel',
        bannerListTitle: this.properties.bannerListTitle || 'BannerImages',
        calendarLayout: this.properties.calendarLayout || 'grid',
        calendarSource: this.properties.calendarSource || 'outlook',
        groupId: this.properties.groupId || '',
        spSiteUrl: this.properties.spSiteUrl || this.context.pageContext.web.absoluteUrl,
        spListTitle: this.properties.spListTitle || 'Calendar',
        teamMembersListTitle: this.properties.teamMembersListTitle || 'Team Members',
        teamMembersTitle: this.properties.teamMembersTitle || 'Our Team',
        teamMembersItemsPerRow: this.properties.teamMembersItemsPerRow || 4,
        showPageBanner: this.properties.showPageBanner !== false,
        showCalendar: this.properties.showCalendar !== false,
        showTeamMembers: this.properties.showTeamMembers !== false,
        quickLinksListTitle: this.properties.quickLinksListTitle || 'QuickLinks',
        quickLinksTitle: this.properties.quickLinksTitle || 'Quick Links',
        showQuickLinks: this.properties.showQuickLinks !== false
      }
    );

    ReactDom.render(element, this.domElement);
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
              groupName: 'Banner Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('bannerLayout', {
                  label: 'Banner Layout',
                  options: bannerLayoutOptions,
                  selectedKey: this.properties.bannerLayout || 'carousel'
                }),
                PropertyPaneTextField('bannerListTitle', {
                  label: 'Banner List Title',
                  value: this.properties.bannerListTitle || 'BannerImages',
                  description: 'The name of the SharePoint list containing banner images'
                })
              ]
            },
            {
              groupName: 'Calendar Settings',
              groupFields: [
                PropertyPaneDropdown('calendarLayout', {
                  label: 'Calendar Layout',
                  options: calendarLayoutOptions,
                  selectedKey: this.properties.calendarLayout || 'grid'
                }),
                PropertyPaneDropdown('calendarSource', {
                  label: 'Calendar Source',
                  options: calendarSourceOptions,
                  selectedKey: this.properties.calendarSource || 'outlook'
                }),
                PropertyPaneTextField('groupId', {
                  label: 'Outlook Group Calendar ID',
                  value: this.properties.groupId || '',
                  description: 'The ID of the Microsoft 365 Group for calendar events'
                }),
                PropertyPaneTextField('spListTitle', {
                  label: 'SharePoint Calendar List Title',
                  value: this.properties.spListTitle || 'Calendar',
                  description: 'The name of the SharePoint list containing calendar events'
                })
              ]
            },
            {
              groupName: 'Team Members Settings',
              groupFields: [
                PropertyPaneTextField('teamMembersListTitle', {
                  label: 'Team Members List Title',
                  value: this.properties.teamMembersListTitle || 'Team Members',
                  description: 'The name of the SharePoint list containing team members'
                }),
                PropertyPaneTextField('teamMembersTitle', {
                  label: 'Team Members Section Title',
                  value: this.properties.teamMembersTitle || 'Our Team',
                  description: 'The title displayed above the team members cards'
                })
              ]
            },
            {
              groupName: 'Quick Links Settings',
              groupFields: [
                PropertyPaneTextField('quickLinksListTitle', {
                  label: 'Quick Links List Title',
                  value: this.properties.quickLinksListTitle || 'QuickLinks',
                  description: 'The name of the SharePoint list containing quick links'
                }),
                PropertyPaneTextField('quickLinksTitle', {
                  label: 'Quick Links Section Title',
                  value: this.properties.quickLinksTitle || 'Quick Links',
                  description: 'The title displayed above the quick links carousel'
                })
              ]
            },
            {
              groupName: 'Visibility Settings',
              groupFields: [
                PropertyPaneToggle('showPageBanner', {
                  label: 'Show Page Banner',
                  checked: this.properties.showPageBanner !== false
                }),
                PropertyPaneToggle('showQuickLinks', {
                  label: 'Show Quick Links',
                  checked: this.properties.showQuickLinks !== false
                }),
                PropertyPaneToggle('showCalendar', {
                  label: 'Show Calendar',
                  checked: this.properties.showCalendar !== false
                }),
                PropertyPaneToggle('showTeamMembers', {
                  label: 'Show Team Members',
                  checked: this.properties.showTeamMembers !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
