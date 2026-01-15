import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDptPageProps {
  description: string;
  context: WebPartContext;
  bannerLayout: 'carousel' | 'tile';
  bannerListTitle: string;
  calendarLayout: 'grid' | 'carousel';
  calendarSource: 'outlook' | 'sharepoint' | 'both';
  groupId?: string;
  spSiteUrl?: string;
  spListTitle?: string;
  teamMembersListTitle?: string;
  teamMembersTitle?: string;
  teamMembersDepartment?: string;
  teamMembersItemsPerRow?: number;
  showPageBanner: boolean;
  showCalendar: boolean;
  showTeamMembers: boolean;
  quickLinksListTitle?: string;
  quickLinksTitle?: string;
  quickLinksDepartment?: string;
  showQuickLinks: boolean;
}
