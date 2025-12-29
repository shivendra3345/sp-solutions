import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBannerProps {
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  /** Toggle whether the welcome panel for new joiners is shown on the page */
  showWelcome?: boolean;
  /** Toggle whether the events carousel is shown */
  showEvents?: boolean;
  /** Toggle whether the news/announcements section is shown */
  showNews?: boolean;
  /** Customizable titles shown in the UI */
  welcomeTitle?: string;
  eventsTitle?: string;
  newsTitle?: string;
  /** Comma-separated list of site URLs (relative or absolute) to exclude when fetching events */
  excludedSites?: string;
  context: WebPartContext; // SPFx web part context
}
