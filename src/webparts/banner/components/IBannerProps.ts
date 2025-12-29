import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBannerProps {
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  /** Toggle whether the welcome panel for new joiners is shown on the page */
  showWelcome?: boolean;
  /** Comma-separated list of site URLs (relative or absolute) to exclude when fetching events */
  excludedSites?: string;
  context: WebPartContext; // SPFx web part context
}
