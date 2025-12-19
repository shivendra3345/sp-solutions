import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBannerProps {
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  /** Toggle whether the welcome panel for new joiners is shown on the page */
  showWelcome?: boolean;
  context: WebPartContext; // SPFx web part context
}
