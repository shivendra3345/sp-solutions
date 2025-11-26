import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBannerProps {
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  context: WebPartContext; // SPFx web part context
}
