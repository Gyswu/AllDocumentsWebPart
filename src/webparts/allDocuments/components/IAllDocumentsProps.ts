import { SPHttpClient } from "@microsoft/sp-http";

export interface IAllDocumentsProps {
  description: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  isDarkTheme?: boolean;
  customColumns: { internalName: string; label: string }[];
  useColumnFormatting: boolean;
}
