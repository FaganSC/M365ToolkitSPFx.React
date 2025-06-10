import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITestWebPartProps {
  context: WebPartContext
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
