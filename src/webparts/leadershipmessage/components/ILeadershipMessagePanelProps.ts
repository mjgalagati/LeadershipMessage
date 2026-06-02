import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILeadershipMessagePanelProps {
  sectionTitle: string;
  sourceList: string;
  executiveList: string;
  context: WebPartContext;
  currentUserId: number;
  isEditor: boolean;
}
