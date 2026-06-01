import { IExecutive } from './IExecutive';

export interface ILeadershipMessageAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface ILeadershipMessageAuthor {
  Id: number;
  Title: string;
}

export interface ILeadershipMessage {
  Id: number;
  Title: string;
  Body?: string;
  Executive?: IExecutive;
  Author?: ILeadershipMessageAuthor;
  Created?: Date;
  Modified?: Date;
  Attachments?: ILeadershipMessageAttachment[];
}
