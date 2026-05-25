export interface IAlertAuthor {
  Id: number;
  Title: string;
}

export interface IAlertAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IAlertAudience {
  Id: number;
  Title: string;
}

export interface IAlert {
  Id: number;
  Title: string;
  Body?: string;
  AlertType: string;
  PublishDate: Date;
  ExpiryDate?: Date;
  Priority: "Critical" | "High" | "Medium" | "Low";
  Status: "Draft" | "Scheduled" | "Active" | "Expired";
  TargetAudienceType?: "All" | "Specific" | "Except";
  TargetAudience?: IAlertAudience[];
  BannerImageUrl?: string;
  Attachments?: IAlertAttachment[];
  Author?: IAlertAuthor;
  Created?: Date;
}
