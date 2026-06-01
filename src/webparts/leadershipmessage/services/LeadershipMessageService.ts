import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import { ILeadershipMessage, ILeadershipMessageAttachment } from "../models/ILeadershipMessage";
import { IExecutive } from "../models/IExecutive";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IMessageListItem {
  Id: number;
  Title: string;
  Body?: string;
  // Only safe text fields — PhotoUrl is a Hyperlink column and cannot be
  // expanded through a lookup; photo data is merged in the panel from getExecutives()
  Executive?: { Id: number; Title: string; Position: string; };
  Author?: { Id: number; Title: string };
  Created: string;
  Modified: string;
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
}

export class LeadershipMessageService {
  private sp: SPFI;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.listName = listName;
    this.sp = spfi().using(SPFx(context));
  }

  public async getMessages(): Promise<ILeadershipMessage[]> {
    const items: IMessageListItem[] = await this.sp.web.lists
      .getByTitle(this.listName)
      .items
      .select(
        "Id", "Title", "Body",
        "Executive/Id", "Executive/Title", "Executive/Position",
        "Author/Id", "Author/Title",
        "Created", "Modified", "AttachmentFiles"
      )
      .expand("Executive", "Author", "AttachmentFiles")
      .orderBy("Created", false)();

    return items.map(i => this._mapItem(i));
  }

  public async addMessage(data: Partial<ILeadershipMessage>, attachments?: File[]): Promise<void> {
    const result = await this.sp.web.lists
      .getByTitle(this.listName)
      .items.add(this._buildItemData(data));
    if (attachments && attachments.length > 0) {
      await this._uploadAttachments(result.Id, attachments);
    }
  }

  public async updateMessage(
    id: number,
    data: Partial<ILeadershipMessage>,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> {
    await this.sp.web.lists.getByTitle(this.listName).items.getById(id).update(this._buildItemData(data));
    if (deletedAttachmentNames && deletedAttachmentNames.length > 0) {
      await this._deleteAttachments(id, deletedAttachmentNames);
    }
    if (attachments && attachments.length > 0) {
      await this._uploadAttachments(id, attachments);
    }
  }

  private _buildItemData(data: Partial<ILeadershipMessage>): Record<string, unknown> {
    const item: Record<string, unknown> = {
      Title: data.Title ?? "",
      Body: data.Body ?? "",
    };
    if (data.Executive) {
      item.ExecutiveId = data.Executive.Id;
    }
    return item;
  }

  private _mapItem(i: IMessageListItem): ILeadershipMessage {
    // Photo / focal-point fields are NOT selected here because PhotoUrl is a
    // Hyperlink column that cannot be expanded through a lookup in SP REST.
    // LeadershipMessagePanel merges full executive data from getExecutives() after load.
    const executive: IExecutive | undefined = i.Executive
      ? {
          Id: i.Executive.Id,
          Title: i.Executive.Title ?? "",
          Position: i.Executive.Position ?? "",
          PhotoUrl: "",
          PhotoFocalX: 50,
          PhotoFocalY: 50,
        }
      : undefined;

    return {
      Id: i.Id,
      Title: i.Title,
      Body: i.Body,
      Executive: executive,
      Author: i.Author ? { Id: i.Author.Id, Title: i.Author.Title } : undefined,
      Created: new Date(i.Created),
      Modified: new Date(i.Modified),
      Attachments: (i.AttachmentFiles ?? []).map(f => ({
        FileName: f.FileName,
        ServerRelativeUrl: f.ServerRelativeUrl,
      } as ILeadershipMessageAttachment)),
    };
  }

  private async _uploadAttachments(itemId: number, files: File[]): Promise<void> {
    for (const file of files) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.add(file.name, file);
      } catch (err) {
        console.error("Failed to upload attachment " + file.name, err);
      }
    }
  }

  private async _deleteAttachments(itemId: number, fileNames: string[]): Promise<void> {
    for (const fileName of fileNames) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.getByName(fileName)
          .delete();
      } catch (err) {
        console.error("Failed to delete attachment " + fileName, err);
      }
    }
  }
}
