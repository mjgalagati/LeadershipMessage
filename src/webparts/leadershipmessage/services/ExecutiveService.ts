import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IExecutive } from "../models/IExecutive";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const PHOTO_FOLDER = "SiteAssets/ExecutivePhotos";

// PhotoUrl is a Hyperlink/Picture column — SP REST returns and expects an object
interface ISpUrlValue {
  Url: string;
  Description: string;
}

interface IExecutiveListItem {
  Id: number;
  Title: string;
  Position: string;
  PhotoUrl?: ISpUrlValue;   // Hyperlink field — comes back as { Url, Description }
  PhotoFocalX: number;
  PhotoFocalY: number;
  Bio?: string;
}

export class ExecutiveService {
  private sp: SPFI;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.listName = listName;
    this.sp = spfi().using(SPFx(context));
  }

  public async getExecutives(): Promise<IExecutive[]> {
    const items: IExecutiveListItem[] = await this.sp.web.lists
      .getByTitle(this.listName)
      .items
      .select("Id", "Title", "Position", "PhotoUrl", "PhotoFocalX", "PhotoFocalY", "Bio")
      .orderBy("Title", true)();

    return items.map(i => ({
      Id: i.Id,
      Title: i.Title ?? "",
      Position: i.Position ?? "",
      PhotoUrl: i.PhotoUrl?.Url ?? "",   // extract the Url string from the hyperlink object
      PhotoFocalX: i.PhotoFocalX ?? 50,
      PhotoFocalY: i.PhotoFocalY ?? 50,
      Bio: i.Bio,
    }));
  }

  public async addExecutive(data: Partial<IExecutive>, photoFile?: File): Promise<void> {
    let photoUrl = data.PhotoUrl ?? "";
    if (photoFile) {
      photoUrl = await this.uploadPhoto(photoFile);
    }
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.add(
        this._buildPayload(data, photoUrl)
      );
    } catch (err) {
      throw new Error(this._extractSpError(err));
    }
  }

  public async updateExecutive(id: number, data: Partial<IExecutive>, photoFile?: File): Promise<void> {
    let photoUrl = data.PhotoUrl ?? "";
    if (photoFile) {
      photoUrl = await this.uploadPhoto(photoFile);
    }
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.getById(id).update(
        this._buildPayload(data, photoUrl)
      );
    } catch (err) {
      throw new Error(this._extractSpError(err));
    }
  }

  private _buildPayload(data: Partial<IExecutive>, photoUrl: string): Record<string, unknown> {
    const payload: Record<string, unknown> = {
      Title: data.Title ?? "",
    };
    if (data.Position !== undefined) payload.Position = data.Position;
    // Hyperlink field requires an object — plain string causes 400
    if (photoUrl) {
      payload.PhotoUrl = { Url: photoUrl, Description: photoUrl };
    }
    if (data.PhotoFocalX !== undefined) payload.PhotoFocalX = data.PhotoFocalX;
    if (data.PhotoFocalY !== undefined) payload.PhotoFocalY = data.PhotoFocalY;
    if (data.Bio) payload.Bio = data.Bio;
    return payload;
  }

  private _extractSpError(err: unknown): string {
    if (err && typeof err === "object") {
      const e = err as Record<string, unknown>;
      const body = e.data as Record<string, unknown> | undefined;
      if (body) {
        const rb = body.responseBody as Record<string, unknown> | undefined;
        const odataErr = rb?.["odata.error"] as Record<string, unknown> | undefined;
        const msgObj = odataErr?.message as Record<string, unknown> | undefined;
        if (msgObj?.value) return String(msgObj.value);
      }
      if (typeof e.message === "string") return e.message;
    }
    return "Unknown error saving executive.";
  }

  public async uploadPhoto(file: File): Promise<string> {
    const fileName = Date.now().toString() + "_" + file.name;
    try {
      try {
        await this.sp.web.getFolderByServerRelativePath(PHOTO_FOLDER).select("Exists")();
      } catch {
        await this.sp.web.folders.addUsingPath(PHOTO_FOLDER);
      }
      const result = await this.sp.web
        .getFolderByServerRelativePath(PHOTO_FOLDER)
        .files.addUsingPath(fileName, file, { Overwrite: true });
      return result.ServerRelativeUrl;
    } catch (err) {
      console.error("Photo upload error:", err);
      throw new Error("Failed to upload executive photo");
    }
  }
}
