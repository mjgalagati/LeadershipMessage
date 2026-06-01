import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LeadershipMessageWebPartStrings';
import LeadershipMessagePanel from './components/LeadershipMessagePanel';
import { ILeadershipMessagePanelProps } from './components/ILeadershipMessagePanelProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";

export interface ILeadershipMessageWebPartProps {
  sectionTitle: string;
  sourceList: string;
  executiveList: string;
}

export default class LeadershipMessageWebPart extends BaseClientSideWebPart<ILeadershipMessageWebPartProps> {

  private _sp!: ReturnType<typeof spfi>;
  private _siteLists: { key: string; text: string }[] = [];
  private _currentUserId: number = 0;

  public render(): void {
    const element: React.ReactElement<ILeadershipMessagePanelProps> = React.createElement(
      LeadershipMessagePanel,
      {
        sectionTitle: this.properties.sectionTitle || "EXECUTIVE MESSAGE",
        sourceList: this.properties.sourceList,
        executiveList: this.properties.executiveList,
        context: this.context,
        currentUserId: this._currentUserId,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    await Promise.all([
      this._loadSiteLists(),
      this._loadCurrentUser(),
    ]);
    return super.onInit();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (newValue !== oldValue) {
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _loadCurrentUser(): Promise<void> {
    try {
      const currentUser = await this._sp.web.currentUser.select("Id")();
      this._currentUserId = currentUser.Id;
    } catch (err) {
      console.error("Failed to load current user:", err);
    }
  }

  private async _loadSiteLists(): Promise<void> {
    try {
      const lists = await this._sp.web.lists
        .filter("BaseTemplate eq 100 and Hidden eq false")
        .select("Title")();
      this._siteLists = lists.map(l => ({ key: l.Title, text: l.Title }));
    } catch (err) {
      console.error("Failed to load site lists:", err);
      this._siteLists = [];
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('sectionTitle', {
                  label: 'Section Title',
                  placeholder: 'EXECUTIVE MESSAGE',
                }),
                PropertyPaneDropdown('sourceList', {
                  label: 'Message Source List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneDropdown('executiveList', {
                  label: 'Executive Directory List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
