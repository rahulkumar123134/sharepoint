import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./SharepointWebPart.module.scss";
import * as strings from "SharepointWebPartStrings";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ISharepointWebPartProps {
  description: string;
  active: boolean;
  title: string;
}

export interface FileRecords {
  value: File[];
}

export interface File {
  Id: string;
  Title: string;
}

export default class SharepointWebPart extends BaseClientSideWebPart<ISharepointWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    this.domElement.innerHTML = `
  <div class="${styles.welcome}">
    <img alt="" src="${
      this._isDarkTheme
        ? require("./assets/welcome-dark.png")
        : require("./assets/welcome-light.png")
    }" class="${styles.welcomeImage}" />
    <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
    <div>${this._environmentMessage}</div>
  </div>
  <div>
    <h3>Welcome to SharePoint Framework!</h3>
    <div>Web part description: <strong>${escape(
      this.properties.description
    )}</strong></div>
    <div>Loading from: <strong>${escape(
      this.context.pageContext.web.title
    )}</strong></div>
  </div>
  <div id="spListContainer" />
</section>`;
    this._renderListAsync();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch((ex) => {
        console.log(ex);
      });
  }

  private _getListData(): Promise<FileRecords> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch(() => {});
  }

  private _renderList(files: File[]): void {
    let html: string = "";
    files.forEach((item: File) => {
      console.log(item);
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    if (this.domElement.querySelector("#spListContainer") !== null) {
      this.domElement.querySelector("#spListContainer")!.innerHTML = html;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneTextField("title", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneCheckbox("active", {
                  text: "Status",
                  checked: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
