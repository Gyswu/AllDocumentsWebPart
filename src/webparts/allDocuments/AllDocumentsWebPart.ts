import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import AllDocuments from "./components/AllDocuments";
import { IAllDocumentsProps } from "./components/IAllDocumentsProps";

export interface IAllDocumentsWebPartProps {
  description: string;
  customColumnsRaw: string;
  useColumnFormatting: boolean;
}

export default class AllDocumentsWebPart extends BaseClientSideWebPart<IAllDocumentsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAllDocumentsProps> = React.createElement(
      AllDocuments,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        customColumns: this._parseCustomColumns(),
        description: this.properties.description,
        environmentMessage: "",
        hasTeamsContext: false,
        userDisplayName: this.context.pageContext.user.displayName,
        useColumnFormatting: this.properties.useColumnFormatting || false,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _parseCustomColumns(): { internalName: string; label: string }[] {
    const raw = this.properties.customColumnsRaw || "";
    return raw
      .split(";")
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        const [internalName, label] = entry.split(",").map((x) => x.trim());
        return {
          internalName,
          label: label || internalName,
        };
      });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Web Part configuration" },
          groups: [
            {
              groupName: "Custom columns",
              groupFields: [
                PropertyPaneTextField("customColumnsRaw", {
                  label:
                    "Custom columns (format: InternalName,Label;...)",
                  multiline: true,
                  description:
                    "Ex: Testeo,Prueba;Status,Estado;Priority,Prioridad",
                }),
              ],
            },
            {
              groupName: "Visualization options",
              groupFields: [
                PropertyPaneToggle("useColumnFormatting", {
                  label: "Use Sharepoint column formating (Custom column colors)",
                  onText: "Enabled",
                  offText: "Disabled",
                  checked: this.properties.useColumnFormatting || false,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
