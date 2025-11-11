import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneTextField, PropertyPaneToggle, } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import AllDocuments from "./components/AllDocuments";
export default class AllDocumentsWebPart extends BaseClientSideWebPart {
    render() {
        const element = React.createElement(AllDocuments, {
            siteUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            customColumns: this._parseCustomColumns(),
            description: this.properties.description,
            environmentMessage: "",
            hasTeamsContext: false,
            userDisplayName: this.context.pageContext.user.displayName,
            useColumnFormatting: this.properties.useColumnFormatting || false,
            showModified: this.properties.showModified !== false,
            showModifiedBy: this.properties.showModifiedBy !== false,
            showLibrary: this.properties.showLibrary !== false,
        });
        ReactDom.render(element, this.domElement);
    }
    _parseCustomColumns() {
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
    onDispose() {
        ReactDom.unmountComponentAtNode(this.domElement);
    }
    get dataVersion() {
        return Version.parse("1.0");
    }
    getPropertyPaneConfiguration() {
        return {
            pages: [
                {
                    header: { description: "Web Part configuration" },
                    groups: [
                        {
                            groupName: "System Columns configuration",
                            groupFields: [
                                PropertyPaneToggle("showModified", {
                                    label: 'System Columns "Modified"',
                                    checked: this.properties.showModified !== false,
                                }),
                                PropertyPaneToggle("showModifiedBy", {
                                    label: 'System Columns "Modified By"',
                                    checked: this.properties.showModifiedBy !== false,
                                }),
                                PropertyPaneToggle("showLibrary", {
                                    label: 'System Columns "Library"',
                                    checked: this.properties.showLibrary !== false,
                                }),
                            ],
                        },
                        {
                            groupName: "Custom columns",
                            groupFields: [
                                PropertyPaneTextField("customColumnsRaw", {
                                    label: "Custom columns (format: InternalName,Label;...)",
                                    multiline: true,
                                    description: "Ex: Testeo,Prueba;Status,Estado;Priority,Prioridad",
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
//# sourceMappingURL=AllDocumentsWebPart.js.map