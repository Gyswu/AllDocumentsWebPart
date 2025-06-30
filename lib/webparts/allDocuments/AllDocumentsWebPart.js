import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import AllDocuments from './components/AllDocuments';
export default class AllDocumentsWebPart extends BaseClientSideWebPart {
    render() {
        const element = React.createElement(AllDocuments, {
            siteUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            customColumns: this._parseCustomColumns(),
            description: this.properties.description,
            environmentMessage: '',
            hasTeamsContext: false,
            userDisplayName: this.context.pageContext.user.displayName,
        });
        ReactDom.render(element, this.domElement);
    }
    _parseCustomColumns() {
        const raw = this.properties.customColumnsRaw || '';
        return raw
            .split(';')
            .map(entry => entry.trim())
            .filter(entry => entry.length > 0)
            .map(entry => {
            const [internalName, label] = entry.split(',').map(x => x.trim());
            return {
                internalName,
                label: label || internalName
            };
        });
    }
    onDispose() {
        ReactDom.unmountComponentAtNode(this.domElement);
    }
    get dataVersion() {
        return Version.parse('1.0');
    }
    getPropertyPaneConfiguration() {
        return {
            pages: [
                {
                    header: { description: "Configuración del Web Part" },
                    groups: [
                        {
                            groupName: "Columnas personalizadas",
                            groupFields: [
                                PropertyPaneTextField('customColumnsRaw', {
                                    label: 'Columnas personalizadas (formato: InternalName,Label;...)',
                                    multiline: true
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
//# sourceMappingURL=AllDocumentsWebPart.js.map