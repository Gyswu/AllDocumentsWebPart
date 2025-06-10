import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import AllDocuments from './components/AllDocuments';
export default class AllDocumentsWebPart extends BaseClientSideWebPart {
    render() {
        const element = React.createElement(AllDocuments, {
            description: this.properties.description,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            environmentMessage: '',
            hasTeamsContext: false,
            userDisplayName: this.context.pageContext.user.displayName,
            isDarkTheme: false,
            customColumns: ['TestColumn'] // 👈 Incluye tus columnas personalizadas aquí
        });
        ReactDom.render(element, this.domElement);
    }
    onDispose() {
        ReactDom.unmountComponentAtNode(this.domElement);
    }
    get dataVersion() {
        return Version.parse('1.0');
    }
    getPropertyPaneConfiguration() {
        return {
            pages: []
        };
    }
}
//# sourceMappingURL=AllDocumentsWebPart.js.map