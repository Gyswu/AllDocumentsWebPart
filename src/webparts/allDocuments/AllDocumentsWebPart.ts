import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AllDocuments from './components/AllDocuments';
import { IAllDocumentsProps } from './components/IAllDocumentsProps';

export interface IAllDocumentsWebPartProps {
  description: string;
}

export default class AllDocumentsWebPart extends BaseClientSideWebPart<IAllDocumentsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAllDocumentsProps> = React.createElement(
      AllDocuments,
      {
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        environmentMessage: '',
        hasTeamsContext: false,
        userDisplayName: this.context.pageContext.user.displayName,
        isDarkTheme: false,
        customColumns: ['TestColumn'] // 👈 Incluye tus columnas personalizadas aquí
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
