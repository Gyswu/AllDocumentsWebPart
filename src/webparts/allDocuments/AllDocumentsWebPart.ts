// AllDocumentsWebPart.ts
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { getSP } from "./pnpjsConfig"; // lo haremos justo ahora
import AllDocuments, { IAllDocumentsProps } from './components/AllDocuments';

export interface IAllDocumentsWebPartProps { /* sin propiedades en este ejemplo */ }

export default class AllDocumentsWebPart extends BaseClientSideWebPart<IAllDocumentsWebPartProps> {

  public async onInit(): Promise<void> {
    await super.onInit();
    // Configurar PnPjs con el contexto SPFx actual
    getSP(this.context); // inicializa PnP
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IAllDocumentsProps> = React.createElement(AllDocuments, {
      // Se podría pasar información adicional vía props si fuera necesario
    });
    ReactDom.render(element, this.domElement);
  }
}
