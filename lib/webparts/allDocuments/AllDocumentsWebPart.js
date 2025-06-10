var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// AllDocumentsWebPart.ts
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { getSP } from "./pnpjsConfig"; // lo haremos justo ahora
import AllDocuments from './components/AllDocuments';
export default class AllDocumentsWebPart extends BaseClientSideWebPart {
    onInit() {
        const _super = Object.create(null, {
            onInit: { get: () => super.onInit }
        });
        return __awaiter(this, void 0, void 0, function* () {
            yield _super.onInit.call(this);
            // Configurar PnPjs con el contexto SPFx actual
            getSP(this.context); // inicializa PnP
            return Promise.resolve();
        });
    }
    render() {
        const element = React.createElement(AllDocuments, {});
        ReactDom.render(element, this.domElement);
    }
}
//# sourceMappingURL=AllDocumentsWebPart.js.map