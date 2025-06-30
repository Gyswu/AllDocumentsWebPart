import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IAllDocumentsWebPartProps {
    description: string;
    customColumnsRaw: string;
}
export default class AllDocumentsWebPart extends BaseClientSideWebPart<IAllDocumentsWebPartProps> {
    render(): void;
    private _parseCustomColumns;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=AllDocumentsWebPart.d.ts.map