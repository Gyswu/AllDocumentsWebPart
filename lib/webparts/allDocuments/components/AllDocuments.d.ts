import * as React from 'react';
export interface IDocumentItem {
    name: string;
    modified: Date;
    modifiedBy: string;
    url: string;
    library: string;
    [key: string]: any;
}
interface ICustomField {
    title: string;
    internalName: string;
}
interface IState {
    items: IDocumentItem[];
    customFields: ICustomField[];
    loading: boolean;
}
export default class AllDocuments extends React.Component<{}, IState> {
    private sp;
    constructor(props: {});
    componentDidMount(): void;
    private loadAllDocuments;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=AllDocuments.d.ts.map