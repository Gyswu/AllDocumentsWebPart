import * as React from 'react';
import { IAllDocumentsProps } from './IAllDocumentsProps';
export interface IDocumentItem {
    name: string;
    modified: string;
    modifiedBy: string;
    library: string;
    editUrl: string;
    customColumns: {
        [key: string]: string;
    };
}
interface IState {
    items: IDocumentItem[];
    loading: boolean;
    filters: {
        [key: string]: string;
    };
    filterOptions: {
        [key: string]: Set<string>;
    };
}
export default class AllDocuments extends React.Component<IAllDocumentsProps, IState> {
    constructor(props: IAllDocumentsProps);
    componentDidMount(): Promise<void>;
    private onFilterChanged;
    private applyFilters;
    render(): React.ReactElement<IAllDocumentsProps>;
}
export {};
//# sourceMappingURL=AllDocuments.d.ts.map