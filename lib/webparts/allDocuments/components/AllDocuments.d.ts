import * as React from "react";
import { IAllDocumentsProps } from "./IAllDocumentsProps";
import { IColumn } from "@fluentui/react/lib/DetailsList";
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
    searchTerm: string;
    columns: IColumn[];
}
export default class AllDocuments extends React.Component<IAllDocumentsProps, IState> {
    constructor(props: IAllDocumentsProps);
    componentDidMount(): Promise<void>;
    private _buildColumns;
    private _onColumnClick;
    private _copyAndSort;
    private _onSearchChange;
    private _onFilterChanged;
    private _getFilteredItems;
    render(): React.ReactElement<IAllDocumentsProps>;
}
export {};
//# sourceMappingURL=AllDocuments.d.ts.map