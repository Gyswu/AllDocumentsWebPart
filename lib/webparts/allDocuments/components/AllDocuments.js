var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import styles from "./AllDocuments.module.scss";
import { initializeFileTypeIcons, getFileTypeIconProps, } from "@fluentui/react-file-type-icons";
import { Icon } from "@fluentui/react/lib/Icon";
import { DetailsList, DetailsListLayoutMode, SelectionMode, } from "@fluentui/react/lib/DetailsList";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { Dropdown } from "@fluentui/react/lib/Dropdown";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Link } from "@fluentui/react/lib/Link";
import { mergeStyles } from "@fluentui/react/lib/Styling";
// Initialize the icons once
initializeFileTypeIcons();
const iconClass = mergeStyles({
    marginRight: 8,
    verticalAlign: "middle",
});
export default class AllDocuments extends React.Component {
    constructor(props) {
        super(props);
        this._onColumnClick = (ev, column) => {
            const { columns, items } = this.state;
            const newColumns = columns.slice();
            const currColumn = newColumns.filter((currCol) => column.key === currCol.key)[0];
            newColumns.forEach((newCol) => {
                if (newCol === currColumn) {
                    currColumn.isSortedDescending = !currColumn.isSortedDescending;
                    currColumn.isSorted = true;
                }
                else {
                    newCol.isSorted = false;
                    newCol.isSortedDescending = true;
                }
            });
            const newItems = this._copyAndSort(items, currColumn.fieldName, currColumn.isSortedDescending);
            this.setState({
                columns: newColumns,
                items: newItems,
            });
        };
        this._onSearchChange = (event, newValue) => {
            this.setState({ searchTerm: newValue || "" });
        };
        this._onFilterChanged = (column, option) => {
            const newFilters = Object.assign(Object.assign({}, this.state.filters), { [column]: (option === null || option === void 0 ? void 0 : option.key) || "" });
            this.setState({ filters: newFilters });
        };
        this.state = {
            items: [],
            loading: true,
            filters: {},
            filterOptions: {},
            searchTerm: "",
            columns: this._buildColumns(),
        };
    }
    componentDidMount() {
        var _a, _b, _c, _d;
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const res = yield this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Id,Title`, SPHttpClient.configurations.v1);
                const json = yield res.json();
                const libraries = json.value;
                const allItems = [];
                const filterOptions = {};
                for (const lib of libraries) {
                    const camlQuery = {
                        ViewXml: `
            <View Scope='RecursiveAll'>
              <Query></Query>
              <ViewFields>
                <FieldRef Name='FileLeafRef' />
                <FieldRef Name='FileRef' />
                <FieldRef Name='Modified' />
                <FieldRef Name='Editor' />
                ${this.props.customColumns
                            .map((col) => `<FieldRef Name='${col.internalName}' />`)
                            .join("")}
              </ViewFields>
            </View>`,
                    };
                    const itemsRes = yield this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists(guid'${lib.Id}')/RenderListDataAsStream`, SPHttpClient.configurations.v1, {
                        headers: {
                            Accept: "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=verbose",
                        },
                        body: JSON.stringify({ parameters: camlQuery }),
                    });
                    const itemsJson = yield itemsRes.json();
                    const rows = itemsJson === null || itemsJson === void 0 ? void 0 : itemsJson.Row;
                    if (!rows || rows.length === 0) {
                        console.warn(`No items found in library: ${lib.Title}`);
                        continue;
                    }
                    for (const file of rows) {
                        if (file.FSObjType === "1" || file.FSObjType === 1)
                            continue;
                        const fileName = file.FileLeafRef;
                        const filePath = file.FileRef;
                        const modified = file.Modified;
                        const editor = ((_b = (_a = file.Editor) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b.title) || ((_c = file.Editor) === null || _c === void 0 ? void 0 : _c.title) || file.Editor || "";
                        const serverRelativePath = filePath;
                        const parentPath = serverRelativePath.substring(0, serverRelativePath.lastIndexOf("/"));
                        const libPath = `${parentPath}/Forms/AllItems.aspx`;
                        const customData = {};
                        for (const col of this.props.customColumns) {
                            const value = file[col.internalName];
                            customData[col.internalName] = value || "";
                            if (!filterOptions[col.internalName])
                                filterOptions[col.internalName] = new Set();
                            if (value)
                                filterOptions[col.internalName].add(value);
                        }
                        const extension = (_d = fileName.split(".").pop()) === null || _d === void 0 ? void 0 : _d.toLowerCase();
                        const officeExtensions = [
                            "docx",
                            "xlsx",
                            "pptx",
                            "doc",
                            "xls",
                            "ppt",
                        ];
                        let editUrl;
                        if (officeExtensions.includes(extension || "")) {
                            editUrl = `${this.props.siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(filePath)}&action=edit&mobileredirect=true`;
                        }
                        else if (extension === "pdf") {
                            editUrl = `${window.location.origin}${libPath}?id=${encodeURIComponent(serverRelativePath)}&parent=${encodeURIComponent(parentPath)}`;
                        }
                        else {
                            editUrl = `${this.props.siteUrl}/${filePath}`;
                        }
                        allItems.push({
                            name: fileName,
                            modified: modified,
                            modifiedBy: editor,
                            library: lib.Title,
                            editUrl: editUrl,
                            customColumns: customData,
                        });
                    }
                }
                this.setState({
                    items: allItems,
                    loading: false,
                    filters: {},
                    filterOptions: filterOptions,
                });
                // Debug Purpose
                //console.log("Loaded document items:", allItems);
            }
            catch (err) {
                console.error("Error loading documents:", err);
                this.setState({ loading: false });
            }
        });
    }
    _buildColumns() {
        const columns = [
            {
                key: "name",
                name: "Name",
                fieldName: "name",
                minWidth: 250,
                maxWidth: 400,
                isResizable: true,
                isSorted: false,
                isSortedDescending: false,
                onColumnClick: this._onColumnClick,
                onRender: (item) => {
                    var _a;
                    const extension = ((_a = item.name.split(".").pop()) === null || _a === void 0 ? void 0 : _a.toLowerCase()) || "unknown";
                    return (React.createElement(Stack, { horizontal: true, verticalAlign: "center" },
                        React.createElement(Icon, Object.assign({}, getFileTypeIconProps({
                            extension: extension,
                            size: 20,
                            imageFileType: "svg",
                        }), { className: iconClass })),
                        React.createElement(Link, { href: item.editUrl, target: "_blank", styles: { root: { whiteSpace: "normal", overflow: "visible" } } }, item.name)));
                },
            },
        ];
        // Add custom columns
        this.props.customColumns.forEach((col) => {
            columns.push({
                key: col.internalName,
                name: col.label,
                fieldName: col.internalName,
                minWidth: 100,
                maxWidth: 200,
                isResizable: true,
                isSorted: false,
                isSortedDescending: false,
                onColumnClick: this._onColumnClick,
                onRender: (item) => {
                    return React.createElement(Text, null, item.customColumns[col.internalName] || "");
                },
            });
        });
        return columns;
    }
    _copyAndSort(items, columnKey, isSortedDescending) {
        return items.slice(0).sort((a, b) => {
            let aValue;
            let bValue;
            if (columnKey === "name" ||
                columnKey === "modified" ||
                columnKey === "modifiedBy" ||
                columnKey === "library") {
                aValue = a[columnKey];
                bValue = b[columnKey];
            }
            else {
                // Custom column
                aValue = a.customColumns[columnKey] || "";
                bValue = b.customColumns[columnKey] || "";
            }
            if (isSortedDescending) {
                return aValue.localeCompare(bValue) * -1;
            }
            else {
                return aValue.localeCompare(bValue);
            }
        });
    }
    _getFilteredItems() {
        const { items, filters, searchTerm } = this.state;
        let filtered = items.filter((item) => Object.entries(filters).every(([key, val]) => val === "" || item.customColumns[key] === val));
        if (searchTerm) {
            filtered = filtered.filter((item) => item.name.toLowerCase().includes(searchTerm.toLowerCase()));
        }
        return filtered;
    }
    render() {
        const { loading, filterOptions, filters, searchTerm } = this.state;
        const allowedSites = ["sites/sp-FIN"];
        const authorized = allowedSites.some((path) => this.props.siteUrl.includes(path));
        if (!authorized) {
            return (React.createElement(MessageBar, { messageBarType: MessageBarType.error }, "\u26A0\uFE0F This webpart is not authorized to be loaded in this site."));
        }
        if (loading || this.props.customColumns.length === 0) {
            return (React.createElement(Stack, { horizontalAlign: "center", tokens: { padding: 20 } },
                React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading Files..." })));
        }
        const filteredItems = this._getFilteredItems();
        return (React.createElement(Stack, { tokens: { childrenGap: 16 }, className: styles.container },
            React.createElement(SearchBox, { placeholder: "Search files...", value: searchTerm, onChange: this._onSearchChange, underlined: true }),
            React.createElement(Stack, { horizontal: true, wrap: true, tokens: { childrenGap: 16 }, styles: { root: { marginBottom: 8 } } }, this.props.customColumns.map((col) => {
                const options = [
                    { key: "", text: "All" },
                    ...[...(filterOptions[col.internalName] || [])].map((val) => ({
                        key: val,
                        text: val,
                    })),
                ];
                return (React.createElement(Stack.Item, { key: col.internalName, styles: { root: { minWidth: 200 } } },
                    React.createElement(Dropdown, { label: col.label, selectedKey: filters[col.internalName] || "", onChange: (e, option) => this._onFilterChanged(col.internalName, option), options: options, styles: { dropdown: { width: 200 } } })));
            })),
            React.createElement(DetailsList, { items: filteredItems, columns: this.state.columns, setKey: "set", layoutMode: DetailsListLayoutMode.justified, selectionMode: SelectionMode.none, isHeaderVisible: true, compact: false })));
    }
}
//# sourceMappingURL=AllDocuments.js.map