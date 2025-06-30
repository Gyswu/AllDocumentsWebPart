var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './AllDocuments.module.scss';
export default class AllDocuments extends React.Component {
    constructor(props) {
        super(props);
        this.onSearchChange = (e) => {
            this.setState({ searchTerm: e.target.value });
        };
        this.onSortColumn = (column) => {
            const { sortConfig } = this.state;
            let direction = 'asc';
            if (sortConfig && sortConfig.column === column && sortConfig.direction === 'asc') {
                direction = 'desc';
            }
            this.setState({ sortConfig: { column, direction } });
        };
        this.renderSortableHeader = (column, label) => {
            const { sortConfig } = this.state;
            const isSorted = (sortConfig === null || sortConfig === void 0 ? void 0 : sortConfig.column) === column;
            const direction = isSorted ? (sortConfig.direction === 'asc' ? '▲' : '▼') : '';
            return (React.createElement("th", { onClick: () => this.onSortColumn(column), className: styles.sortableHeader },
                label,
                " ",
                direction));
        };
        this.onFilterChanged = (column, e) => {
            const newFilters = Object.assign(Object.assign({}, this.state.filters), { [column]: e.target.value });
            this.setState({ filters: newFilters });
        };
        this.state = {
            items: [],
            loading: true,
            filters: {},
            filterOptions: {},
            searchTerm: '',
            sortConfig: null
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
                ${this.props.customColumns.map(col => `<FieldRef Name='${col.internalName}' />`).join('')}
              </ViewFields>
            </View>`
                    };
                    const itemsRes = yield this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists(guid'${lib.Id}')/RenderListDataAsStream`, SPHttpClient.configurations.v1, {
                        headers: {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-Type': 'application/json;odata=verbose'
                        },
                        body: JSON.stringify({ parameters: camlQuery })
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
                        const editor = ((_b = (_a = file.Editor) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b.title) || ((_c = file.Editor) === null || _c === void 0 ? void 0 : _c.title) || file.Editor || '';
                        const serverRelativePath = filePath; // e.g., "/sites/sp-FINA/Fundamentals/71_Oleksandr-Lugovskyy_2024-06-19.pdf"
                        const parentPath = serverRelativePath.substring(0, serverRelativePath.lastIndexOf("/"));
                        const libPath = `${parentPath}/Forms/AllItems.aspx`;
                        const customData = {};
                        for (const col of this.props.customColumns) {
                            const value = file[col.internalName];
                            customData[col.internalName] = value || '';
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
                        if (officeExtensions.includes(extension || '')) {
                            editUrl = `${this.props.siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(filePath)}&action=edit&mobileredirect=true`;
                        }
                        else if (extension === 'pdf') {
                            // Directly open PDF in browser
                            editUrl =
                                editUrl = `${window.location.origin}${libPath}?id=${encodeURIComponent(serverRelativePath)}&parent=${encodeURIComponent(parentPath)}`;
                        }
                        else {
                            // Fallback for other files
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
                    filterOptions: filterOptions
                });
                console.log("Loaded document items:", allItems);
            }
            catch (err) {
                console.error("Error loading documents:", err);
                this.setState({ loading: false });
            }
        });
    }
    applyFilters(items) {
        const { filters, searchTerm, sortConfig } = this.state;
        let filtered = items.filter(item => Object.entries(filters).every(([key, val]) => val === '' || item.customColumns[key] === val));
        if (searchTerm) {
            filtered = filtered.filter(item => item.name.toLowerCase().includes(searchTerm.toLowerCase()));
        }
        if (sortConfig) {
            const { column, direction } = sortConfig;
            filtered.sort((a, b) => {
                const valA = column === 'name' || column === 'modified' || column === 'modifiedBy' || column === 'library'
                    ? a[column]
                    : a.customColumns[column] || '';
                const valB = column === 'name' || column === 'modified' || column === 'modifiedBy' || column === 'library'
                    ? b[column]
                    : b.customColumns[column] || '';
                return direction === 'asc'
                    ? valA.localeCompare(valB)
                    : valB.localeCompare(valA);
            });
        }
        return filtered;
    }
    render() {
        const { items, filters, filterOptions } = this.state;
        const allowedSites = [
            "sites/sp-FIN"
        ];
        const authorized = allowedSites.some(path => this.props.siteUrl.includes(path));
        if (!authorized) {
            return (React.createElement("div", { style: { padding: 16, color: 'red', fontWeight: 'bold' } }, "\u26A0\uFE0F This webpart is not authorized to be loaded in this site."));
        }
        const filteredItems = this.applyFilters(items);
        return (React.createElement("div", { className: styles.container },
            React.createElement("h3", null, "Todos los documentos"),
            React.createElement("div", { style: { marginBottom: '1rem' } },
                React.createElement("input", { type: "text", placeholder: "Buscar por nombre...", value: this.state.searchTerm, onChange: this.onSearchChange, className: styles.searchBox })),
            this.props.customColumns.map(col => (React.createElement("div", { key: col.internalName, className: styles.filterWrapper },
                React.createElement("label", { className: styles.filterLabel, htmlFor: col.internalName },
                    React.createElement("strong", null, col.label)),
                React.createElement("select", { className: styles.filterDropdown, onFocus: e => { e.currentTarget.style.borderColor = '#0078d4'; }, onBlur: e => { e.currentTarget.style.borderColor = '#8a8886'; }, id: col.internalName, value: filters[col.internalName] || '', onChange: e => this.onFilterChanged(col.internalName, e) },
                    React.createElement("option", { value: '' }, "Todos"),
                    [...(filterOptions[col.internalName] || [])].map(option => (React.createElement("option", { key: option, value: option }, option))))))),
            React.createElement("table", { className: styles.fluentliketable },
                React.createElement("thead", null,
                    React.createElement("tr", null,
                        this.renderSortableHeader('name', 'Nombre'),
                        this.renderSortableHeader('modified', 'Modificado'),
                        this.renderSortableHeader('modifiedBy', 'Modificado por'),
                        this.renderSortableHeader('library', 'Biblioteca'),
                        this.props.customColumns.map(col => this.renderSortableHeader(col.internalName, col.label)))),
                React.createElement("tbody", null, filteredItems.map((item, idx) => (React.createElement("tr", { key: idx },
                    React.createElement("td", null,
                        React.createElement("a", { href: item.editUrl, target: "_blank", rel: "noreferrer", className: styles.linkStyle }, item.name)),
                    React.createElement("td", null, new Date(item.modified).toLocaleString()),
                    React.createElement("td", null, item.modifiedBy),
                    React.createElement("td", null, item.library),
                    this.props.customColumns.map(col => (React.createElement("td", { key: col.internalName }, item.customColumns[col.internalName]))))))))));
    }
}
//# sourceMappingURL=AllDocuments.js.map