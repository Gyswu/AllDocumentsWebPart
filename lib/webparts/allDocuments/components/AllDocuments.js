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
export default class AllDocuments extends React.Component {
    constructor(props) {
        super(props);
        this.onFilterChanged = (column, e) => {
            const newFilters = Object.assign(Object.assign({}, this.state.filters), { [column]: e.target.value });
            this.setState({ filters: newFilters });
        };
        this.state = {
            items: [],
            loading: true,
            filters: {},
            filterOptions: {}
        };
    }
    componentDidMount() {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const res = yield this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`, SPHttpClient.configurations.v1);
                const json = yield res.json();
                const libraries = json.value;
                const allItems = [];
                const filterOptions = {};
                for (const lib of libraries) {
                    const libUrl = lib.RootFolder.ServerRelativeUrl;
                    const itemsRes = yield this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/getFolderByServerRelativeUrl('${libUrl}')/Files?$expand=ListItemAllFields,Author&$select=Name,TimeLastModified,Author/Title,ListItemAllFields/ID,ListItemAllFields/TestColumn,ListItemAllFields,Author`, SPHttpClient.configurations.v1);
                    const itemsJson = yield itemsRes.json();
                    const files = itemsJson.value;
                    for (const file of files) {
                        const customData = {};
                        for (const col of this.props.customColumns) {
                            const value = (_a = file.ListItemAllFields) === null || _a === void 0 ? void 0 : _a[col];
                            customData[col] = value || '';
                            if (!filterOptions[col])
                                filterOptions[col] = new Set();
                            if (value)
                                filterOptions[col].add(value);
                        }
                        allItems.push({
                            name: file.Name,
                            modified: file.TimeLastModified,
                            modifiedBy: ((_b = file.Author) === null || _b === void 0 ? void 0 : _b.Title) || '',
                            library: lib.Title,
                            editUrl: `${this.props.siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(libUrl + '/' + file.Name)}&action=edit&mobileredirect=true`,
                            customColumns: customData
                        });
                    }
                }
                this.setState({
                    items: allItems,
                    loading: false,
                    filters: {},
                    filterOptions: filterOptions
                });
            }
            catch (err) {
                console.error("Error loading documents:", err);
                this.setState({ loading: false });
            }
        });
    }
    applyFilters(items) {
        const { filters } = this.state;
        return items.filter(item => Object.entries(filters).every(([key, val]) => val === '' || item.customColumns[key] === val));
    }
    render() {
        const { items, filters, filterOptions } = this.state;
        const filteredItems = this.applyFilters(items);
        return (React.createElement("div", null,
            React.createElement("h3", null, "Todos los documentos"),
            this.props.customColumns.map(col => (React.createElement("div", { key: col, style: { marginBottom: 10 } },
                React.createElement("label", { htmlFor: col },
                    React.createElement("strong", null, col)),
                React.createElement("select", { id: col, value: filters[col] || '', onChange: e => this.onFilterChanged(col, e) },
                    React.createElement("option", { value: '' }, "Todos"),
                    [...(filterOptions[col] || [])].map(option => (React.createElement("option", { key: option, value: option }, option))))))),
            React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse' } },
                React.createElement("thead", { style: { background: '#ddd' } },
                    React.createElement("tr", null,
                        React.createElement("th", null, "Nombre"),
                        React.createElement("th", null, "Modificado"),
                        React.createElement("th", null, "Modificado por"),
                        React.createElement("th", null, "Biblioteca"),
                        this.props.customColumns.map(col => (React.createElement("th", { key: col }, col))))),
                React.createElement("tbody", null, filteredItems.map((item, idx) => (React.createElement("tr", { key: idx },
                    React.createElement("td", null,
                        React.createElement("a", { href: item.editUrl, target: "_blank", rel: "noreferrer" }, item.name)),
                    React.createElement("td", null, new Date(item.modified).toLocaleString()),
                    React.createElement("td", null, item.modifiedBy),
                    React.createElement("td", null, item.library),
                    this.props.customColumns.map(col => (React.createElement("td", { key: col }, item.customColumns[col]))))))))));
    }
}
//# sourceMappingURL=AllDocuments.js.map