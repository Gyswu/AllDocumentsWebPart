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
import { getSP } from '../pnpjsConfig';
export default class AllDocuments extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            items: [],
            customFields: [],
            loading: true,
        };
        this.sp = getSP();
    }
    componentDidMount() {
        this.loadAllDocuments();
    }
    loadAllDocuments() {
        return __awaiter(this, void 0, void 0, function* () {
            this.setState({ loading: true });
            try {
                const libs = yield this.sp.web.lists
                    .filter("BaseTemplate eq 101")
                    .select("Title", "Id")();
                const allItems = [];
                const customFieldsMap = new Map();
                for (const lib of libs) {
                    const fields = yield this.sp.web.lists.getById(lib.Id).fields
                        .filter("Group eq 'Columnas personalizadas' or Group eq 'Custom Columns'")
                        .select("Title", "InternalName")();
                    fields.forEach((f) => customFieldsMap.set(f.InternalName, f.Title));
                    const customInternalNames = Array.from(customFieldsMap.keys());
                    const selectFields = ["FileLeafRef", "Modified", "Editor/Title", "FileRef", ...customInternalNames];
                    const items = yield this.sp.web.lists.getById(lib.Id).items
                        .select(...selectFields)
                        .expand("Editor")();
                    const mappedItems = items.map(item => {
                        var _a;
                        const base = {
                            name: item.FileLeafRef,
                            modified: new Date(item.Modified),
                            modifiedBy: ((_a = item.Editor) === null || _a === void 0 ? void 0 : _a.Title) || '',
                            url: window.location.origin + item.FileRef,
                            library: lib.Title
                        };
                        customInternalNames.forEach(key => base[key] = item[key]);
                        return base;
                    });
                    allItems.push(...mappedItems);
                }
                const customFields = Array.from(customFieldsMap.entries()).map(([internalName, title]) => ({
                    internalName,
                    title
                }));
                this.setState({ items: allItems, customFields, loading: false });
            }
            catch (err) {
                console.error("Error loading documents:", err);
                this.setState({ loading: false });
            }
        });
    }
    render() {
        return (React.createElement("div", null,
            React.createElement("h3", null, "Todos los documentos"),
            React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse' } },
                React.createElement("thead", { style: { background: '#ddd' } },
                    React.createElement("tr", null,
                        React.createElement("th", null, "Nombre"),
                        React.createElement("th", null, "Modificado"),
                        React.createElement("th", null, "Modificado por"),
                        React.createElement("th", null, "Biblioteca"),
                        this.state.customFields.map(cf => (React.createElement("th", { key: cf.internalName }, cf.title))))),
                React.createElement("tbody", null, this.state.items.map((item, idx) => (React.createElement("tr", { key: idx },
                    React.createElement("td", null,
                        React.createElement("a", { href: item.url, target: "_blank", rel: "noreferrer" }, item.name)),
                    React.createElement("td", null, item.modified.toLocaleDateString()),
                    React.createElement("td", null, item.modifiedBy),
                    React.createElement("td", null, item.library),
                    this.state.customFields.map(cf => (React.createElement("td", { key: cf.internalName }, item[cf.internalName])))))))),
            this.state.loading && React.createElement("p", null, "Cargando...")));
    }
}
//# sourceMappingURL=AllDocuments.js.map