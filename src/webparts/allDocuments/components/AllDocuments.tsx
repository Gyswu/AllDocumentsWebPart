import * as React from 'react';
import { IAllDocumentsProps } from './IAllDocumentsProps';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './AllDocuments.module.scss';

export interface IDocumentItem {
  name: string;
  modified: string;
  modifiedBy: string;
  library: string;
  editUrl: string;
  customColumns: { [key: string]: string };
}

interface IState {
  items: IDocumentItem[];
  loading: boolean;
  filters: { [key: string]: string };
  filterOptions: { [key: string]: Set<string> };
  searchTerm: string;
  sortConfig: {
    column: string;
    direction: 'asc' | 'desc';
  } | null;
}

export default class AllDocuments extends React.Component<IAllDocumentsProps, IState> {
  constructor(props: IAllDocumentsProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      filters: {},
      filterOptions: {},
      searchTerm: '',
      sortConfig: null
    };
  }

  public async componentDidMount(): Promise<void> {
    
    try {
      const res = await this.props.spHttpClient.get(
        `${this.props.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Id,Title`,
        SPHttpClient.configurations.v1
      );
      const json = await res.json();
      const libraries = json.value;

      const allItems: IDocumentItem[] = [];
      const filterOptions: { [key: string]: Set<string> } = {};

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

        const itemsRes = await this.props.spHttpClient.post(
          `${this.props.siteUrl}/_api/web/lists(guid'${lib.Id}')/RenderListDataAsStream`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({ parameters: camlQuery })
          }
        );

        const itemsJson = await itemsRes.json();
        const rows = itemsJson?.Row;
        if (!rows || rows.length === 0) {
          console.warn(`No items found in library: ${lib.Title}`);
          continue;
        }

        for (const file of rows) {
          if (file.FSObjType === "1" || file.FSObjType === 1) continue;

          const fileName = file.FileLeafRef;
          const filePath = file.FileRef;
          const modified = file.Modified;
          const editor = file.Editor?.[0]?.title || file.Editor?.title || file.Editor || '';

          const customData: { [key: string]: string } = {};
          for (const col of this.props.customColumns) {
            const value = file[col.internalName];
            customData[col.internalName] = value || '';
            if (!filterOptions[col.internalName]) filterOptions[col.internalName] = new Set<string>();
            if (value) filterOptions[col.internalName].add(value);
          }

          allItems.push({
            name: fileName,
            modified: modified,
            modifiedBy: editor,
            library: lib.Title,
            editUrl: `${this.props.siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(filePath)}&action=edit&mobileredirect=true`,
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

      console.log("Loaded document items:", allItems);
    } catch (err) {
      console.error("Error loading documents:", err);
      this.setState({ loading: false });
    }
  }

  private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ searchTerm: e.target.value });
  };

  private onSortColumn = (column: string): void => {
    const { sortConfig } = this.state;
    let direction: 'asc' | 'desc' = 'asc';

    if (sortConfig && sortConfig.column === column && sortConfig.direction === 'asc') {
      direction = 'desc';
    }

    this.setState({ sortConfig: { column, direction } });
  };

  private renderSortableHeader = (column: string, label: string): JSX.Element => {
    const { sortConfig } = this.state;
    const isSorted = sortConfig?.column === column;
    const direction = isSorted ? (sortConfig.direction === 'asc' ? '▲' : '▼') : '';

    return (
      <th onClick={() => this.onSortColumn(column)} className={styles.sortableHeader}>
        {label} {direction}
      </th>
    );
  };

  private onFilterChanged = (column: string, e: React.ChangeEvent<HTMLSelectElement>): void => {
    const newFilters = { ...this.state.filters, [column]: e.target.value };
    this.setState({ filters: newFilters });
  };

  private applyFilters(items: IDocumentItem[]): IDocumentItem[] {
    const { filters, searchTerm, sortConfig } = this.state;

    let filtered = items.filter(item =>
      Object.entries(filters).every(([key, val]) => val === '' || item.customColumns[key] === val)
    );

    if (searchTerm) {
      filtered = filtered.filter(item =>
        item.name.toLowerCase().includes(searchTerm.toLowerCase())
      );
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

  public render(): React.ReactElement<IAllDocumentsProps> {
    const {items, filters, filterOptions } = this.state;
    const allowedSites = [
      "sites/sp-FIN"
    ]

    const authorized =allowedSites.some(path => this.props.siteUrl.includes(path))
    if (!authorized) {
      return (
        <div style={{ padding: 16, color: 'red', fontWeight: 'bold' }}>
          ⚠️ This webpart is not authorized to be loaded in this site.
        </div>
      );
    }
  
    const filteredItems = this.applyFilters(items);

    return (
      <div className={styles.container}>
        <h3>Todos los documentos</h3>

        <div style={{ marginBottom: '1rem' }}>
          <input
            type="text"
            placeholder="Buscar por nombre..."
            value={this.state.searchTerm}
            onChange={this.onSearchChange}
            className={styles.searchBox}
          />
        </div>

        {this.props.customColumns.map(col => (
          <div key={col.internalName} className={styles.filterWrapper}>
            <label className={styles.filterLabel} htmlFor={col.internalName}><strong>{col.label}</strong></label>
            <select
              className={styles.filterDropdown}
              onFocus={e => { e.currentTarget.style.borderColor = '#0078d4'; }}
              onBlur={e => { e.currentTarget.style.borderColor = '#8a8886'; }}
              id={col.internalName}
              value={filters[col.internalName] || ''}
              onChange={e => this.onFilterChanged(col.internalName, e)}
            >
              <option value=''>Todos</option>
              {[...(filterOptions[col.internalName] || [])].map(option => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
        ))}

        <table className={styles.fluentliketable}>
          <thead>
            <tr>
              {this.renderSortableHeader('name', 'Nombre')}
              {this.renderSortableHeader('modified', 'Modificado')}
              {this.renderSortableHeader('modifiedBy', 'Modificado por')}
              {this.renderSortableHeader('library', 'Biblioteca')}
              {this.props.customColumns.map(col =>
                this.renderSortableHeader(col.internalName, col.label)
              )}
            </tr>
          </thead>
          <tbody>
            {filteredItems.map((item, idx) => (
              <tr key={idx}>
                <td><a href={item.editUrl} target="_blank" rel="noreferrer" className={styles.linkStyle}>{item.name}</a></td>
                <td>{new Date(item.modified).toLocaleString()}</td>
                <td>{item.modifiedBy}</td>
                <td>{item.library}</td>
                {this.props.customColumns.map(col => (
                  <td key={col.internalName}>{item.customColumns[col.internalName]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}
