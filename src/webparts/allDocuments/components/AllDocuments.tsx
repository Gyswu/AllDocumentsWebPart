import * as React from 'react';
import { IAllDocumentsProps } from './IAllDocumentsProps';
import { SPHttpClient } from '@microsoft/sp-http';

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
}

export default class AllDocuments extends React.Component<IAllDocumentsProps, IState> {
  constructor(props: IAllDocumentsProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      filters: {},
      filterOptions: {}
    };
  }

  public async componentDidMount(): Promise<void> {
    try {
      const res = await this.props.spHttpClient.get(
        `${this.props.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        SPHttpClient.configurations.v1
      );
      const json = await res.json();
      const libraries = json.value;

      const allItems: IDocumentItem[] = [];
      const filterOptions: { [key: string]: Set<string> } = {};

      for (const lib of libraries) {
        const libUrl = lib.RootFolder.ServerRelativeUrl;
        const itemsRes = await this.props.spHttpClient.get(
          `${this.props.siteUrl}/_api/web/getFolderByServerRelativeUrl('${libUrl}')/Files?$expand=ListItemAllFields,Author&$select=Name,TimeLastModified,Author/Title,ListItemAllFields/ID,ListItemAllFields/TestColumn,ListItemAllFields,Author`,
          SPHttpClient.configurations.v1
        );
        const itemsJson = await itemsRes.json();
        const files = itemsJson.value;

        for (const file of files) {
          const customData: { [key: string]: string } = {};

          for (const col of this.props.customColumns) {
            const value = file.ListItemAllFields?.[col];
            customData[col] = value || '';
            if (!filterOptions[col]) filterOptions[col] = new Set<string>();
            if (value) filterOptions[col].add(value);
          }

          allItems.push({
            name: file.Name,
            modified: file.TimeLastModified,
            modifiedBy: file.Author?.Title || '',
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
    } catch (err) {
      console.error("Error loading documents:", err);
      this.setState({ loading: false });
    }
  }

  private onFilterChanged = (column: string, e: React.ChangeEvent<HTMLSelectElement>): void => {
    const newFilters = { ...this.state.filters, [column]: e.target.value };
    this.setState({ filters: newFilters });
  };

  private applyFilters(items: IDocumentItem[]): IDocumentItem[] {
    const { filters } = this.state;
    return items.filter(item =>
      Object.entries(filters).every(([key, val]) => val === '' || item.customColumns[key] === val)
    );
  }

  public render(): React.ReactElement<IAllDocumentsProps> {
    const { items, filters, filterOptions } = this.state;
    const filteredItems = this.applyFilters(items);

    return (
      <div>
        <h3>Todos los documentos</h3>

        {this.props.customColumns.map(col => (
          <div key={col} style={{ marginBottom: 10 }}>
            <label htmlFor={col}><strong>{col}</strong></label>
            <select
              id={col}
              value={filters[col] || ''}
              onChange={e => this.onFilterChanged(col, e)}
            >
              <option value=''>Todos</option>
              {[...(filterOptions[col] || [])].map(option => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
        ))}

        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead style={{ background: '#ddd' }}>
            <tr>
              <th>Nombre</th>
              <th>Modificado</th>
              <th>Modificado por</th>
              <th>Biblioteca</th>
              {this.props.customColumns.map(col => (
                <th key={col}>{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredItems.map((item, idx) => (
              <tr key={idx}>
                <td><a href={item.editUrl} target="_blank" rel="noreferrer">{item.name}</a></td>
                <td>{new Date(item.modified).toLocaleString()}</td>
                <td>{item.modifiedBy}</td>
                <td>{item.library}</td>
                {this.props.customColumns.map(col => (
                  <td key={col}>{item.customColumns[col]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}
