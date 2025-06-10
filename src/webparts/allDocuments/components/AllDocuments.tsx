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
                ${this.props.customColumns.map(col => `<FieldRef Name='${col}' />`).join('')}
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
          // Skip folders (FSObjType = 1)
          if (file.FSObjType === "1" || file.FSObjType === 1) continue;
  
          const fileName = file.FileLeafRef;
          const filePath = file.FileRef;
          const modified = file.Modified;
          const editor = file.Editor?.[0]?.title || file.Editor?.title || file.Editor || '';
  
          const customData: { [key: string]: string } = {};
          for (const col of this.props.customColumns) {
            const value = file[col];
            customData[col] = value || '';
            if (!filterOptions[col]) filterOptions[col] = new Set<string>();
            if (value) filterOptions[col].add(value);
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
          <div key={col} className={styles.filterWrapper}>
            <label className={styles.filterLabel} htmlFor={col}><strong>{col}</strong></label>
            <select
              className={styles.filterDropdown}
              onFocus={e => {
                e.currentTarget.style.borderColor = '#0078d4';
              }}
              onBlur={e => {
                e.currentTarget.style.borderColor = '#8a8886';
              }}
              id={col}
              value={filters[col] || ''}
              onChange={e => this.onFilterChanged(col, e)
              }
            >
              <option value=''>Todos</option>
              {[...(filterOptions[col] || [])].map(option => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>
        ))}

        <table className={styles.fluentliketable}>
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
                <td><a href={item.editUrl} target="_blank" rel="noreferrer" className={styles.linkStyle} >{item.name}</a></td>
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
