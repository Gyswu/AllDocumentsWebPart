import * as React from 'react';
import { getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';

export interface IAllDocumentsProps {}

export interface IDocumentItem {
  name: string;
  fileType: string;
  modified: Date;
  modifiedBy: string;
  url: string;
  library: string;
}

export interface IAllDocumentsState {
  items: IDocumentItem[];
  filterName: string;
  filterType: string;
  filterUser: string;
  filterDateFrom?: string;
  filterDateTo?: string;
  loading: boolean;
}

export default class AllDocuments extends React.Component<IAllDocumentsProps, IAllDocumentsState> {
  private sp: SPFI;

  constructor(props: IAllDocumentsProps) {
    super(props);
    this.state = {
      items: [],
      filterName: '',
      filterType: 'Todos',
      filterUser: '',
      filterDateFrom: undefined,
      filterDateTo: undefined,
      loading: false,
    };

    this.handleFilterChange = this.handleFilterChange.bind(this);
    this.sp = getSP();
  }

  public componentDidMount(): void {
    this.loadAllLibraries();
  }

  private async loadAllLibraries(): Promise<void> {
    this.setState({ loading: true });

    try {
      const libs = await this.sp.web.lists
        .filter("BaseTemplate eq 101")
        .select("Title", "Id")();

      console.log("Bibliotecas detectadas:", libs);

      const allItems: IDocumentItem[] = [];

      for (const lib of libs) {
        try {
          const items = await this.sp.web.lists
            .getById(lib.Id)
            .items
            .filter("FSObjType eq 0")
            .select("FileLeafRef", "File_x0020_Type", "Modified", "Editor/Title", "FileRef")
            .expand("Editor")();

          const mapped = items.map(item => ({
            name: item.FileLeafRef,
            fileType: item.File_x0020_Type,
            modified: new Date(item.Modified),
            modifiedBy: item.Editor?.Title || '',
            url: window.location.origin + item.FileRef,
            library: lib.Title
          }));

          allItems.push(...mapped);
        } catch (err) {
          console.warn(`No se pudo acceder a la biblioteca: ${lib.Title}`, err);
        }
      }

      this.setState({ items: allItems, loading: false });
    } catch (error) {
      console.error("Error cargando bibliotecas:", error);
      this.setState({ loading: false });
    }
  }

  private handleFilterChange(e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>): void {
    const { name, value } = e.target;
    this.setState({ [name]: value } as unknown as Pick<IAllDocumentsState, keyof IAllDocumentsState>);
  }

  public render(): React.ReactElement<IAllDocumentsProps> {
    const filteredItems = this.state.items.filter(item => {
      const matchesName = this.state.filterName === '' || item.name.toLowerCase().includes(this.state.filterName.toLowerCase());
      const matchesUser = this.state.filterUser === '' || (item.modifiedBy && item.modifiedBy.toLowerCase().includes(this.state.filterUser.toLowerCase()));
      const matchesType = this.state.filterType === 'Todos' || item.fileType === this.state.filterType;
      const matchesFromDate = !this.state.filterDateFrom || item.modified >= new Date(this.state.filterDateFrom);
      const matchesToDate = !this.state.filterDateTo || item.modified <= new Date(this.state.filterDateTo);
      return matchesName && matchesUser && matchesType && matchesFromDate && matchesToDate;
    });

    const fileTypes = Array.from(new Set(this.state.items.map(i => i.fileType).filter(ft => ft && ft !== ''))).sort();
    fileTypes.unshift('Todos');

    return (
      <div>
        <div style={{ padding: '8px', background: '#f3f3f3' }}>
          <strong>Filtros:</strong>{' '}
          Nombre: <input type="text" name="filterName" value={this.state.filterName} onChange={this.handleFilterChange} />{' '}
          Tipo de archivo: <select name="filterType" value={this.state.filterType} onChange={this.handleFilterChange}>
            {fileTypes.map(type => <option key={type} value={type}>{type}</option>)}
          </select>{' '}
          Modificado por: <input type="text" name="filterUser" value={this.state.filterUser} onChange={this.handleFilterChange} />{' '}
          Desde: <input type="date" name="filterDateFrom" value={this.state.filterDateFrom || ''} onChange={this.handleFilterChange} />{' '}
          Hasta: <input type="date" name="filterDateTo" value={this.state.filterDateTo || ''} onChange={this.handleFilterChange} />
        </div>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead style={{ background: '#ddd' }}>
            <tr>
              <th style={{ textAlign: 'left', padding: '4px' }}>Nombre</th>
              <th style={{ textAlign: 'left', padding: '4px' }}>Tipo</th>
              <th style={{ textAlign: 'left', padding: '4px' }}>Modificado</th>
              <th style={{ textAlign: 'left', padding: '4px' }}>Modificado por</th>
              <th style={{ textAlign: 'left', padding: '4px' }}>Biblioteca</th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.map((item, idx) => (
              <tr key={idx}>
                <td style={{ padding: '4px' }}>
                  <a href={item.url} target="_blank" rel="noopener noreferrer">{item.name}</a>
                </td>
                <td style={{ padding: '4px' }}>{item.fileType}</td>
                <td style={{ padding: '4px' }}>{item.modified.toLocaleDateString()}</td>
                <td style={{ padding: '4px' }}>{item.modifiedBy}</td>
                <td style={{ padding: '4px' }}>{item.library}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {this.state.loading && (
          <div style={{ textAlign: 'center', padding: '8px' }}>
            Cargando documentos...
          </div>
        )}
      </div>
    );
  }
}
