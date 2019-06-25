import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';

const BTSP_FILEMANAGER_API_URL = "http://localhost:7071/api/FileManager";
const BTSP_EDITBEGIN_API_URL = "http://localhost:7071/api/FileEditBegin";
const BTSP_EDITEND_API_URL = "http://localhost:7071/api/FileEditEnd";

export interface IFilesListItem {
  key: string;
  name: string;
  modified: number;
  isLoading: boolean;
  isEditing: boolean;
  webUrl: string;
}

export interface IFilesListState {
  items: IFilesListItem[];
  error: string;
}

export default class FilesList extends React.Component<{}, IFilesListState> {
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this._columns = [
      { key: 'operations', name: '', fieldName: 'operations', minWidth: 60, maxWidth: 60, isResizable: false },
      { key: 'name', name: 'Name', fieldName: 'name', minWidth: 300, maxWidth: 500, isResizable: true },
      { key: 'modified', name: 'Last Modified', fieldName: 'modified', minWidth: 150, maxWidth: 250, isResizable: true },
    ];

    this.state = {
      items: [],
      error: ''
    };
  }

  public async componentDidMount() {
    const response: Response = await fetch(BTSP_FILEMANAGER_API_URL);
    const result = await response.json();

    if (result.error) {
      this.setState({ error: result.error });
    } 
    else  {
      const items: IFilesListItem[] = result.data.map((file: { name: string; lastModified: Date; }) => ({
        key: file.name,
        name: file.name,
        modified: file.lastModified,
        isLoading: false,
        isEditing: false,
        webUrl: '',
      }))
      this.setState({ items })
    }
  }

  private _startEditFile = async (item: IFilesListItem) => {
    console.log(item);
    this._updateItem(item.key, { isLoading: true });
    const response: Response = await fetch(BTSP_EDITBEGIN_API_URL, {
      body: JSON.stringify({ blobName: item.name }),
      method: 'POST'
    });
    const result = await response.json();
    this._updateItem(item.key, { isLoading: false, webUrl: result.webUrl });

    console.log(result);
  }

  private _openFile = async (item: IFilesListItem) => {
    this._updateItem(item.key, { isEditing: true });
    window.open(item.webUrl, '_blank');
  }

  private _stopEditFile = async (item: IFilesListItem) => {
    this._updateItem(item.key, { isLoading: true });
    setTimeout(_ => this._updateItem(item.key, { isLoading: false, isEditing: false, webUrl: '' }), 2000);
    console.log(item);
  }

  private _updateItem = (key: string, updates: any) => {
    let updatedItems = this.state.items.map(i => {
      return i.key === key ? { ...i, ...updates } : i;
    })
    if (JSON.stringify(updatedItems) !== JSON.stringify(this.state.items)) {
      this.setState({ items: updatedItems });
    }
  }

  private _renderItemColumn = (item: IFilesListItem, index?: number, column?: any) => {
    const fieldContent = item[column.fieldName as keyof IFilesListItem] as string;

    switch (column.key) {
      case 'operations':
        return (
          <div style={{textAlign: 'center'}}>
            {item.isLoading
             ? <Spinner size={SpinnerSize.small} />
             : (item.isEditing 
                ? (<>
                    <IconButton iconProps={{ iconName: 'Completed' }} title="Finish Editing" style={{height: 18}} onClick={_ => this._stopEditFile(item)} />
                    <IconButton iconProps={{ iconName: 'ErrorBadge' }} title="Discard" style={{height: 18}} onClick={_ => this._stopEditFile(item)} />
                  </>)
                : (item.webUrl
                   ? <IconButton iconProps={{ iconName: 'OpenInNewWindow' }} title="Open" style={{height: 18}} onClick={_ => this._openFile(item)} />
                   : <IconButton iconProps={{ iconName: 'EditSolid12' }} title="Edit" style={{height: 18}} onClick={_ => this._startEditFile(item)} />
                )
             )            
            }
          </div>
        );    
      default:
        return <div>{fieldContent}</div>;
    }
  }

  public render(): JSX.Element {
    const { items } = this.state;

    return (
      <Fabric>
        {this.state.error && <div>{this.state.error}</div>}
        <DetailsList
          // compact={true}
          items={items}
          columns={this._columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          onRenderItemColumn={this._renderItemColumn}
        />
      </Fabric>
    );
  }
}