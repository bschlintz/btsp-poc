import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { runInNewContext } from 'vm';

const BTSP_API_URL = "http://localhost:7071/api/FileManager"

export interface IFilesListItem {
  key: number;
  name: string;
  modified: number;
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
      { key: 'operations', name: '', fieldName: 'operations', minWidth: 40, maxWidth: 40, isResizable: false },
      { key: 'name', name: 'Name', fieldName: 'name', minWidth: 300, maxWidth: 500, isResizable: true },
      { key: 'modified', name: 'Last Modified', fieldName: 'modified', minWidth: 150, maxWidth: 250, isResizable: true },
    ];

    this.state = {
      items: [],
      error: ''
    };
  }

  public async componentDidMount() {
    const response: Response = await fetch(BTSP_API_URL);
    const result = await response.json();

    if (result.error) {
      this.setState({ error: result.error });
    } 
    else  {
      const items: IFilesListItem[] = result.data.map((file: { name: string; lastModified: Date; }) => ({
        key: file.name,
        name: file.name,
        modified: file.lastModified
      }))
      this.setState({ items })
    }
  }

  private _editFile = (item: IFilesListItem) => {
    console.log(item);
  }

  private _renderItemColumn = (item: IFilesListItem, index?: number, column?: any) => {
    const fieldContent = item[column.fieldName as keyof IFilesListItem] as string;

    switch (column.key) {
      case 'operations':
        return (
          <div>
            <IconButton iconProps={{ iconName: 'EditSolid12' }} title="Edit" style={{height: 18}} ariaLabel="Edit" onClick={_ => this._editFile(item)} />
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