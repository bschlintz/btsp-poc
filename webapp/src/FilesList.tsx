import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { IFilesListItem, GetFilesList, BeginEditSession, EndEditSession, NewItem } from './Api';

const ITEMS_CACHE_KEY = "BTSPItemsCache";

export interface IFilesListState {
  items: IFilesListItem[];
  error: any;
}

export default class FilesList extends React.Component<{}, IFilesListState> {
  private _columns: IColumn[];

  constructor(props: {}) {
    super(props);

    this._columns = [
      { key: 'operations', name: '', fieldName: 'operations', minWidth: 70, maxWidth: 70, isResizable: false },
      { key: 'name', name: 'Name', fieldName: 'name', minWidth: 300, maxWidth: 500, isResizable: true },
      { key: 'modified', name: 'Last Modified', fieldName: 'modified', minWidth: 150, maxWidth: 250, isResizable: true },
    ];

    this.state = {
      items: [],
      error: ''
    };
  }

  public async componentDidMount() {
    let cachedItems = this._getCache(ITEMS_CACHE_KEY);
    if (cachedItems) {
      //Reset loading status back to false if any were cached while loading
      cachedItems = cachedItems.map((item: IFilesListItem) => { 
        item.isLoading = false; 
        return item;
      })
      this.setState({items: cachedItems});
    }
    await this._refreshItems();
  }

  private _refreshItems = async (): Promise<void> => {
    try {
      const items = await GetFilesList();
      const mergedItems = this._mergeItemCollections(items, this.state.items);
      this.setState({ items: mergedItems });
      this._setCache(ITEMS_CACHE_KEY, mergedItems);
    }
    catch (error) {
      this.setState({ error });
    }
  }

  private _mergeItemCollections = (itemsColA: IFilesListItem[], itemsColB: IFilesListItem[]): IFilesListItem[] => {
    if (itemsColA && itemsColA.length > 0) {
      if (itemsColB && itemsColB.length > 0) {
        return itemsColA.map(newItem => {
          let existingItemMatch = itemsColB.filter(ei => ei.key === newItem.key);
          let existingItem = existingItemMatch.length === 1 ? existingItemMatch[0] : null;
          return !existingItem ? newItem : { ...newItem, ...existingItem, modified: newItem.modified };
        });
      }
      else {
        return itemsColA;
      }
    }
    else if (itemsColB && itemsColB.length > 0) {
      return itemsColB;
    }
    else {
      return [];
    }
  }

  private _startEditFile = async (item: IFilesListItem): Promise<void> => {

    this._updateItem(item.key, { isLoading: true });    
    
    try {
      const editSession = await BeginEditSession(item.name);
      console.log(editSession);
      this._updateItem(item.key, { ...editSession, isLoading: false, isEditing: true });
    }
    catch (error) {
      this.setState({ error });
    }
  }

  private _openFile = async (item: IFilesListItem): Promise<void> => {
    // this._updateItem(item.key, { isEditing: true });
    window.open(item.webUrl, '_blank');
  }

  private _stopEditFile = async (item: IFilesListItem, discardEdit: boolean = false) => {

    this._updateItem(item.key, { isLoading: true });

    try {
      await EndEditSession(item.name, item.itemId, discardEdit);
      this._updateItem(item.key, NewItem(item.name, item.modified));
      await this._refreshItems();
    }
    catch (error) {
      this.setState({ error });
    }
  }

  private _updateItem = (key: string, updates: any) => {
    let updatedItems = this.state.items.map(i => {
      return i.key === key ? { ...i, ...updates } : i;
    })
    if (JSON.stringify(updatedItems) !== JSON.stringify(this.state.items)) {
      this.setState({ items: updatedItems });
      this._setCache(ITEMS_CACHE_KEY, updatedItems);
    }
  }

  private _setCache = (key: string, data: any): any => {
    sessionStorage.setItem(key, JSON.stringify(data));
    return this._getCache(key);
  }

  private _getCache = (key: string): any => {
    const data = sessionStorage.getItem(key);
    return data ? JSON.parse(data) : null;
  }

  private _renderItemColumn = (item: IFilesListItem, index?: number, column?: any) => {
    const fieldContent = item[column.fieldName as keyof IFilesListItem] as string;
    const iconStyle = { height: 18, width: 24 };
    switch (column.key) {
      case 'operations':
        return (
          <div style={{textAlign: 'center'}}>
            {item.isLoading
             ? <Spinner size={SpinnerSize.small} />
             : (item.isEditing 
                ? (<>
                    <IconButton iconProps={{ iconName: 'OpenInNewWindow' }} title="Open" onClick={_ => this._openFile(item)} style={iconStyle} />
                    <IconButton iconProps={{ iconName: 'Completed' }} title="Finish Editing" onClick={_ => this._stopEditFile(item)} style={iconStyle} />
                    <IconButton iconProps={{ iconName: 'ErrorBadge' }} title="Discard" onClick={_ => this._stopEditFile(item, true)} style={iconStyle} />
                  </>)
                : (<>
                    <IconButton iconProps={{ iconName: 'EditSolid12' }} title="Edit" onClick={_ => this._startEditFile(item)} style={iconStyle} />
                  </>)
             )            
            }
          </div>
        );    
      case 'modified':  
        const date = new Date(fieldContent);
        return <div>{date.toLocaleString()}</div>;
      default:
        return <div>{fieldContent}</div>;
    }
  }

  public render(): JSX.Element {
    const { items } = this.state;

    return (
      <Fabric>
        {this.state.error ? <div style={{color: 'red'}}>{this.state.error.toString()}</div> : ''}
        <DetailsList
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