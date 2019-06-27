const BTSP_FILEMANAGER_API_URL = "http://localhost:7071/api/FileManager";
const BTSP_EDITBEGIN_API_URL = "http://localhost:7071/api/FileEditBegin";
const BTSP_EDITEND_API_URL = "http://localhost:7071/api/FileEditEnd";

export interface IFilesListItem {
  key: string;
  name: string;
  modified: Date;
  isLoading: boolean;
  isEditing: boolean;
  webUrl: string;
  itemId: string;
}

export interface StartEditSessionResult {
  webUrl: string;
  itemId: string;
}

export interface ICommonResponse {
  error?: string;
  message?: string;
  data?: any;
}

export const GetFilesList = async (): Promise<IFilesListItem[]> => {
  const response: Response = await fetch(BTSP_FILEMANAGER_API_URL);
  const result: ICommonResponse = await response.json();

  if (result.error) {
    throw new Error(result.error);
  }

  const fileInfos: IFilesListItem[] = result.data.map(
    (file: { name: string; lastModified: Date; }) => NewItem(file.name, file.lastModified)
  );

  return fileInfos;
}

export const BeginEditSession = async (blobName: string): Promise<StartEditSessionResult> => {
  const payload = { blobName };
  const response: Response = await fetch(BTSP_EDITBEGIN_API_URL, {
    body: JSON.stringify(payload),
    method: 'POST'
  });
  const result: ICommonResponse = await response.json();

  if (result.error) {
    throw new Error(result.error);
  }

  return result.data;
}

export const EndEditSession = async (blobName: string, itemId: string, discard: boolean): Promise<void> => {
  const payload = { blobName, itemId, discard };
  const response = await fetch(BTSP_EDITEND_API_URL, {
    body: JSON.stringify(payload),
    method: 'POST'
  });
  const result: ICommonResponse = await response.json();

  if (result.error) {
    throw new Error(result.error);
  }
}

export const NewItem = (name: string, modified: Date): IFilesListItem => {
  return {
    name,
    modified,
    key: name,
    isLoading: false,
    isEditing: false,
    webUrl: '',
    itemId: '',
  };
}