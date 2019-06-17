import { IFileInfo } from "../models/IFileInfo";

export interface IStorageService {
  getFileList(containerName: string): Promise<IFileInfo[]>;
}