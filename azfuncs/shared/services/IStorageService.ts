import { IFileInfo } from "../models/IFileInfo";

export interface IStorageService {
  GetBlobList(containerName: string): Promise<IFileInfo[]>;
  UploadBlobContent(containerName: string, blobName: string, blobStream: any): Promise<any>;
  // getBlobContent(containerName: string, blobName: string): Promise<any>;
}