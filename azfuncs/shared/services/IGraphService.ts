export interface IGraphService {
  UploadFileToSite(siteId: string, driveId: string, path: string, name: string, blob: Buffer): Promise<any>;
  DownloadFileFromSite(siteId: string, driveId: string, itemId: string): Promise<Blob>;
}