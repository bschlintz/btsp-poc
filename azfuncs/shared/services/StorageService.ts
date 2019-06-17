import { IStorageService } from './IStorageService';
import {   
  Aborter,
  BlobURL,
  BlockBlobURL,
  ContainerURL,
  ServiceURL,
  StorageURL,
  SharedKeyCredential,
  TokenCredential,
  Models 
} from '@azure/storage-blob';
import { IFileInfo } from '../models/IFileInfo';


class StorageService implements IStorageService {

  private _serviceUrl: ServiceURL;

  constructor(accountName: string, accountKey: string) {
    this._serviceUrl = this._getServiceUrl(accountName, accountKey);
  }

  private _getServiceUrl(accountName: string, accountKey: string): ServiceURL {

    // Use SharedKeyCredential with storage account and account key
    const sharedKeyCredential = new SharedKeyCredential(accountName, accountKey);

    // Use sharedKeyCredential to create a pipeline
    const pipeline = StorageURL.newPipeline(sharedKeyCredential);

    // Construct service URL
    const serviceURL = new ServiceURL(`https://${accountName}.blob.core.windows.net`, pipeline);
        
    return serviceURL;
  }

  public async getFileList(containerName: string): Promise<IFileInfo[]> {

    const containerUrl = ContainerURL.fromServiceURL(this._serviceUrl, containerName);
    let blobInfos: IFileInfo[] = [];
    let marker = undefined;

    do {
      const listBlobsResponse: Models.ContainerListBlobFlatSegmentResponse = await containerUrl.listBlobFlatSegment(
        Aborter.none,
        marker
      );

      marker = listBlobsResponse.nextMarker;
      blobInfos.push(
        ...listBlobsResponse.segment.blobItems.map(blob => {
          return { 
            name: blob.name,
            contentType: blob.properties.contentType,
            lastModified: blob.properties.lastModified
          }
        })
      );
    } while (marker);

    return blobInfos;
  }

}

export default StorageService;