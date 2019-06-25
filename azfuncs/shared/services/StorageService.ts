import { IStorageService } from './IStorageService';
import {   
  Aborter,
  ContainerURL,
  ServiceURL,
  StorageURL,
  SharedKeyCredential,
  Models,
  BlockBlobURL,
  uploadStreamToBlockBlob,
  BlobUploadCommonResponse,
} from '@azure/storage-blob';
import { IFileInfo } from '../models/IFileInfo';
import { Stream } from 'stream';

const ONE_MEGABYTE = 1024 * 1024;
const FOUR_MEGABYTES = 4 * ONE_MEGABYTE;
const ONE_MINUTE = 60 * 1000;

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

  public async GetBlobList(containerName: string): Promise<IFileInfo[]> {

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

  public async UploadBlobContent(containerName: string, blobName: string, blobStream: any): Promise<any> {
    const containerUrl = ContainerURL.fromServiceURL(this._serviceUrl, containerName);
    const blobUrl = BlockBlobURL.fromContainerURL(containerUrl, blobName);

    const aborter = Aborter.timeout(ONE_MINUTE);

    const opts = {
      bufferSize: FOUR_MEGABYTES,
      maxBuffers: 5,
    };

    const uploadResult: BlobUploadCommonResponse = await uploadStreamToBlockBlob(aborter, blobStream, blobUrl, opts.bufferSize, opts.maxBuffers);
    return uploadResult;    
  }

  // public async getBlobContent(containerName: string, blobName: string): Promise<any> {
  //   const containerUrl = ContainerURL.fromServiceURL(this._serviceUrl, containerName);
  //   const blobUrl = BlockBlobURL.fromContainerURL(containerUrl, blobName);

  //   const ONE_MINUTE = 60 * 1000;
  //   const aborter = Aborter.timeout(ONE_MINUTE);

  //   const downloadResponse: Models.BlobDownloadResponse = await blobUrl.download(aborter, 0);
    
  //   const downloadContent = await this._streamToString(downloadResponse.readableStreamBody);

  //   return downloadContent;
  // }

  // // A helper method used to read a Node.js readable stream into string
  // private async _streamToString(readableStream) {
  //   return new Promise((resolve, reject) => {
  //     const chunks = [];
  //     readableStream.on("data", data => {
  //       chunks.push(data.toString());
  //     });
  //     readableStream.on("end", () => {
  //       resolve(chunks.join(""));
  //     });
  //     readableStream.on("error", reject);
  //   });
  // }
}

export default StorageService;