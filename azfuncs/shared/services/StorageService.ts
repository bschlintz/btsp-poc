import {   
  Aborter,
  ContainerURL,
  ServiceURL,
  StorageURL,
  SharedKeyCredential,
  Models,
  BlockBlobURL,
  uploadStreamToBlockBlob,
  IUploadStreamToBlockBlobOptions,
} from '@azure/storage-blob';
import { IFileInfo } from '../models/IFileInfo';
import { Readable } from 'stream';
import * as mime from 'mime-types';

const ONE_MEGABYTE = 1024 * 1024;
const FOUR_MEGABYTES = 4 * ONE_MEGABYTE;
const THIRTY_SECONDS = 30 * 1000;

class StorageService {

  private _serviceUrl: ServiceURL;

  constructor(accountName: string, accountKey: string) {
    this._serviceUrl = this._getServiceUrl(accountName, accountKey);
  }

  public async GetBlobList(containerName: string): Promise<IFileInfo[]> {
    let blobInfos: IFileInfo[] = [];

    try {
      const containerUrl = ContainerURL.fromServiceURL(this._serviceUrl, containerName);
      let marker = undefined;
  
      do {
        const aborter = Aborter.timeout(THIRTY_SECONDS);    
        const listBlobsResponse: Models.ContainerListBlobFlatSegmentResponse = await containerUrl.listBlobFlatSegment(aborter, marker);
  
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
    }
    catch (error) {
      throw new Error(`Unable to get blob list. ${error.message}`);
    }

    return blobInfos;
  }

  public async UploadBlobContent(containerName: string, blobName: string, blobBuffer: Buffer): Promise<any> {
    let response = null;

    try {
      const containerUrl = ContainerURL.fromServiceURL(this._serviceUrl, containerName);
      const blobUrl = BlockBlobURL.fromContainerURL(containerUrl, blobName);
  
      const aborter = Aborter.timeout(THIRTY_SECONDS);    
      const bufferSize = FOUR_MEGABYTES;
      const maxBuffers = 5;
      const uploadOpts: IUploadStreamToBlockBlobOptions = {
        blobHTTPHeaders: {
          blobContentType: mime.lookup(blobName) || 'application/octet-stream'
        }
      }
  
      //Convert buffer to stream for more performant uploads
      const blobStream = this._bufferToReadable(blobBuffer);
  
      response = await uploadStreamToBlockBlob(aborter, blobStream, blobUrl, bufferSize, maxBuffers, uploadOpts);
    }
    catch (error) {
      throw new Error(`Unable to upload blob content. ${error.message}`);
    }

    return response;
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

  private _bufferToReadable = (buffer: Buffer): Readable => {
    let readable = new Readable();
    readable._read = () => {}; 
    readable.push(buffer); 
    readable.push(null);
    return readable;
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