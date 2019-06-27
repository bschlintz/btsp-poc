import { Client as GraphClient, AuthenticationProvider, LargeFileUploadTask, FileObject, LargeFileUploadTaskOptions, MiddlewareOptions, RetryHandlerOptions, RetryHandler } from "@microsoft/microsoft-graph-client";
import * as request from 'request-promise-native';

class GraphService {
  private _graphClient: GraphClient;

  constructor(authProvider?: AuthenticationProvider, accessToken?: string) {
    if (authProvider) {
      this._graphClient = GraphClient.initWithMiddleware({ authProvider });
    }
    else if (accessToken) {
      this._graphClient = GraphClient.initWithMiddleware({ authProvider: { 
        getAccessToken: () => Promise.resolve(accessToken)
      }});
    }
    else {
      throw new Error(`Instantiating an instance of GraphService requires either an AuthenticationProvider or AccessToken.`);
    }
  }

  public async UploadFileToSite(siteId: string, driveId: string, path: string, name: string, blob: Buffer): Promise<any> {
    let uploadSession = null;

    try {
      uploadSession = await this._getUploadSession(siteId, driveId, path, name);   
    } 
    catch (error) {
      throw new Error(`Unable to get upload session. ${error.message}`);
    }

    try {
      const size = blob.byteLength - blob.byteOffset;
      const content = blob.buffer.slice(blob.byteOffset, blob.byteOffset + blob.byteLength);
      const fileObj: FileObject = { content, size, name };
      const uploadOpts: LargeFileUploadTaskOptions = {
        rangeSize: (1024 * 1024)
      }
      
      const uploadTask = new LargeFileUploadTask(this._graphClient, fileObj, uploadSession, uploadOpts);
      const uploadResponse = await uploadTask.upload();

      return uploadResponse;

    } 
    catch (error) {
      throw new Error(`Unable to upload blob. ${error.message}`);
    }
  }

  public async DownloadFileFromSite(siteId: string, driveId: string, itemId: string): Promise<Buffer> {
    const itemUrl = `/sites/${siteId}/drives/${driveId}/items/${itemId}`;
    let response = null;

    try {
      const itemResponse = await this._graphClient.api(itemUrl).version('v1.0').get();
  
      const downloadUrl = itemResponse["@microsoft.graph.downloadUrl"];
      const downloadOpts = {
        url: downloadUrl,
        resolveWithFullResponse: true,
        encoding: null
      }
      
      const downloadResponse = await request.get(downloadOpts);
      response = downloadResponse.body;
    }
    catch (error) {
      throw new Error(`Unable to download file. ${error.message}`);
    }

    return response;
  }

  public async DeleteFileFromSite(siteId: string, driveId: string, itemId: string): Promise<any> {
    const itemUrl = `/sites/${siteId}/drives/${driveId}/items/${itemId}`;
    const FILE_LOCK_MESSAGE = 'The file is currently checked out or locked for editing by another user.';
    const MAX_RETRIES = 3;
    const DELAY = 1;
    let response = null;
    
    try {
      let isDeleteSuccess = false;
      let retryAttempt = 1;
      do {        
        try {
          response = await this._graphClient.api(itemUrl).delete(); 
          isDeleteSuccess = true;       
        }
        catch (error) {
          if (error.message.indexOf(FILE_LOCK_MESSAGE) > -1 && retryAttempt <= MAX_RETRIES) {
            await this._sleep(DELAY);
            retryAttempt += 1;
          }
          else {
            throw error;
          }
        }
      } while(!isDeleteSuccess && retryAttempt <= MAX_RETRIES);
    }
    catch (error) {
      throw new Error(`Unable to delete file. ${error.message}`);
    }

    return response;
  }

  private async _sleep(delaySeconds: number): Promise<void> {
		const delayMilliseconds = delaySeconds * 1000;
		return new Promise((resolve) => setTimeout(resolve, delayMilliseconds));
	}

  private async _getUploadSession(siteId: string, driveId: string, path: string, name: string): Promise<any> {
    const url = `/sites/${siteId}/drives/${driveId}/root:/${path}/${name}:/createUploadSession`;
    const sessionOptions = { 
      "item": { 
        "@microsoft.graph.conflictBehavior": "fail",
      } 
    };
    const uploadSession = await LargeFileUploadTask.createUploadSession(this._graphClient, encodeURI(url), sessionOptions);
    
    return uploadSession;
  }  
}

export default GraphService;