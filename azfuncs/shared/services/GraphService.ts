import { IGraphService } from './IGraphService';
import { Client as GraphClient, AuthenticationProvider, LargeFileUploadTask, FileObject, LargeFileUploadTaskOptions } from "@microsoft/microsoft-graph-client";

class GraphService implements IGraphService {
  private _graphClient: GraphClient;
  private _writeLog: Function;

  constructor(authProvider?: AuthenticationProvider, accessToken?: string, writeLog?: Function) {
    this._writeLog = writeLog || (() => { });
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
    try {
      const uploadSession = await this._getUploadSession(siteId, driveId, path, name);   

      const size = blob.byteLength - blob.byteOffset;
      const content = blob.buffer.slice(blob.byteOffset, blob.byteOffset + blob.byteLength);
      const fileObj: FileObject = {
        content,
        size,
        name,
      }     
      const uploadOpts: LargeFileUploadTaskOptions = {
        rangeSize: (1024 * 1024)
      }
      
      const uploadTask = new LargeFileUploadTask(this._graphClient, fileObj, uploadSession, uploadOpts);
      const uploadResponse = await uploadTask.upload();

      return uploadResponse;

    } catch (err) {
      console.log(err);
    }
  }

  public async DownloadFileFromSite(siteId: string, driveId: string, itemId: string): Promise<any> {
    const url = `/sites/${siteId}/drives/${driveId}/items/${itemId}/content`;
    const response = await this._graphClient.api(url).getStream();
    // const itemResponse = await fetch(url);
    // const blobResult = await itemResponse.blob(); 
    return response;
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