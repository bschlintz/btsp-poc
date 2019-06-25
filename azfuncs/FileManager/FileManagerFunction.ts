import * as df from "durable-functions";
import { Context, HttpRequest } from "@azure/functions";
import { IStorageService } from "../shared/services/IStorageService";
import { IGraphService } from "../shared/services/IGraphService";
import config from "../shared/config";
import { IFileManagerResponse } from "./IFileManagerResponse";
import { IFileInfo } from "../shared/models/IFileInfo";
import { IHttpResponse } from "durable-functions/lib/src/classes";

class FileManagerFunction {
  private _storageService: IStorageService;
  private _adminGraphService: IGraphService;

  constructor(storageService: IStorageService, adminGraphService: IGraphService) {
    this._storageService = storageService;
    this._adminGraphService = adminGraphService;
  }

  public async processRequest(context: Context, req: HttpRequest): Promise<Response> {
    switch (context.req.method) {
      case "GET":  {
        return await this._handleHttpGet(context, req);
      }
      case "POST": {
        const blobName = req.body && req.body.blobName;
        if (!blobName) {
          return this._getJsonRes(400, { error: "You must provide a POST body when using this endpoint with HTTP POST. Example POST body: {'name': '<PATH TO FILENAME IN BLOB STOARGE>'}" })
        }
        return await this._handleHttpPost(context, req);
      }
  
      default:
        return this._getJsonRes(400, { error: "Invalid HTTP Verb" });
    }
  }

  private async _handleHttpGet(context: Context, req: HttpRequest): Promise<Response> {
    try {
      const fileInfos: IFileInfo[] = await this._storageService.GetBlobList(config.AzureStorageBlobContainer);

      context.log(`Found ${fileInfos.length} blobs`);
      
      return this._getJsonRes(200, { 
        message: `Found ${fileInfos.length} blobs`,  
        data: fileInfos
      });   
    }
    catch (error) {
      context.log(`ERROR: Exception occured in GET request: ${error}`);
      return this._getJsonRes(400, { error });
    }    
  }


  private async _handleHttpPost(context: Context, req: HttpRequest): Promise<Response> {
    const blobName = req.body.blobName;
    
    const client = df.getClient(context);
    
    const instanceId = await client.startNew("FileOrchestrator", undefined, { blobName });
    
    context.log(`Started orchestration with ID = '${instanceId}'.`);

    return this._getJsonRes(200, {data: client.createCheckStatusResponse(context.bindingData.req, instanceId)});
    
    
    // const blob = await this._storageService.getBlobContent(config.AzureStorageBlobContainer, blobName);

    // console.log('Blob Name', blobName);
    // return this._getJsonRes(200, { message: blobName });
  }

  private _getJsonRes = (status: number, body: IFileManagerResponse): any => {
    return {
      status,
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    }
  }
}

export default FileManagerFunction;