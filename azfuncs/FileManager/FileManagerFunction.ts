import { Context, HttpRequest } from "@azure/functions";
import { IStorageService } from "../shared/services/IStorageService";
import { IGraphService } from "../shared/services/IGraphService";
import config from "../shared/config";
import { IFileManagerResponse } from "./IFileManagerResponse";
import { IFileInfo } from "../shared/models/IFileInfo";

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
      // case "POST": {
      //   const group = req.body && req.body as IGroup;
      //   if (!group || !group.customAttributes) {
      //     return this._getJsonRes(400, { error: "You must provide a POST body when using this endpoint with HTTP POST. Example POST body: {'id': '<GROUP ID>', 'customAttributes': [{'internalName': 'department', 'value': 'Information Technology'}]}" })
      //   }
      //   return await this._handleHttpPost(context, req);
      // }
  
      default:
        return this._getJsonRes(400, { error: "Invalid HTTP Verb" });
    }
  }

  private async _handleHttpGet(context: Context, req: HttpRequest): Promise<Response> {
    try {
      const fileInfos: IFileInfo[] = await this._storageService.getFileList(config.AzureStorageBlobContainer);

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
    return null;
    // try {
    //   const group = req.body && req.body as IGroup;

    //   const isOwnerOfGroup = await this._adminGraphService.isUserOwnerOfGroup(this._callerUpn, group.id);
      
    //   if (isOwnerOfGroup) {
    //     const updatedGroup = await this._storageService.updateGroupCustomAttributes(group.id, group.customAttributes);
    //     return this._getJsonRes(201, { group: updatedGroup });
    //   }
    //   else {
    //     return this._getJsonRes(401, { error: `User not authorized to update group details`});
    //   }
    // }
    // catch (error) {
    //   context.log(`ERROR: Exception occured in POST request: ${error}`);
    //   return this._getJsonRes(400, { error });
    // }
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