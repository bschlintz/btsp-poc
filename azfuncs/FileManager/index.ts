import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import config from "../shared/config";
import StorageService from "../shared/services/StorageService";
import { IFileInfo } from "../shared/models/IFileInfo";
import { JsonResponse } from '../shared/services/Utils';

let storageService: StorageService;

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  context.log('[FileManager] New Request');

  //Persistent across function invocations
  if (!storageService) {
    storageService = new StorageService(config.AzureStorageBlobAccount, config.AzureStorageBlobKey);
  }
  
  let response;
  try {
    const fileInfos: IFileInfo[] = await storageService.GetBlobList(config.AzureStorageBlobContainer);

    response = JsonResponse(200, { 
      message: `Found ${fileInfos.length} blobs`,  
      data: fileInfos
    });   
  }
  catch (error) {
    context.log(`[FileManager] Error Occurred`, error);
    response = JsonResponse(400, { 
      error: error.message
    });
  }   

  return response;
};

export default httpTrigger;