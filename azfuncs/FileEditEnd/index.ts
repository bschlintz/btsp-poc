import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import ADALAuthenticationProvider from "../shared/auth/ADALAuthenticationProvider";
import GraphService from "../shared/services/GraphService";
import config from "../shared/config";
import StorageService from "../shared/services/StorageService";
import { JsonResponse } from "../shared/services/Utils";

let graphService: GraphService;
let storageService: StorageService;

const FileEditEnd: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  const itemId: string = req.body && req.body.itemId ? req.body.itemId : '';
  const blobName: string = req.body && req.body.blobName ? req.body.blobName : '';
  const isDiscard: boolean = req.body && req.body.discard === true ? true : false;

  context.log(`[FileEditEnd] New Request. Item ID: ${itemId}. Blob Name: ${blobName}`);

  //Persistent across function invocations
  if (!graphService) {
    const authProvider = new ADALAuthenticationProvider(config.AdalClientId, config.AdalClientSecret, config.AdalTenant, 'https://graph.microsoft.com');
    graphService = new GraphService(authProvider);
  }
  if (!storageService) {
    storageService = new StorageService(config.AzureStorageBlobAccount, config.AzureStorageBlobKey);
  }

  let response;
  try {

    if (!isDiscard) { 
      const fileBuffer = await graphService.DownloadFileFromSite(config.GraphSharePointSiteId, config.GraphSharePointDriveId, itemId);
      
      const uploadResult = await storageService.UploadBlobContent(config.AzureStorageBlobContainer, blobName, fileBuffer);
    }

    const deleteResult = await graphService.DeleteFileFromSite(config.GraphSharePointSiteId, config.GraphSharePointDriveId, itemId);

    response = JsonResponse(200, {
      message: 'Success',
    });    
  }
  catch (error) {
    context.log(`[FileEditEnd] Error Occurred.`, error);    
    response = JsonResponse(400, {
      error: error.message
    });
  }

  return response;
};

export default FileEditEnd;
