import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { IGraphService } from "../shared/services/IGraphService";
import ADALAuthenticationProvider from "../shared/auth/ADALAuthenticationProvider";
import GraphService from "../shared/services/GraphService";
import config from "../shared/config";
import { IStorageService } from "../shared/services/IStorageService";
import StorageService from "../shared/services/StorageService";

let adminGraphService: IGraphService;
let storageService: IStorageService;

const FileEditEnd: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  const itemId: string = req.body && req.body.itemId ? req.body.itemId : '';
  const blobName: string = req.body && req.body.blobName ? req.body.blobName : '';

  context.log("[FileEditEnd] New Request \n File ID:", itemId, "\n Blob Name:", blobName);

  if (!adminGraphService) {
    const authProvider = new ADALAuthenticationProvider(config.AdalClientId, config.AdalClientSecret, config.AdalTenant, 'https://graph.microsoft.com');
    adminGraphService = new GraphService(authProvider, null, context.log);
  }

  //Persistent across function invocations
  if (!storageService) {
    storageService = new StorageService(config.AzureStorageBlobAccount, config.AzureStorageBlobKey);
  }

  try {
    const downloadResult = await adminGraphService.DownloadFileFromSite(config.GraphSharePointSiteId, config.GraphSharePointDriveId, itemId);
    context.log("[FileEditEnd] Downloaded File from SPO \n File URL:"/*, result*/);

    const timestamp = new Date().toISOString();
    const extension = blobName.substr(blobName.lastIndexOf('.'));
    const uniqueBlobName = `${blobName.replace(extension, '')}-${timestamp}${extension}`;
    
    const uploadResult = await storageService.UploadBlobContent(config.AzureStorageBlobContainer, uniqueBlobName, downloadResult);
    context.log("[FileEditEnd] Uploaded File to Azure Blob \n File URL:"/*, result*/);
    return uploadResult;
  }
  catch (error) {
    context.log("[FileEditEnd] Error Downloading File from SPO \n Message:", error);
  }
};

export default FileEditEnd;
