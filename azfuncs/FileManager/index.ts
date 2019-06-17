import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import FileManagerFunction from "./FileManagerFunction";
import config from "../shared/config";
import ADALAuthenticationProvider from "../shared/auth/ADALAuthenticationProvider";
import GraphService from "../shared/services/GraphService";
import StorageService from "../shared/services/StorageService";
import { getTokenFromAuthorizationHeader, getUpnFromToken } from "../shared/auth/TokenHelper";
import { IStorageService } from "../shared/services/IStorageService";
import { IGraphService } from "../shared/services/IGraphService";

let storageService: IStorageService;
let adminGraphService: IGraphService;

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  context.log('HTTP trigger function processed a request.');

  //Persistent across function invocations
  if (!storageService) {
    storageService = new StorageService(config.AzureStorageBlobAccount, config.AzureStorageBlobKey);
  }

  // storageService

  //Persistent across function invocations
  // if (!adminGraphService) {
  //   const authProvider = new ADALAuthenticationProvider(config.AdalClientId, config.AdalClientSecret, config.AdalTenant, 'https://graph.microsoft.com');
  //   adminGraphService = new GraphService(authProvider, null, context.log);
  // }
  
  const groupSiteFunction = new FileManagerFunction(storageService, adminGraphService);
  const result = await groupSiteFunction.processRequest(context, req);
  context.log(`HTTP trigger function finished processing a request. Status Code: ${result.status}`);

  return result;
};

export default httpTrigger;
