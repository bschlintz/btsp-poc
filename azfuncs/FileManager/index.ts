import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import FileManagerFunction from "./FileManagerFunction";
import config from "../shared/config";
import StorageService from "../shared/services/StorageService";
import { IStorageService } from "../shared/services/IStorageService";

let storageService: IStorageService;

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  context.log('HTTP trigger function processed a request.');

  //Persistent across function invocations
  if (!storageService) {
    storageService = new StorageService(config.AzureStorageBlobAccount, config.AzureStorageBlobKey);
  }
  
  const groupSiteFunction = new FileManagerFunction(storageService, null);
  const result = await groupSiteFunction.processRequest(context, req);
  context.log(`HTTP trigger function finished processing a request. Status Code: ${result.status}`);

  return result;
};

export default httpTrigger;