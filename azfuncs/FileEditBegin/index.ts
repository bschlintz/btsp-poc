import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { IGraphService } from "../shared/services/IGraphService";
import ADALAuthenticationProvider from "../shared/auth/ADALAuthenticationProvider";
import GraphService from "../shared/services/GraphService";
import config from "../shared/config";
import { v4 as uuidv4 } from 'uuid';

let adminGraphService: IGraphService;

const FileEditBegin: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  const inputBlob = context.bindings.inputBlob;
  const blobName: string = req.body && req.body.blobName ? req.body.blobName : '';

  context.log("[FileEditBegin] New Request \n Name:", blobName, "\n Blob Size:", inputBlob.length, "Bytes");

  if (!adminGraphService) {
    const authProvider = new ADALAuthenticationProvider(config.AdalClientId, config.AdalClientSecret, config.AdalTenant, 'https://graph.microsoft.com');
    adminGraphService = new GraphService(authProvider, null, context.log);
  }

  try {
    const uniqueId = uuidv4();
    const fileName = blobName.substr(blobName.lastIndexOf('/') + 1);
    const result = await adminGraphService.UploadFileToSite(config.GraphSharePointSiteId, config.GraphSharePointDriveId, uniqueId, fileName, inputBlob);
    context.log("[FileEditBegin] Uploaded File to SPO \n Web URL:", result.webUrl);
    return result;
  }
  catch (error) {
    context.log("[FileEditBegin] Error Uploading File to SPO \n Message:", error);
  }
};

export default FileEditBegin;
