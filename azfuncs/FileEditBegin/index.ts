import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import ADALAuthenticationProvider from "../shared/auth/ADALAuthenticationProvider";
import GraphService from "../shared/services/GraphService";
import config from "../shared/config";
import { JsonResponse } from "../shared/services/Utils";

let graphService: GraphService;

const FileEditBegin: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
  const inputBlob = context.bindings.inputBlob;
  const blobName: string = req.body && req.body.blobName ? req.body.blobName : '';

  context.log(`[FileEditBegin] New Request. Blob Name: ${blobName}. Blob Bytes: ${inputBlob.length}.`);

  if (!graphService) {
    const authProvider = new ADALAuthenticationProvider(config.AdalClientId, config.AdalClientSecret, config.AdalTenant, 'https://graph.microsoft.com');
    graphService = new GraphService(authProvider);
  }

  let response;
  try {    
    const timestamp = new Date().toISOString().replace(/[-.:'TZ]/g, '');
    const fileName = blobName.substr(blobName.lastIndexOf('/') + 1);
    const extension = fileName.substr(fileName.lastIndexOf('.'));
    const uniqueFileName = `${fileName.replace(extension, '')}${timestamp}${extension}`;

    const uploadResult = await graphService.UploadFileToSite(config.GraphSharePointSiteId, config.GraphSharePointDriveId, '', uniqueFileName, inputBlob);

    response = JsonResponse(200, {
      message: `Success`,
      data: { 
        webUrl: uploadResult.webUrl,
        itemId: uploadResult.id
      }
    });
  }
  catch (error) {
    context.log(`[FileEditBegin] Error Occurred.`, error);
    response = JsonResponse(400, {
      error: error.message
    });
  }

  return response;
};

export default FileEditBegin;
