import { IGraphService } from './IGraphService';
import { Client as GraphClient, AuthenticationProvider, PageIterator, PageIteratorCallback } from "@microsoft/microsoft-graph-client";

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

  // public async getGroup(groupId: string): Promise<IGroup> {
  //   try {
  //     const selectProperties = ['id', 'createdDateTime', 'displayName', 'mailNickname', 'visibility', 'classification']
  //     const response = await this._graphClient.api(`/groups/${groupId}`).version('v1.0').select(selectProperties).get();
      
  //     if (response) {
  //       this._writeLog(`INFO: Found group ID ${groupId}`);
  //       const group: IGroup = this._convertGroupResponseToGroup(response);
  //       return group;
  //     }
  //     else {
  //       this._writeLog(`WARNING: Unable to find group ID ${groupId}`);
  //       return null;
  //     }
  //   }
  //   catch (error) {
  //     throw new Error(`ERROR: getGroup - Exception occured when calling the Microsoft Graph API. Message: ${error}`);
  //   }
  // }  
  
}

export default GraphService;