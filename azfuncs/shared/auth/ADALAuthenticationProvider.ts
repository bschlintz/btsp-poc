import { AuthenticationContext } from 'adal-node';
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

class ADALAuthenticationProvider implements AuthenticationProvider {
  private _clientId: string;
  private _clientSecret: string;
  private _resource: string;
  private _authContext: AuthenticationContext;

  constructor(clientId: string, clientSecret: string, tenant: string, resource: string) {
    this._clientId = clientId;
    this._clientSecret = clientSecret;
    this._resource = resource;
    this._authContext = new AuthenticationContext(`https://login.microsoftonline.com/${tenant}`);
  }
  
  public async getAccessToken(): Promise<any>  {
    return new Promise((resolve, reject) => {
      this._authContext.acquireTokenWithClientCredentials(this._resource, this._clientId, this._clientSecret, (err, tokenRes: any) => {
        if (err) { reject(err); }
        resolve(tokenRes.accessToken);
      });
    });    
  }
}

export default ADALAuthenticationProvider;