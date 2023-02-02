import { Client } from "@microsoft/microsoft-graph-client";

export class GraphService {
  private client: Client | undefined = undefined;

  public ensureClient = (auth: any) => {
    if (!this.client) {
      this.client = Client.initWithMiddleware({
        authProvider: auth,
      });
    }
    return this.client;
  };

  public getUser = async (auth: any): Promise<any> => {
    this.ensureClient(auth);
    const user: any = await this.client!.api("/me").select("displayName,mail,mailboxSettings,userPrincipalName").get();
    return user;
  };
}
