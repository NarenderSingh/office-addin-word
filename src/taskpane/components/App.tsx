import * as React from "react";
import Progress from "./Progress";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import config from "./Config";
import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { GraphService } from "./GraphService";

const graphService = new GraphService();

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App = (props: AppProps) => {
  const { title, isOfficeInitialized } = props;
  const [currentUser, setCurrentUser] = React.useState<any>({});
  const [documentPath, setDocumentPath] = React.useState<any>(null);
  const [documentName, setDocumentName] = React.useState<any>(null);

  const msal = useMsal();
  const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(msal.instance as PublicClientApplication, {
    account: msal.instance.getActiveAccount()!,
    scopes: config.scopes,
    interactionType: InteractionType.Popup,
  });

  const signIn = async () => {
    await msal.instance.loginPopup({
      scopes: config.scopes,
      prompt: "select_account",
    });
    const user = await graphService.getUser(authProvider);
    setCurrentUser(user);
    getFileInfo();
  };

  const getFileInfo = async () => {
    return Word.run(async (context) => {
      Office.context.document.getFilePropertiesAsync((asyncResult) => {
        const fileUrl = asyncResult.value.url;
        setDocumentPath(fileUrl);

        var fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
        setDocumentName(fileName);
      });
      await context.sync();
    });
  };

  const createMeeting = () => {
    getFileInfo();
    console.log(currentUser, documentName);
  };

  return (
    <React.Fragment>
      {!isOfficeInitialized && (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      )}
      <img src={require("./../../../assets/icon-16.png")} style={{ height: "50px" }} alt="" />
      <label>
        <b style={{ fontSize: "18px" }}>Convene in Teams</b>
      </label>
      <br />
      <br />
      <AuthenticatedTemplate>
        <div>
          <h5>
            <b>Hello {currentUser?.displayName}</b>
          </h5>
        </div>
        <table className="table table-border">
          <thead>
            <tr>
              <th>Name</th>
              <th>Value</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <b>Id</b>
              </td>
              <td>{currentUser?.id}</td>
            </tr>
            <tr>
              <td>
                <b>Email Address</b>
              </td>
              <td>{currentUser?.mail}</td>
            </tr>
            <tr>
              <td>
                <b>Display Name</b>
              </td>
              <td>{currentUser?.displayName}</td>
            </tr>
            <tr>
              <td>
                <b>Time Zone</b>
              </td>
              <td>{currentUser?.mailboxSettings?.timeZone}</td>
            </tr>
            <tr>
              <td>
                <b>Language</b>
              </td>
              <td>{currentUser?.mailboxSettings?.language?.displayName}</td>
            </tr>
            <tr>
              <td>
                <b>File URL</b>
              </td>
              <td>{documentPath}</td>
            </tr>
            <tr>
              <td>
                <b>File Name</b>
              </td>
              <td>{documentName}</td>
            </tr>
          </tbody>
        </table>
        <br />
        <button type="button" className="ms-welcome__action" onClick={createMeeting}>
          Create Meeting in CiT Admin with this Document
        </button>
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <button type="button" className="ms-welcome__action" onClick={signIn}>
          Sign In
        </button>
      </UnauthenticatedTemplate>
    </React.Fragment>
  );
};

export default App;
