import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import { AuthenticatedTemplate, useMsal } from "@azure/msal-react";
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
    const result = await msal.instance.loginPopup({
      scopes: config.scopes,
      prompt: "select_account",
    });

    console.log("Result", result);
    const user = await graphService.getUser(authProvider);
    console.log("user", user);
    setCurrentUser(user);
  };

  const signOut = async () => {
    await msal.instance.logoutPopup();
  };

  const click = async () => {
    signIn();
    return Word.run(async (context) => {
      const document = context.document.body;
      context.load(document, ["*"]);
      context
        .sync()
        .then(context.sync)
        .then(() => {
          console.log("document", document);
        })
        .then(context.sync);

      Office.context.document.getFilePropertiesAsync((asyncResult) => {
        const fileUrl = asyncResult.value.url;
        setDocumentPath(fileUrl);

        var fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
        setDocumentName(fileName);
      });

      // console.log(Office.context.host.toString());
      // console.log(Office.context.contentLanguage.toString());
      // console.log(Office.context.document.url);
      await context.sync();
    });
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
        {/* <button type="button" className="btn btn-secondary btn-sm" onClick={signOut}>
          Sign Out
        </button> */}
        <br />
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
      </AuthenticatedTemplate>
      {/* <UnauthenticatedTemplate>
        <button type="button" className="btn btn-warning" onClick={signIn}>
          Sign In
        </button>
      </UnauthenticatedTemplate> */}
      <div className="ms-welcome">
        <DefaultButton className="ms-welcome__action" onClick={click}>
          Get File Info
        </DefaultButton>
      </div>
    </React.Fragment>
  );
};

export default App;
