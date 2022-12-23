import * as React from "react";
import Progress from "./Progress";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import config from "./Config";
import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { GraphService } from "./GraphService";
import { toast, ToastContainer } from "react-toastify";
import { DefaultButton } from "@fluentui/react";

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

  React.useEffect(() => {
    authProvider.getAccessToken().then(() => {
      getUserInfo();
    });
  }, []);

  const signIn = async () => {
    await msal.instance.loginPopup({
      scopes: config.scopes,
      prompt: "select_account",
    });
    getUserInfo();
  };

  const getUserInfo = () => {
    graphService.getUser(authProvider).then((user: any) => {
      setCurrentUser(user);
      getFileInfo();
    });
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

  const createCiTMeeting = () => {
    getFileInfo();
    setTimeout(() => {
      toast.success("CiT Meeting created successfully!");
      // toast.error("An Error occured");
    }, 500);
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
      <img src={require("./../../../assets/icon-16.png")} style={{ height: "45px" }} alt="" />
      <label>
        <b style={{ fontSize: "16px" }}>Convene in Teams</b>
      </label>
      <br />
      <AuthenticatedTemplate>
        <div>
          <h5>
            <b>Hello {currentUser?.displayName}</b>
          </h5>
        </div>
        <table className="table table-border">
          <tbody>
            <tr>
              <td>
                <b>Name</b>
              </td>
              <td>{currentUser?.displayName}</td>
            </tr>
            <tr>
              <td>
                <b>Email</b>
              </td>
              <td>{currentUser?.mail}</td>
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
                <b>File Name</b>
              </td>
              <td>{documentName}</td>
            </tr>
            <tr>
              <td>
                <b>File URL</b>
              </td>
              <td>{documentPath}</td>
            </tr>
          </tbody>
        </table>
        <br />
        <div className="text-center">
          <button type="button" className="btn btn-primary" data-bs-toggle="modal" data-bs-target="#exampleModal">
            Create Meeting in CiT Admin with this Document
          </button>
        </div>

        <div
          className="modal fade"
          id="exampleModal"
          tabIndex={-1}
          aria-labelledby="exampleModalLabel"
          aria-hidden="true"
        >
          <div className="modal-dialog">
            <div className="modal-content">
              <div className="modal-header">
                <h1 className="modal-title fs-5" id="exampleModalLabel">
                  Create Meeting Confirmation
                </h1>
                <button type="button" className="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
              </div>
              <div className="modal-body">
                <p>
                  By clicking Create Meeting button, you will be creating a new meeting in CiT Admin portal with this
                  document as attachment in agenda.
                </p>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-bs-dismiss="modal">
                  Close
                </button>
                <button type="button" className="btn btn-primary" data-bs-dismiss="modal" onClick={createCiTMeeting}>
                  Create Meeting
                </button>
              </div>
            </div>
          </div>
        </div>
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <div className="ms-welcome">
          <DefaultButton className="ms-welcome__action" onClick={signIn}>
            Get File Info
          </DefaultButton>
        </div>
      </UnauthenticatedTemplate>
      <ToastContainer />
    </React.Fragment>
  );
};

export default App;
