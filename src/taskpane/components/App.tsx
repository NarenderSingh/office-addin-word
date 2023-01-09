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

export interface IAppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface IFields {
  meetingTitle: string;
  meetingDescription: string;
  meetingNotes: string;
  videoConferencing: number;
}

const App = (props: IAppProps) => {
  const { title, isOfficeInitialized } = props;
  const [currentUser, setCurrentUser] = React.useState<any>({});
  const [documentPath, setDocumentPath] = React.useState<any>(null);
  const [blocking, setBlocking] = React.useState<boolean>(true);
  const [fields, SetFields] = React.useState<IFields>({
    meetingTitle: "",
    meetingDescription: "",
    meetingNotes: "",
    videoConferencing: 0,
  });

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
        SetFields({
          ...fields,
          meetingTitle: trimExtension(fileName),
        });
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

  const onInputChange = (e: any) => {
    const value = e.target.value;
    SetFields({
      ...fields,
      [e.target.name]: value,
    });
  };

  const onCheckboxChange = (e: any) => {
    const value = e.target.checked;
    SetFields({
      ...fields,
      [e.target.name]: value,
    });
  };

  const trimExtension = (filename: string) => {
    return filename.substring(0, filename.lastIndexOf(".")) || filename;
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
        <div className="pt-4">
          <h6>
            <b>Hello {currentUser?.displayName}</b>
          </h6>
        </div>
        <hr />
        <p className="mt-2">Please enter the details to create CiT Meeting</p>
        <div>
          <form>
            {/* <div className="mb-3">
              <label className="form-label">Workspace</label>
              <select className="form-select">
                <option value="0">Select Workspaces</option>
                <option value="1">Azeus</option>
              </select>
            </div>
            <div className="mb-3">
              <label className="form-label">Meeting Type</label>
              <select className="form-select">
                <option value="0">Board</option>
                <option value="1">Companey</option>
              </select>
            </div> */}
            <div className="mb-3">
              <label className="form-label">Meeting Title</label>
              <input
                type="text"
                className="form-control"
                name="meetingTitle"
                value={fields.meetingTitle}
                onChange={onInputChange}
              />
              {/* <div className="form-text">Meeting description</div> */}
            </div>
            <div className="mb-3">
              <label className="form-label">Meeting Description</label>
              <textarea
                name="meetingDescription"
                className="form-control"
                cols={30}
                rows={2}
                value={fields.meetingDescription}
                onChange={onInputChange}
              ></textarea>
            </div>
            <div className="mb-3">
              <label className="form-label">Meeting Notes for Participants</label>
              <textarea
                name="meetingNotes"
                className="form-control"
                cols={30}
                rows={2}
                value={fields.meetingNotes}
                onChange={onInputChange}
              ></textarea>
            </div>
            <div className="mb-3">
              <label className="form-label">Video Conferencing</label>
              <div className="input-text">
                <input
                  className="form-check-input m-1"
                  name="videoConferencing"
                  type="checkbox"
                  value={fields.videoConferencing}
                  onChange={onCheckboxChange}
                />
                Schedule a video conference meeting in Teams
              </div>
            </div>
            <div className="text-center">
              <button type="button" className="btn btn-primary" data-bs-toggle="modal" data-bs-target="#exampleModal">
                Create Meeting in CiT Admin with this Document
              </button>
            </div>
          </form>
        </div>
        <br />

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
                  By clicking Create Meeting option, a new meeting in CiT Admin portal will be created with this
                  document as attachment in agenda & with current date and time as schedule. However, you can edit the
                  details at CiT Admin portal.
                </p>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-bs-dismiss="modal">
                  Cancel
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
            Sign In
          </DefaultButton>
        </div>
      </UnauthenticatedTemplate>
      <ToastContainer />
    </React.Fragment>
  );
};

export default App;
