import * as React from "react";
import Progress from "./Progress";
// import {
//   useMsal
// } from "@azure/msal-react";
// import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
// import config from "./Config";
// import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
// import { GraphService } from "./GraphService";
// import { toast, ToastContainer } from "react-toastify";

// const graphService = new GraphService();

export interface IAppProps {
  isOfficeInitialized: boolean;
}

export interface IFields {
  meetingTitle: string;
  meetingVenue: string;
  scheduleDate: string;
  scheduleTime: string;
  meetingDescription: string;
  meetingNotes: string;
  includeDocument: number;
  videoConferencing: number;
}

const App = (props: IAppProps) => {
  const { isOfficeInitialized } = props;
  // const [currentUser, setCurrentUser] = React.useState<any>({});
  // const [documentPath, setDocumentPath] = React.useState<any>(null);
  const [fields, SetFields] = React.useState<IFields>({
    meetingTitle: "",
    meetingVenue: "",
    scheduleDate: "",
    scheduleTime: "",
    meetingDescription: "",
    meetingNotes: "",
    includeDocument: 1,
    videoConferencing: 0,
  });

  // const msal = useMsal();
  // const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(
  //   msal.instance as PublicClientApplication,
  //   {
  //     account: msal.instance.getActiveAccount()!,
  //     scopes: config.scopes,
  //     interactionType: InteractionType.Popup,
  //   }
  // );

  React.useEffect(() => {
    // authProvider.getAccessToken().then(() => {
    //   getFileInfo();
    // });
    getFileInfo();
  }, []);

  // const getUserInfo = () => {
  //   graphService.getUser(authProvider).then((user: any) => {
  //     setCurrentUser(user);
  //     getFileInfo();
  //   });
  // };

  const getFileInfo = async () => {
    Office.context.document.getFilePropertiesAsync((asyncResult) => {
      const fileUrl = asyncResult.value.url;
      var fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
      SetFields({
        ...fields,
        meetingTitle: trimExtension(fileName),
      });
    });
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
          title={"Loading..."}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      )}
      <img src={require("./../../../assets/icon-16.png")} style={{ height: "45px" }} alt="" />
      <label>
        <b style={{ fontSize: "16px" }}>Convene in Teams</b>
      </label>
      <hr />
      <div>
        <form>
          <div className="mb-3">
            <label className="form-label">Meeting Title</label>
            <input
              type="text"
              className="form-control hello"
              name="meetingTitle"
              title="meetingTitle"
              value={fields.meetingTitle}
              onChange={onInputChange}
            />
          </div>
          <div className="row mb-3">
            <label className="form-label">Schedule</label>
            <div className="col-7">
              <input
                type="date"
                name="scheduleDate"
                title="scheduleDate"
                className="form-control"
                value={fields.scheduleDate}
                onChange={onInputChange}
              />
            </div>
            <div className="col-5">
              <input
                name="scheduleTime"
                title="scheduleTime"
                type="time"
                className="form-control"
                value={fields.scheduleTime}
                onChange={onInputChange}
              />
            </div>
          </div>
          <div className="mb-3">
            <label className="form-label">Venue</label>
            <input
              name="meetingVenue"
              title="meetingVenue"
              type="text"
              className="form-control"
              value={fields.meetingVenue}
              onChange={onInputChange}
            />
          </div>
          <div className="mb-3">
            <label className="form-label">Description</label>
            <textarea
              name="meetingDescription"
              title="meetingDescription"
              className="form-control"
              cols={30}
              rows={2}
              value={fields.meetingDescription}
              onChange={onInputChange}
            ></textarea>
          </div>
          <div className="mb-3">
            <label className="form-label">Notes for Participants</label>
            <textarea
              title="meetingNotes"
              name="meetingNotes"
              className="form-control"
              cols={30}
              rows={2}
              value={fields.meetingNotes}
              onChange={onInputChange}
            ></textarea>
          </div>
          <div className="mb-3">
            <div className="input-text">
              <input
                name="includeDocument"
                title="includeDocument"
                className="form-check-input m-1"
                type="checkbox"
                value={fields.includeDocument}
                defaultChecked={true}
                onChange={onCheckboxChange}
              />
              Include this document in the meeting
            </div>
          </div>
          <div className="mb-3">
            <div className="input-text">
              <input
                name="videoConferencing"
                title="meetingNotes"
                className="form-check-input m-1"
                type="checkbox"
                value={fields.videoConferencing}
                onChange={onCheckboxChange}
              />
              Schedule a video conference meeting in Teams
            </div>
          </div>
          <div className="text-center">
            <button type="button" className="btn btn-primary" data-bs-toggle="modal" data-bs-target="#exampleModal">
              Schedule New Meeting
            </button>
          </div>
        </form>
      </div>
      {/* <ToastContainer /> */}

      <div
        className="modal fade"
        id="exampleModal"
        tabIndex={-1}
        aria-labelledby="exampleModalLabel"
        aria-hidden="true"
      >
        <div className="modal-dialog modal-dialog-centered">
          <div className="modal-content">
            <div className="modal-header">
              <h1 className="modal-title fs-5" id="exampleModalLabel">
                Disclaimer
              </h1>
              <button type="button" className="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div className="modal-body">
              <p>
                This is to inform that by clicking on Confirm, you will be redirected to the Convene in Teams app to
                schedule a new meeting with existing details.
              </p>
            </div>
            <div className="modal-footer">
              <button type="button" className="btn btn-secondary" data-bs-dismiss="modal">
                Cancel
              </button>
              <button type="button" className="btn btn-primary" data-bs-dismiss="modal">
                Confirm
              </button>
            </div>
          </div>
        </div>
      </div>
    </React.Fragment>
  );
};

export default App;
