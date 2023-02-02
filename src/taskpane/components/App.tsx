import * as React from "react";
import Progress from "./Progress";
import { ToastContainer } from "react-toastify";
import Button from "react-bootstrap/Button";
import Modal from "react-bootstrap/Modal";

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
  isMeetingTitleDirty: boolean;
  isScheduleDateDirty: boolean;
  isScheduleTimeDirty: boolean;
  filePath: string;
}

export enum MEETING {
  MEETING_TITLE = "meetingTitle",
  SCHEDULE_DATE = "scheduleDate",
  SCHEDULE_TIME = "scheduleTime",
}

const App = (props: IAppProps) => {
  const { isOfficeInitialized } = props;
  const [fields, SetFields] = React.useState<IFields>({
    meetingTitle: "",
    meetingVenue: "",
    scheduleDate: "",
    scheduleTime: "",
    meetingDescription: "",
    meetingNotes: "",
    includeDocument: 1,
    videoConferencing: 0,
    isMeetingTitleDirty: false,
    isScheduleDateDirty: false,
    isScheduleTimeDirty: false,
    filePath: "",
  });
  const [show, setShow] = React.useState(false);

  React.useEffect(() => {
    getFileInfo();
  }, []);

  const getFileInfo = async () => {
    Office.context.document.getFilePropertiesAsync((asyncResult) => {
      const filePath = asyncResult.value.url;
      var fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
      SetFields({
        ...fields,
        meetingTitle: trimExtension(fileName),
        filePath: filePath,
      });
    });
  };

  const handleShow = () => setShow(true);
  const handleClose = () => setShow(false);

  const onInputChange = (e: any) => {
    const value = e.target.value;

    if (e.target.name === MEETING.MEETING_TITLE) {
      SetFields({
        ...fields,
        isMeetingTitleDirty: true,
        [e.target.name]: value,
      });
    } else if (e.target.name === MEETING.SCHEDULE_DATE) {
      SetFields({
        ...fields,
        isScheduleDateDirty: true,
        [e.target.name]: value,
      });
    } else if (e.target.name === MEETING.SCHEDULE_TIME) {
      SetFields({
        ...fields,
        isScheduleTimeDirty: true,
        [e.target.name]: value,
      });
    } else {
      SetFields({
        ...fields,
        [e.target.name]: value,
      });
    }
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

  const onScheduleNewMeeting = () => {
    let _fields = {
      isMeetingTitleDirty: false,
      isScheduleDateDirty: false,
      isScheduleTimeDirty: false,
      isDirty: false,
    };
    if (fields?.meetingTitle == "" || !fields.isMeetingTitleDirty) {
      _fields.isMeetingTitleDirty = true;
      _fields.isDirty = true;
    }
    if (fields?.scheduleDate == "" || !fields.isScheduleDateDirty) {
      _fields.isScheduleDateDirty = true;
      _fields.isDirty = true;
    }
    if (fields?.scheduleTime == "" || !fields.isScheduleTimeDirty) {
      _fields.isScheduleTimeDirty = true;
      _fields.isDirty = true;
    }

    const value: IFields = {
      ...fields,
      isMeetingTitleDirty: _fields.isMeetingTitleDirty,
      isScheduleDateDirty: _fields.isScheduleDateDirty,
      isScheduleTimeDirty: _fields.isScheduleTimeDirty,
    };

    if (_fields.isDirty) {
      SetFields({
        ...value,
      });
    } else {
      handleShow();
    }
  };

  const onClickContinue = () => {
    const model = {
      meetingTitle: fields.meetingTitle,
      meetingVenue: fields.meetingVenue,
      scheduleDate: fields.scheduleDate,
      scheduleTime: fields.scheduleTime,
      meetingDescription: fields.meetingDescription,
      meetingNotes: fields.meetingNotes,
      includeDocument: fields.includeDocument == 1 ? true : false,
      videoConferencing: fields.videoConferencing == 1 ? true : false,
      filePath: fields.filePath,
      timeZoneOffset: new Date().getTimezoneOffset(),
    };

    console.log(model);
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
      <div className="mb-3 mt-2">
        <img src={require("./../../../assets/icon-16.png")} style={{ height: "45px" }} alt="" />
        <label>
          <b style={{ fontSize: "16px" }}>Convene in Teams</b>
        </label>
      </div>

      <div>
        <form>
          <div className="mb-3">
            <label className="form-label">
              Meeting Title <span className="required">*</span>
            </label>
            <input
              type="text"
              className={`form-control ${fields?.meetingTitle == "" && fields.isMeetingTitleDirty ? "danger" : ""}`}
              name="meetingTitle"
              title="meetingTitle"
              value={fields.meetingTitle}
              onChange={onInputChange}
            />
            {fields?.meetingTitle == "" && fields.isMeetingTitleDirty && (
              <span className="required">
                <i className="fa fa-exclamation-circle" aria-hidden="true"></i> Please enter the Meeting Title
              </span>
            )}
          </div>
          <div className="row mb-3">
            <label className="form-label">
              Schedule <span className="required">*</span>
            </label>
            <div className="col-7">
              <input
                type="date"
                name="scheduleDate"
                title="scheduleDate"
                className={`form-control ${fields?.scheduleDate == "" && fields.isScheduleDateDirty ? "danger" : ""}`}
                value={fields.scheduleDate}
                onChange={onInputChange}
              />
            </div>
            <div className="col-5">
              <input
                name="scheduleTime"
                title="scheduleTime"
                type="time"
                className={`form-control ${fields?.scheduleTime == "" && fields.isScheduleTimeDirty ? "danger" : ""}`}
                value={fields.scheduleTime}
                onChange={onInputChange}
              />
            </div>
            {((fields?.scheduleDate == "" && fields.isScheduleDateDirty) || (fields?.scheduleTime == "" && fields.isScheduleTimeDirty)) && (
              <span className="required">
                <i className="fa fa-exclamation-circle" aria-hidden="true"></i> Please select the Schedule
              </span>
            )}
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
          <div className="row add-left">
            <button
              type="button"
              className="btn btn-primary"
              // data-bs-toggle="modal"
              // data-bs-target="#scheduleModal"
              onClick={onScheduleNewMeeting}
            >
              Schedule New Meeting
            </button>
          </div>
        </form>
      </div>
      <ToastContainer />

      <Modal show={show} onHide={handleClose} centered>
        <Modal.Header closeButton>
          <Modal.Title>Schedule New Meeting</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          <p>
            You will be redirected to the Convene in Teams app to schedule the new meeting with the details provided.
          </p>
        </Modal.Body>
        <Modal.Footer>
          <Button variant="secondary" onClick={handleClose}>
            Cancel
          </Button>
          <Button variant="primary" onClick={onClickContinue}>
            Continue
          </Button>
        </Modal.Footer>
      </Modal>
    </React.Fragment>
  );
};

export default App;
