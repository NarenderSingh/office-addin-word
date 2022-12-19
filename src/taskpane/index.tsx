import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { EventType, PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import config from "./components/Config";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: config.appId,
    redirectUri: config.redirectUri,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true,
  },
});

const accounts = msalInstance.getAllAccounts();
if (accounts && accounts.length > 0) {
  msalInstance.setActiveAccount(accounts[0]);
}

msalInstance.addEventCallback((event: any) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
    const authResult = event.payload;
    msalInstance.setActiveAccount(authResult.account);
  }
});

initializeIcons();
let isOfficeInitialized = false;
const title = "CiT Add-in";

const render = (Component: any) => {
  ReactDOM.render(
    <AppContainer>
      <MsalProvider instance={msalInstance}>
        <ThemeProvider>
          <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </ThemeProvider>
      </MsalProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
