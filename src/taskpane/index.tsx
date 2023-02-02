import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();
let isOfficeInitialized = false;

const render = (Component: any) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("root-office-app")
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
