import { TaskPane } from "./TaskPane";
import { AppContainer } from "react-hot-loader";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { addToLog } from "./logState";

/* global document, Word, Office, OfficeExtension, require,  console, module*/

const render = (Component: React.FC) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

const now = () => new Date().toISOString().replace(/^.*?T/, "").replace(/\..*/, "");

export async function selectionChanged(_event: Office.DocumentSelectionChangedEventArgs) {
  addToLog(`${now()}: selection changed`);
}

Office.onReady(() => {
  OfficeExtension.config.extendedErrorLogging = true;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionChanged, () => {});
  render(TaskPane);
});

if ((module as any).hot) {
  (module as any).hot.accept("./TaskPane", () => {
    const NextTaskPane = require("./TaskPane").default;
    render(NextTaskPane);
  });
}
