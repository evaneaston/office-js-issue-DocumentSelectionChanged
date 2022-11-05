import { TaskPane } from "./TaskPane";
import { AppContainer } from "react-hot-loader";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { addToLog } from "./logState";

/* global document, Word, Office, OfficeExtension, require,  console, module*/

export async function selectionChanged(event: Office.DocumentSelectionChangedEventArgs) {
  console.log("selection changed", event);
  addToLog(`selection changed @ ${new Date().toISOString()}`);
}

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
