import * as React from "react";
import { useLogState, clearLogState } from "./logState";

/* global Word, console, setTimeout */

const repro1 = async () => {
  Word.run(async (context: Word.RequestContext) => {
    try {
      context.document.body.clear();
      context.sync();

      const textRange = context.document.body.insertText("A", Word.InsertLocation.start);
      const inlinePicture = context.document.body.insertInlinePictureFromBase64(
        "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",
        Word.InsertLocation.end
      );
      textRange.select(Word.SelectionMode.end);

      await context.sync();
      setTimeout(() => clearLogState(), 250);
    } catch (e) {
      console.error("Error during repro1", e);
    }
  });
};

const Observation: React.FC<React.PropsWithChildren<{ bad?: boolean }>> = ({ bad, children }) => (
  <div style={{ marginLeft: 4, fontSize: "90%", fontStyle: "italic", color: bad ? "red" : "auto" }}>{children}</div>
);

const ResetLink: React.FC = () => {
  return (
    <>
      Click{" "}
      <a href="#" onClick={repro1}>
        here
      </a>{" "}
      to reset document and clear log (below)
    </>
  );
};

export const TaskPane: React.FC = () => {
  const logState = useLogState();

  return (
    <>
      <div style={{ padding: "4px 4px", display: "flex", flexDirection: "column" }}>
        <div>
          <b>Reproduction #1</b>
        </div>
        <ol>
          <li>
            <ResetLink />
            <Observation>The selection will be between the A and the icon.</Observation>
          </li>
          <li>
            Click <b>Clear</b> on log below if it's not empty.
            <Observation>Selection changed log is empty</Observation>
          </li>
          <li>
            Click on logo
            <Observation>See icon get selected</Observation>
            <Observation bad>No new event logged</Observation>
          </li>
        </ol>

        <div>
          <b>Reproduction #2</b>
        </div>
        <ol>
          <li>
            <ResetLink />
            <Observation>The selection will be between the A and the icon.</Observation>
          </li>
          <li>Click document body</li>
          <li>
            Press Home <Observation>Selection changed event logged</Observation>
          </li>
          <li>
            Press Right Arrow
            <Observation>Selection changed event logged</Observation>
            <Observation>Caret/selection is between the A and the icon</Observation>
          </li>
          <li>
            Press <b>Shift+Right Arrow</b>
            <Observation>See icon get selected</Observation>
            <Observation bad>No new event logged</Observation>
          </li>
          <li>
            Press <b>Shift+Right Arrow</b> again
            <Observation>Selection changed event logged</Observation>
          </li>
        </ol>
      </div>

      <div>
        <b>Selection Changed Events Log</b> (
        <a href="#" onClick={clearLogState}>
          Clear
        </a>
        )
      </div>
      <div>
        {logState.get().map((e, i) => (
          <div key={i}>{e}</div>
        ))}
      </div>
    </>
  );
};
