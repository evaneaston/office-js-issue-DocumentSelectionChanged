import * as React from "react";
import { useLogState, clearLogState } from "./logState";

export const TaskPane: React.FC = () => {
  const logState = useLogState();
  return (
    <div onClick={clearLogState}>
      {logState.get().map((e, i) => (
        <div key={i}>{e}</div>
      ))}
    </div>
  );
};
