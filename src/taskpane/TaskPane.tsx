import * as React from "react";
import { useLogState } from "./logState";

export const TaskPane: React.FC = () => {
  const logState = useLogState();
  return (
    <div>
      {logState.get().map((e, i) => (
        <div key={i}>{e}</div>
      ))}
    </div>
  );
};
