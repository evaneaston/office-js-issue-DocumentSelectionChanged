import { hookstate, useHookstate } from "@hookstate/core";

export const logState = hookstate<string[]>([]);

export const useLogState = () => {
  return useHookstate(logState);
};

export const addToLog = (message: string) => {
  logState.set((e) => {
    e.unshift(message);
    return e.slice(0, 20);
  });
};
