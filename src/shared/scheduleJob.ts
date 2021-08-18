import { tokenServices, updatedStateTable, previousStateTable, compareTableDiffArray } from "../taskpane/taskpane";
import { saveTableDataAPI } from "../service/taskpane-api.service";

var returnRowNumber: string;
export var scheduleJobInterval: any;

const job = async () => {
  returnRowNumber = null;
  await updatedStateTable();
  try {
    if (compareTableDiffArray.length > 0) {
      await submitUpdatedArray();
    }
  } catch (e) {
    console.error(e);
  }
};

export const selectScheduleJob = async (time: number) => {
  clearInterval(scheduleJobInterval);
  scheduleJobInterval = setInterval(() => {
    job();
  }, time * 1000);
};

export const submitUpdatedArray = async () => {
  let now = new Date();
  let hr = now.getHours();
  let mins = now.getMinutes();
  let sec = now.getSeconds();

  returnRowNumber = await saveTableDataAPI(tokenServices, compareTableDiffArray);
  document.getElementById("spanLastPush").textContent = `Pushed ${returnRowNumber} Prices at ${hr}:${mins}:${sec}`;
  compareTableDiffArray.length = 0;
  previousStateTable();
};
