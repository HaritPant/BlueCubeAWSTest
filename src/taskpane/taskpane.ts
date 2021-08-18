//***********************************************************************************************************
//   CONFIDENTIAL
//
//   COPYRIGHT 2018 - 2021
//   Enscape Solutions, LLC DBA BlueCube Energy
//   All Rights Reserved
//
//   NOTICE:  All information contained herein is, and remains the property of Enscape Solutions.
//   The intellectual and technical concepts contained herein are proprietary to Enscape Solutions
//   and are protected by trade secret or copyright law. Dissemination of this information or
//   reproduction of this material is strictly prohibited unless prior written permission is obtained
//   from Enscape Solutions.
//
//   Included third party assets:
//   Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License.
//
//   CONFIDENTIAL
//***********************************************************************************************************

import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-64.png";
import "../../assets/icon-80.png";
import "../../assets/icon-128.png";
import * as AWS from "aws-sdk";
import { getGraphData } from "./../helpers/ssoauthhelper";
import { getTableDatafromAPI } from "../service/taskpane-api.service";
import data from "../shared/constant.json";
import { ColumnAPI } from "../model/column-API.model";
import { ColumnName } from "../model/column-name.enum";
import { selectScheduleJob, submitUpdatedArray, scheduleJobInterval } from "../shared/scheduleJob";

var tokenID: string;
var updatedRow: Object;
var usernameValue: string;
var passwordValue: string;
var pushFrequency: number;
var tableData: any;
var previousStateTableArray: any;
var updatedStateTableArray: any;
var outFocusPushFrequencyStoreRange;
var outFocusPushFrequencyStoreHideRange;
var pushFrequencyStoreRange;
export var compareTableDiffArray = [];
export var pushUpdatedArray = [];
export var tokenServices: string;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function () {
      $("#getGraphDataButton").click(getGraphData);
    });
    try {
      document.getElementById("divFetchTemplate").style.display = "none";
      document.getElementById("divMonitorRange").style.display = "none";
      document.getElementById("divRefreshMarket").style.display = "none";
      document.getElementById("spanInvalidMessage").style.visibility = "hidden";

      Excel.run(function (context) {
        let sheets = context.workbook.worksheets;
        let sheet = sheets.add(data.workSheetName);
        sheet.load("name, position");
        sheet.activate();
        return context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
});

export const writeDataToOfficeDocument = async (result: Object): Promise<any> => {
  if (result != null) {
    tokenServices = await fetchAWSAPItoken();
    document.getElementById("divFetchTemplate").style.display = "block";
    document.getElementById("divSignIn").style.display = "none";
  }
};

export const fetchAWSAPItoken = async () => {
  try {
    const cognitoClient = new AWS.CognitoIdentityServiceProvider({
      region: data.region,
    });
    const payload = {
      AuthFlow: data.authFlow,
      ClientId: data.clientId,
      AuthParameters: {
        USERNAME: data.email,
        PASSWORD: data.password,
      },
    };
    const response = await cognitoClient.initiateAuth(payload).promise();
    tokenID = response.AuthenticationResult.IdToken;
    return tokenID;
  } catch (error) {
    console.error(error);
  }
};

export const userLoginFetchAWSAPItoken = async () => {
  try {
    const cognitoClient = new AWS.CognitoIdentityServiceProvider({
      region: data.region,
    });
    const payload = {
      AuthFlow: data.authFlow,
      ClientId: data.clientId,
      AuthParameters: {
        USERNAME: usernameValue,
        PASSWORD: passwordValue,
      },
    };
    const response = await cognitoClient.initiateAuth(payload).promise();
    tokenID = response.AuthenticationResult.IdToken;
    tokenServices = tokenID;

    document.getElementById("divFetchTemplate").style.display = "block";
    document.getElementById("spanInvalidMessage").style.visibility = "hidden";
    document.getElementById("divSignIn").style.display = "none";
    document.getElementById("divMonitorRange").style.visibility = "none";
    return tokenID;
  } catch (error) {
    console.error(error);
    document.getElementById("spanInvalidMessage").style.visibility = "visible";
  }
};

const validateUser = async () => {
  usernameValue = (<HTMLInputElement>document.getElementById("inputUserName")).value;
  passwordValue = (<HTMLInputElement>document.getElementById("inputPassword")).value;
  if (usernameValue === null || passwordValue === null) {
    document.getElementById("spanInvalidMessage").style.visibility = "visible";
  } else {
    document.getElementById("spanInvalidMessage").style.visibility = "hidden";
    await userLoginFetchAWSAPItoken();
  }
};

const validatePushFrequencyOnFocus = async () => {
  document.getElementById("spanPushFrequencyMessage").style.visibility = "visible";
};

const fetchServerDataAPICall = async () => {
  tableData = await getTableDatafromAPI(tokenServices);
};

const outFocusPushFrequencyCache = async () => {
  return new Promise(async (reject) => {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(data.workSheetName);
      outFocusPushFrequencyStoreRange = sheet.getRange(data.hideStoreRange);
      outFocusPushFrequencyStoreRange.values = [[pushFrequency]];
      outFocusPushFrequencyStoreHideRange = context.workbook.worksheets
        .getItem(data.workSheetName)
        .getRange(data.hideRange);
      outFocusPushFrequencyStoreHideRange.columnHidden = true;
      await selectScheduleJob(pushFrequency);
      return context.sync();
    }).catch((createError) => {
      reject(createError);
    });
  });
};

const validateNewPushFrequencyOutFocus = async () => {
  let outFocusPushFreqInputValue = (<HTMLInputElement>document.getElementById("inputPushFrequency")).value;
  let btnRefreshVisbility = document.getElementById("divRefreshMarket").style.visibility;
  let outFocusPushFreqInputNumber = parseInt(outFocusPushFreqInputValue);

  if (outFocusPushFreqInputNumber < 15 || outFocusPushFreqInputNumber > 3600 || isNaN(outFocusPushFreqInputNumber)) {
    (<HTMLInputElement>document.getElementById("inputPushFrequency")).value = "30";
    document.getElementById("spanPushFrequencyMessage").style.visibility = "hidden";

    if (btnRefreshVisbility != "visible") {
      pushFrequency = 30;
      outFocusPushFrequencyCache();
    }
  } else {
    if (btnRefreshVisbility != "visible") {
      pushFrequency = parseInt((<HTMLInputElement>document.getElementById("inputPushFrequency")).value);
      outFocusPushFrequencyCache();
    }
    document.getElementById("spanPushFrequencyMessage").style.visibility = "hidden";
  }
};

const setPushFrequencyCacheHide = async () => {
  Excel.run(function (context) {
    let range = context.workbook.worksheets.getItem(data.workSheetName).getRange(data.hideRange);
    range.columnHidden = true;
    return context.sync();
  });
};

const fetchPushFrequencyCache = async () => {
  return new Promise(async (reject) => {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(data.workSheetName);
      pushFrequencyStoreRange = sheet.getRange(data.hideStoreRange);
      pushFrequencyStoreRange.load("values");

      return context.sync().then(() => {
        if (pushFrequencyStoreRange.values == "") {
          pushFrequency = parseInt((<HTMLInputElement>document.getElementById("inputPushFrequency")).value);
          pushFrequencyStoreRange.values = [[pushFrequency]];
          selectScheduleJob(pushFrequency);
        } else {
          pushFrequency = pushFrequencyStoreRange.values;
          (<HTMLInputElement>document.getElementById("inputPushFrequency")).value = pushFrequencyStoreRange.values;
          selectScheduleJob(pushFrequency);
        }
        setPushFrequencyCacheHide();
      });
    }).catch((createError) => {
      reject(createError);
    });
  });
};

const displayFetchData = async () => {
  return new Promise(async (reject) => {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(data.workSheetName);
      let expensesTable = sheet.tables.add(data.columnRange, true);
      expensesTable.name = data.tableName;
      expensesTable.getHeaderRowRange().values = [
        [
          ColumnName.TP,
          ColumnName.Mkt,
          ColumnName.B,
          ColumnName.BVol,
          ColumnName.A,
          ColumnName.AVol,
          ColumnName.IBAdd,
          ColumnName.IBVol,
          ColumnName.IAAdd,
          ColumnName.IAVol,
          ColumnName.IWap,
        ],
      ];

      await fetchServerDataAPICall();

      let transactions = tableData;
      let newData = await transactions.data.map((item) => [
        item.TP,
        item.Mkt,
        item.B,
        item.BVol,
        item.A,
        item.AVol,
        item.IBAdd,
        item.IBVol,
        item.IAAdd,
        item.IAVol,
        item.IWap,
      ]);

      expensesTable.rows.add(null, newData);
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      sheet.activate();
      previousStateTable();
      return context.sync();
    }).catch((createError) => {
      reject(createError);
    });
  });
};

var refreshModal = document.getElementById("myRefreshModal");
const userConfirmRefreshMarketList = async () => {
  refreshModal.style.display = "block";
};

var fetchModal = document.getElementById("myFetchModal");
const userConfirmFetchMarket = async () => {
  fetchModal.style.display = "block";
};

const notRefreshMarket = async () => {
  refreshModal.style.display = "none";
};

const noFetchServerData = async () => {
  fetchModal.style.display = "none";
  await cancelFetchServerData();
};

const refeshMarketData = async () => {
  refreshModal.style.display = "none";
  Excel.run(function (ctx) {
    let tableName = data.tableName;
    let table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    displayFetchData();
    return ctx.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

export const previousStateTable = async () => {
  return new Promise(async (reject) => {
    await Excel.run(async (context) => {
      let fisrtTableArray = [];
      let sheet = context.workbook.worksheets.getItem(data.workSheetName);
      let expensesTable = sheet.tables.getItem(data.tableName);
      let bodyRange = expensesTable.getDataBodyRange().load("values");

      return context.sync().then(() => {
        fisrtTableArray = bodyRange.values;
        previousStateTableArray = JSON.parse(JSON.stringify(fisrtTableArray));
      });
    }).catch((createError) => {
      reject(createError);
    });
  });
};

export const updatedStateTable = async () => {
  await Excel.run(async (context) => {
    let secondTableArray = [];
    let sheet = context.workbook.worksheets.getItem(data.workSheetName);
    let expensesTable = sheet.tables.getItem(data.tableName);
    let bodyRange = expensesTable.getDataBodyRange().load("values");
    return context.sync().then(() => {
      secondTableArray = bodyRange.values;
      updatedStateTableArray = JSON.parse(JSON.stringify(secondTableArray));
      compareTableDiffArray = [];

      for (let i = 0; i < previousStateTableArray.length; i++) {
        if (JSON.stringify(previousStateTableArray[i]) != JSON.stringify(updatedStateTableArray[i])) {
          let rowValue = updatedStateTableArray[i];
          updatedRow = new ColumnAPI(
            rowValue[0],
            rowValue[1],
            parseInt(rowValue[2]),
            parseInt(rowValue[3]),
            parseInt(rowValue[4]),
            parseInt(rowValue[5]),
            parseInt(rowValue[6]),
            parseInt(rowValue[7]),
            parseInt(rowValue[8]),
            parseInt(rowValue[9]),
            parseInt(rowValue[10])
          );
          compareTableDiffArray.push(updatedRow);
        }
      }
    });
  }).catch((createError) => {
    console.log(createError);
  });
};

const fetchServerData = async () => {
  fetchPushFrequencyCache();
  document.getElementById("divFetchTemplate").style.display = "none";
  document.getElementById("spanPushFrequencyMessage").style.visibility = "hidden";
  document.getElementById("divMonitorRange").style.display = "block";
  await refeshMarketData();
};

const cancelFetchServerData = async () => {
  document.getElementById("divFetchTemplate").style.display = "none";
  document.getElementById("spanPushFrequencyMessage").style.visibility = "hidden";
  document.getElementById("divMonitorRange").style.display = "block";
  fetchPushFrequencyCache();
};

const tooglePauseResume = async () => {
  try {
    if (pauseResume.innerText === data.pause) {
      pauseResume.innerText = data.resume;
      clearInterval(scheduleJobInterval);
      document.getElementById("divRefreshMarket").style.display = "block";
    } else {
      pauseResume.innerText = data.pause;
      await updatedStateTable();
      if (compareTableDiffArray.length > 0) {
        await submitUpdatedArray();
      }
      pushFrequency = parseInt((<HTMLInputElement>document.getElementById("inputPushFrequency")).value);
      await selectScheduleJob(pushFrequency);
      document.getElementById("divRefreshMarket").style.display = "none";
    }
  } catch (error) {
    console.error(error);
  }
};

var userSignIn = document.getElementById("btnSignIn");
userSignIn.addEventListener("click", validateUser);

var fetchTemplate = document.getElementById("btnFetchTemplate");
fetchTemplate.addEventListener("click", userConfirmFetchMarket);

var yesFetchMarket = document.getElementById("btnYesFetchMarket");
yesFetchMarket.addEventListener("click", fetchServerData);

var noFetchMarket = document.getElementById("btnCancelFetchMarket");
noFetchMarket.addEventListener("click", noFetchServerData);

var refreshMarketList = document.getElementById("btnRefreshMarket");
refreshMarketList.addEventListener("click", userConfirmRefreshMarketList);

var yesRefreshMarketList = document.getElementById("btnYesRefreshMarket");
yesRefreshMarketList.addEventListener("click", refeshMarketData);

var noRefreshMarketList = document.getElementById("btnCancelRefreshMarket");
noRefreshMarketList.addEventListener("click", notRefreshMarket);

var pauseResume = document.getElementById("btnPauseResume");
pauseResume.addEventListener("click", tooglePauseResume);

var validatePushFreqOnFocus = document.getElementById("inputPushFrequency");
validatePushFreqOnFocus.addEventListener("focus", validatePushFrequencyOnFocus);

var validatePushFeeqOutFocus = document.getElementById("inputPushFrequency");
validatePushFeeqOutFocus.addEventListener("focusout", validateNewPushFrequencyOutFocus);
