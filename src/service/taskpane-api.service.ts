import data from "../shared/constant.json";

var fetchAPITableData: Object;
var submitUserTableDataResponse: Object;
var updatedRowsSend: string;

export const getTableDatafromAPI = async (token: string) => {
  try {
    const axios = await require("axios");

    fetchAPITableData = await axios.get(data.fetchAPIURL + data.coIdValue, {
      headers: {
        Authorization: token,
      },
    });
    return fetchAPITableData;
  } catch (error) {
    console.error(error);
  }
};

export const saveTableDataAPI = async (token: string, body: Object) => {
  try {
    const axios = await require("axios");

    submitUserTableDataResponse = await axios.post(data.pushAPIURL + data.coIdValue + "&AcctId=" + data.acctId, body, {
      headers: {
        Authorization: token,
      },
    });
    updatedRowsSend = Object.values(submitUserTableDataResponse).shift();
    return updatedRowsSend;
  } catch (error) {
    console.error(error);
  }
};





