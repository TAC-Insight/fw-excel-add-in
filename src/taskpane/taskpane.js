/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { flatten } from "flat";
import { createOrUpdateQuery, getQuery, listQueries, deleteQuery } from "./db";
import setAlertMsg from "./setAlertMsg";
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // register handlers here
    document.getElementById("run").onclick = run;
    document.getElementById("save").onclick = save;
    document.getElementById("del").onclick = del;
    document.getElementById("clearAlertMsg").onclick = clearAlertMsg;
    document.getElementById("queryList").onchange = loadSelectedQuery;
    loadQueryList();
  }
});

export async function run() {
  try {
    // set alert msg
    setAlertMsg(`Running query...`);

    const url = document.getElementById("url").value;
    const apiKey = document.getElementById("apiKey").value;
    const queryName = document.getElementById("queryName").value;
    const query = document.getElementById("query").value;

    let response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
      },
      body: JSON.stringify({ query }),
    });

    let { data } = await response.json();

    const key = Object.keys(data)[0];
    response = data[key.toString()];

    console.log("response");
    console.log(response);

    // Flatten and standardize results
    const flattenedResponse = response.map((item) => {
      return flatten(item);
    });

    console.log("flatternedResponse");
    console.log(flattenedResponse);

    const keys = [...new Set(flattenedResponse.flatMap(Object.keys))];
    const standardizedResponse = flattenedResponse.map((doc) => {
      keys.forEach((key) => {
        if (doc[key] === undefined) doc[key] = "";
      });
      return doc;
    });

    console.log("standardizedResponse");
    console.log(standardizedResponse);

    // Fire up the Excel engine
    await Excel.run(async (context) => {
      // get sheet
      let sheet = context.workbook.worksheets.getItemOrNullObject(queryName);
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(queryName);
      }

      // clear sheet
      sheet.getUsedRange().clear();

      // set range values
      let headers = Object.keys(standardizedResponse[0]);
      const rangeBody = standardizedResponse.map((doc) => {
        let row = [];
        headers.forEach((key) => {
          row.push(doc[key]);
        });
        return row;
      });
      rangeBody.unshift(headers);

      // make sure no arrays are in the range
      rangeBody.forEach((row) => {
        row.forEach((cell, index) => {
          if (Array.isArray(cell)) {
            row[index] = "";
          }
        });
      });

      console.log("rangeBody");
      console.log(rangeBody);
      let range = sheet.getCell(0, 0).getResizedRange(rangeBody.length - 1, rangeBody[0].length - 1);
      range.values = rangeBody;
      range.load("address");
      await context.sync();

      // autofit cells
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getRange(range.address).format.autofitColumns();
        sheet.getRange(range.address).format.autofitRows();
      }

      // set alert msg
      setAlertMsg(`Worksheet "${queryName}" created.`);

      // activate sheet and sync
      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    setAlertMsg(`Error: ${error.message}`);
    console.error(error);
    console.error("Error code: " + error.code);
    console.error("Error message: " + error.message);
    console.error("Error debugInfo: " + error.debugInfo);
  }
}

export async function save() {
  try {
    setAlertMsg(`Saving query...`);
    const queryName = document.getElementById("queryName").value;
    const url = document.getElementById("url").value;
    const apiKey = document.getElementById("apiKey").value;
    const query = document.getElementById("query").value;
    createOrUpdateQuery(queryName, url, apiKey, query);
    setAlertMsg("Query saved to your list");
    loadQueryList();
  } catch (error) {
    setAlertMsg(`Error: ${error.message}`);
    console.error(error);
    console.error("Error code: " + error.code);
    console.error("Error message: " + error.message);
    console.error("Error debugInfo: " + error.debugInfo);
  }
}

export async function del() {
  try {
    setAlertMsg("Deleting query...");
    const queryName = document.getElementById("queryList").value;
    deleteQuery(queryName);
    document.getElementById("queryName").value = "";
    document.getElementById("url").value = "";
    document.getElementById("apiKey").value = "";
    document.getElementById("query").value = "";
    setAlertMsg("Query deleted from your list");
    loadQueryList();
  } catch (error) {
    setAlertMsg(`Error: ${error.message}`);
    console.error(error);
    console.error("Error code: " + error.code);
    console.error("Error message: " + error.message);
    console.error("Error debugInfo: " + error.debugInfo);
  }
}

export async function loadQueryList() {
  const queryList = document.getElementById("queryList");
  queryList.innerHTML = "";
  let defaultOption = document.createElement("option");
  defaultOption.value = "Select a query";
  defaultOption.innerHTML = "Select a query";
  queryList.appendChild(defaultOption);
  const queries = listQueries();
  queries.forEach((queryName) => {
    let option = document.createElement("option");
    option.value = queryName;
    option.innerText = queryName;
    queryList.appendChild(option);
  });
}

export async function loadSelectedQuery() {
  const queryName = document.getElementById("queryList").value;
  if (queryName === "Select a query") {
    return;
  }
  const query = getQuery(queryName);
  document.getElementById("queryName").value = queryName;
  document.getElementById("url").value = query.url;
  document.getElementById("apiKey").value = query.apiKey;
  document.getElementById("query").value = query.query;
  setAlertMsg("Query loaded");
}

export async function clearAlertMsg() {
  setAlertMsg("");
}
