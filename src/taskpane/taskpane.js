/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import getRangeString from "./getRangeString";
import { flatten } from "flat";
import { createOrUpdateQuery, getQuery, listQueries, deleteQuery } from "./db";
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // register handlers here
    document.getElementById("run").onclick = run;
    document.getElementById("save").onclick = save;
    document.getElementById("del").onclick = del;
    document.getElementById("queryList").onchange = loadSelectedQuery;
    loadQueryList();
  }
});

export async function run() {
  try {
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
      const rangeString = getRangeString(headers.length, rangeBody.length);
      let range = sheet.getRange(rangeString);
      range.values = rangeBody;

      // autofit cells
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      // convert to table
      const table = sheet.tables.add(rangeString, true);
      table.name = queryName;

      // activate sheet and sync
      sheet.activate();
      return context.sync();
    });
  } catch (error) {
    console.error(error);
    console.error("Error code: " + error.code);
    console.error("Error message: " + error.message);
    console.error("Error debugInfo: " + error.debugInfo);
  }
}

export async function save() {
  const queryName = document.getElementById("queryName").value;
  const url = document.getElementById("url").value;
  const apiKey = document.getElementById("apiKey").value;
  const query = document.getElementById("query").value;
  createOrUpdateQuery(queryName, url, apiKey, query);
  loadQueryList();
}

export async function del() {
  const queryName = document.getElementById("queryList").value;
  deleteQuery(queryName);
  document.getElementById("queryName").value = "";
  document.getElementById("url").value = "";
  document.getElementById("apiKey").value = "";
  document.getElementById("query").value = "";

  loadQueryList();
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
}
