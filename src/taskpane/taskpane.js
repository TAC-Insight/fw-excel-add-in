/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import getRangeString from "./getRangeString";
import { flatten } from "flat";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    const url = "https://fwt.fast-weigh.dev/v1/graphql";
    const apiKey = "bb9274ff-15d0-446a-b897-5a23b9f34f22";
    const sheetName = "GetCustomers";
    const tableName = "GetCustomers";
    const query = `
      query GetTickets {
        LoadTicket(limit: 1000, order_by: {TicketKey: desc}) {
          TicketNumber
          TicketDate
          TicketDateTime
          GrossWeight
          TareWeight
          NetWeight
          Order {
            OrderNumber
            Salesperson {
              Name
            }
          }
          Truck {
            TruckID
            Hauler {
              HaulerID
            }
          }
        }
      }
    `;

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
      let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add(sheetName);
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
      table.name = tableName;

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
