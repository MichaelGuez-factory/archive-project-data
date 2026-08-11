function archiveQuotes_v2(scoroIdsToArchive) {

  if (scoroIdsToArchive.length > 0) {

    /*Get Item data*/
    const sourceSheetId = `1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8`;
    const sourceSheetFirstRowNumber = 2;
    const sourceSheetFirstColumnNumber = 1;
    const sourceSheetName = `Pdata_Active Projects`;
    const sourceSheetTabName = `Quote Export`;
    const spreadsheetTotalColumnCount = 10; //Only pull data up to column J (10).  

    const collectedData = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, spreadsheetTotalColumnCount);

    const dataToArchive = [];
    const archivalTimestamp = new Date();

    Logger.log(`********** Starting archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(`Searching ${sourceSheetName} (${sourceSheetTabName}) for records to archive. `);

    collectedData.forEach((record) => {

      let ScoroReference = parseInt(record[3]).toString().trim();

      if (scoroIdsToArchive.includes(ScoroReference)) {

        record.push(archivalTimestamp);
        dataToArchive.push(record)


      }
    });

    if (dataToArchive.length > 0) {
      Logger.log(`Discovered ${dataToArchive.length} records to archive.  Item records associated with projects in archive protocols: `);
      dataToArchive.forEach((record) => {

        Logger.log(record)

      });
    } else {

      Logger.log(`No Item records found for archival.`);

    }

    /*Append to Archive*/
    if (dataToArchive.length > 0) {
      const targetSheetId = `1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI`;
      const targetSheetFirstRowNumber = 2;
      const targetSheetFirstColumnNumber = 1;
      const targetSheetName = `Pdata_Project Archive_Current`;
      const targetSheetTabName = `Quotes`;

      const target_ss = openSpreadsheetSafe(targetSheetId, targetSheetName);
      const target_sheet = getSheetSafe(target_ss, targetSheetTabName, targetSheetName);
      const lastRow = target_sheet.getLastRow();


      target_sheet.getRange(lastRow + 1, targetSheetFirstColumnNumber, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
      Logger.log(``);
      Logger.log(`Appended ${dataToArchive.length} rows of data to ${targetSheetName} (${targetSheetTabName})`);
    }

    Logger.log(`********** End archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(``);
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function removeQuotes_v2(scoroIdsToArchive) {

  if (scoroIdsToArchive.length > 0) {

    /*Get routing data*/
    const sourceSheetId = `1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8`;
    const sourceSheetFirstRowNumber = 2;
    const sourceSheetFirstColumnNumber = 1;
    const sourceSheetName = `Pdata_Active Projects`;
    const sourceSheetTabName = `Quotes`;
    const sourceSheetTotalColumnCount = 11; //Only pull data up to column K (11).  There are calculated values starting with column L.
    const ss = openSpreadsheetSafe(sourceSheetId, sourceSheetName);
    const sheet = getSheetSafe(ss, sourceSheetTabName, sourceSheetName);

    const collectedData = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, sourceSheetTotalColumnCount);

    const dataToRemove = [];
    const dataToRetain = [];

    Logger.log(`********** Starting removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(`Searching ${sourceSheetName} (${sourceSheetTabName}) for records to remove.`);

    collectedData.forEach((record) => {

      let ScoroReference = parseInt(record[3]).toString().trim();

      if (scoroIdsToArchive.includes(ScoroReference)) {

        dataToRemove.push(record)

      } else {

        dataToRetain.push(record);

      }
    });

    if (dataToRemove.length > 0) {
      Logger.log(`Discovered ${dataToRemove.length} records to remove: `);
      dataToRemove.forEach((record) => {

        Logger.log(record)

      });
    } else {

      Logger.log(`No quote records found to remove.`);

    }

    /*Rebuild Active Routing*/
    if (dataToRetain.length > 0) {

      sheet.getRange(sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, collectedData.length + 1, sourceSheetTotalColumnCount).clear();
      Utilities.sleep(5000); // Pause for 5 seconds (5000 milliseconds)
      sheet.getRange(sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, dataToRetain.length, sourceSheetTotalColumnCount).setValues(dataToRetain);

    }

    Logger.log(`********** End removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(``);


  }
}


