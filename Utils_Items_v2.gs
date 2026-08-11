function archiveItems_v2(scoroIdsToArchive) {

  if (scoroIdsToArchive.length > 0) {

    /*Get Item data*/
    const sourceSheetId = `1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8`;
    const sourceSheetFirstRowNumber = 2;
    const sourceSheetFirstColumnNumber = 1;
    const sourceSheetName = `Pdata_Active Projects`;
    const sourceSheetTabName = `Item Data Export`;
    const spreadsheetTotalColumnCount = 11; //Only pull data up to column K (11).  There remianing data is of no long term use.

    const collectedData = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, spreadsheetTotalColumnCount);

    const dataToArchive = [];
    const archivalTimestamp = new Date();

    Logger.log(`********** Starting archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(`Searching ${sourceSheetName} (${sourceSheetTabName}) for records to archive. `);

    collectedData.forEach((record) => {

      let ScoroReference = parseInt(record[0]).toString().trim();

      if (scoroIdsToArchive.includes(ScoroReference)) {

        record.push(archivalTimestamp);
        dataToArchive.push(record);


      }
    });

    if (dataToArchive.length > 0) {
      Logger.log(`Discovered ${dataToArchive.length} records to archive.  Item records associated with projects in archive protocols: `);
      dataToArchive.forEach((record) => {

        Logger.log(record);

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
      const targetSheetTabName = `Items`;

      const target_ss = SpreadsheetApp.openById(targetSheetId);
      const target_sheet = target_ss.getSheetByName(targetSheetTabName);
      const lastRow = target_sheet.getLastRow();


      target_sheet.getRange(lastRow + 1, targetSheetFirstColumnNumber, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
      Logger.log(``);
      Logger.log(`Appended ${dataToArchive.length} rows of data to ${targetSheetName} (${targetSheetTabName})`);
    }

    Logger.log(`********** End archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
    Logger.log(``);
  }
}

