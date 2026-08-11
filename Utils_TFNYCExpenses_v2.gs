function archiveTFNYCExpenses_v2(projectReferencesToArchive) {

  if (projectReferencesToArchive.length > 0) {

    /*Get routing data*/
    const sourceSheetId = `1znzFr87-_EqkxfvribcZNimweO9r99tj2Q_RU_nIKxc`;
    const sourceSheetFirstRowNumber = 2;
    const sourceSheetFirstColumnNumber = 1;
    const sourceSheetName = `Pdata_CollectedProjectBudgets`;
    const sourceSheetTabName = `Active Project Budget Export`;

    const collectedExpenses = collectData(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName);

    const dataToArchive = [];
    const dataToRetain = [];
    const archivalTimestamp = new Date();

    collectedExpenses.forEach((record) => {

      let expensesProjectReference = record[0];

      if (projectReferencesToArchive.includes(expensesProjectReference)) {

        record.push(archivalTimestamp);
        dataToArchive.push(record);

      } else {

        dataToRetain.push(record)

      }
    });

    if (dataToArchive.length > 0) {
      Logger.log(`Expense records associated with projects in archive protocols: `);

      dataToArchive.forEach((record) => {

        Logger.log(record)

      });
    } else {

      Logger.log(`No expense records found for archive.`);

    }

    /*Append to Archive*/
    if (dataToArchive.length > 0) {
      const targetSheetId = `1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI`;
      const targetSheetFirstRowNumber = 2;
      const targetSheetFirstColumnNumber = 1;
      const targetSheetName = `Pdata_Project Archive_Current`;
      const targetSheetTabName = `Archived Project Expenses`;

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

function removeExpenses_v2(projectReferencesToArchive) {

  if (projectReferencesToArchive.length > 0) {

    Logger.log(`Searching the last active PM utility for records to remove....`);
    Logger.log(``);

    const projectUtilities = [
      [`1CnoKvQLvCOiOjwqnBOdHN8EZ6N2yzLE5QwYB7g8nVkU`, `PM Project Utility_v11P1`, `Project Materials  & Services Tracking`]
    ];

    /*Get routing data*/

    projectUtilities.forEach((record) => {

      const sourceSheetId = record[0];
      const sourceSheetFirstColumnNumber = 2;
      const sourceSheetName = record[1];
      const sourceSheetTabName = record[2];

      const sourceHeaderRowNumber = findHeader(sourceSheetId, sourceSheetTabName, sourceSheetName, `A1:A`, `Headers`);

      const sourceSheetFirstRowNumber = sourceHeaderRowNumber + 1;
      const sourceSheetTotalColumnCount = 10; //Only pull data up to column K (11).  Since the data starts at column 2, I subtracted 1.
      const ss = openSpreadsheetSafe(sourceSheetId, sourceSheetName);
      const sheet = getSheetSafe(ss, sourceSheetTabName, sourceSheetName);

      //Do no execute if a header row was not found (findHeader returned 0)
      let collectedData = [];
      if (sourceHeaderRowNumber > 0) {
        collectedData = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, sourceSheetTotalColumnCount);
        //Logger.log(`Collected Resource Data: `);
        //Logger.log(collectedData)
      }

      /*Create a data log*/
      const dataToRemove = [];
      const dataToRetain = [];


      //Do not execute if collectedData is undefined
      if (collectedData.length > 0) {
        collectedData.forEach((record) => {

          if (projectReferencesToArchive.includes(record[0])) {

            dataToRemove.push(record)

          } else {

            dataToRetain.push(record)

          }
        });
      }

      /*Remove data*/
      Logger.log(`********** Starting removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
      //Do not execute if there is no data to reomce or if a header row was never found to determine the removal range
      if (dataToRemove.length > 0 && sourceHeaderRowNumber > 0) {
        Logger.log(`Discovered ${dataToRemove.length} records to remove: `);
        dataToRemove.forEach((record) => {

          Logger.log(record)

        });
      } else {

        Logger.log(`No expense records found to remove.`);

      }

      if (dataToRemove.length > 0) {

        for (i = collectedData.length + sourceSheetFirstRowNumber; i >= sourceSheetFirstRowNumber; i--) {

          const rowData = sheet.getRange(i, sourceSheetFirstColumnNumber, 1, sourceSheetTotalColumnCount).getValues();

          let projectReference = rowData[0][0];

          if (projectReferencesToArchive.includes(projectReference)) {
            Logger.log(`Removing row:  ${rowData}`);
            sheet.deleteRow(i);
          }
        }
      }

      Logger.log(`********** End removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
      Logger.log(``);
    });
  }
}




