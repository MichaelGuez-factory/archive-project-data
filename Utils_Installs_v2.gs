function archiveInstalls_v2(thresholdDate, activeProjectList) {

  /*Get install data*/
  const sourceSheetId = `18o-aMpo4cLMh-g6P0iyqjLFyfACzXAGZVI7di5g-8t4`;
  const sourceSheetFirstRowNumber = 11;
  const sourceSheetFirstColumnNumber = 2;
  const sourceSheetName = `TFNYC Install Schedule Utility`;
  const sourceSheetTabName = `Install Schedule`;
  const sourceSheetTotalColumnCount = 14; //Only pull data up to column O (14 because I am starting on Column B, not A).  There remianing data is of no long term use.

  const collectedInstalls = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, sourceSheetTotalColumnCount);
  const installsToArchive = [];
  const activeInstalls = [];
  const archivalTimestamp = new Date();

  /* Create the Set of Active Project References using the SECOND item (index 1) from activeProjectList.  I want to ignore anything that is still on this list, even if it is over 14 days old.  */
  const activeProjectIds = new Set(activeProjectList.flat());


  Logger.log(`********** Starting archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
  Logger.log(`********** Active Project List **********`);
  Logger.log(activeProjectList);
  Logger.log(`********** Active Project Ids **********`);
  Logger.log([...activeProjectIds]);

  collectedInstalls.forEach((record) => {

    if (record[5] < thresholdDate && record[5] != `` && record[5] instanceof Date && !isNaN(record[5]) && !activeProjectIds.has(record[0])) {

      record.push("Event", archivalTimestamp);
      installsToArchive.push(record)

    } else {

      activeInstalls.push(record)

    }
  });

  if (installsToArchive.length > 0) {
    Logger.log(`Install records that were completed before ${thresholdDate}: `);

    installsToArchive.forEach((record) => {

      Logger.log(record)

    });
  } else {

    Logger.log(`No install records found for archive.`);

  }

  /*Append to Archive*/
  if (installsToArchive.length > 0) {
    const targetSheetId = `1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI`;
    const targetSheetFirstRowNumber = 2;
    const targetSheetFirstColumnNumber = 1;
    const targetSheetName = `Pdata_Project Archive_Current`;
    const targetSheetTabName = `Installs`;

    const target_ss = SpreadsheetApp.openById(targetSheetId);
    const target_sheet = target_ss.getSheetByName(targetSheetTabName);
    const lastRow = target_sheet.getLastRow();


    target_sheet.getRange(lastRow + 1, targetSheetFirstColumnNumber, installsToArchive.length, installsToArchive[0].length).setValues(installsToArchive);
    Logger.log(``);
    Logger.log(`Appended ${installsToArchive.length} rows of data to ${targetSheetName} (${targetSheetTabName})`);

  }

  Logger.log(`********** End archival for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
  Logger.log(``);

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function removeInstalls_v2(thresholdDate, activeProjectList) {

  /*Get routing data*/
  const sourceSheetId = `18o-aMpo4cLMh-g6P0iyqjLFyfACzXAGZVI7di5g-8t4`;
  const sourceSheetFirstRowNumber = 11;
  const sourceSheetFirstColumnNumber = 2;
  const sourceSheetName = `TFNYC Install Schedule Utility_v100`;
  const sourceSheetTabName = `Install Schedule`;
  const sourceSheetTotalColumnCount = 13; //Only pull data up to column N (14).  Since the data starts at column 2, I subtracted 1.
  const ss = SpreadsheetApp.openById(sourceSheetId);
  const sheet = ss.getSheetByName(sourceSheetTabName);

  const collectedData = collectDataByDefinedRange(sourceSheetId, sourceSheetFirstRowNumber, sourceSheetFirstColumnNumber, sourceSheetName, sourceSheetTabName, sourceSheetTotalColumnCount);
  //Logger.log(`Collected Install Data: `);
  //Logger.log(collectedData)

  /*Create a data log*/
  const dataToRemove = [];
  const dataToRetain = [];

  /* Create the Set of Active Project References using the SECOND item (index 1) from activeProjectList.  I want to ignore anything that is still on this list, even if it is over 14 days old.  */
  const activeProjectIds = new Set(activeProjectList.flat());

  Logger.log(`********** Starting removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
  Logger.log(`Searching ${sourceSheetName} (${sourceSheetTabName}) for records to remove.`);

  collectedData.forEach((record) => {

    if (record[5] < thresholdDate && record[5] != `` && record[5] instanceof Date && !isNaN(record[5]) && !activeProjectIds.has(record[0])) {

      dataToRemove.push(record)

    } else {

      dataToRetain.push(record)

    }
  });

  /*Report findings*/
  if (dataToRemove.length > 0) {
    Logger.log(`Discovered ${dataToRemove.length} records to remove: `);
    dataToRemove.forEach((record) => {

      Logger.log(record)

    });
  } else {

    Logger.log(`No install records found to remove.`);

  }

  /*Remove Records*/
  if (dataToRemove.length > 0) {

    for (i = collectedData.length + sourceSheetFirstRowNumber - 1; i >= sourceSheetFirstRowNumber; i--) {

      const rowData = sheet.getRange(i, sourceSheetFirstColumnNumber, 1, sourceSheetTotalColumnCount).getValues();

      let eventEndDate = rowData[0][5];
      let projectId = rowData[0][0];

      if (eventEndDate < thresholdDate && eventEndDate instanceof Date && !isNaN(eventEndDate) && !activeProjectIds.has(projectId)) {
        Logger.log(`Removing row:  ${rowData}`);
        sheet.deleteRow(i);
      }
    }
  }

  Logger.log(`********** End removal for sheet(${sourceSheetName}), tab (${sourceSheetTabName}) **********`);
  Logger.log(``);
}


