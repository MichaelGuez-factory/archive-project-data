function getListOfScoroIdsForArchive(thresholdDate, projectsToArchive) {

  const scoroIdsToArchive = [];

  projectsToArchive.forEach((record) => {

    //Logger.log(record);
    scoroIdsToArchive.push(parseInt(record[0]).toString().trim());

  });

  return scoroIdsToArchive;

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getListOfprojectReferencesForArchive(thresholdDate, projectsToArchive) {

  const projectReferencesToArchive = [];

  projectsToArchive.forEach((record) => {

    //Logger.log(record);
    projectReferencesToArchive.push(record[11]);

  });

  return projectReferencesToArchive;

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function collectData(sheetId, spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetName, spreadsheetTabName) {

  const ss = SpreadsheetApp.openById(sheetId);
  const spreadsheetTab = ss.getSheetByName(spreadsheetTabName);
  const spreadsheetTabFirstColumn = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTab.getLastRow(), 1).getValues();
  const spreadsheetLastRowOfData = getLastRowOfData(spreadsheetTabFirstColumn, spreadsheetFirstRowNumber);
  const spreadsheetTotalRowCount = spreadsheetLastRowOfData - spreadsheetFirstRowNumber + 1 // Add one back to account for the first row itself
  const spreadsheetTotalColumnCount = spreadsheetTab.getLastColumn() - spreadsheetFirstColumnNumber + 1; // Add one back to account for the first column itself
  Logger.log(`********** Collecting data from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(`First row of data for sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetFirstRowNumber)
  Logger.log(`Actual last row of sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetLastRowOfData);
  Logger.log(`Total record count for (${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetTotalRowCount);
  Logger.log(`********** End data collection from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(``);

  const collectedData = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTotalRowCount, spreadsheetTotalColumnCount).getValues();

  return collectedData;

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function collectDataByDefinedRange(sheetId, spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetName, spreadsheetTabName, spreadsheetTotalColumnCount) {

  const ss = SpreadsheetApp.openById(sheetId);
  const spreadsheetTab = ss.getSheetByName(spreadsheetTabName);
  const spreadsheetTabFirstColumn = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTab.getLastRow(), 1).getValues();
  const spreadsheetLastRowOfData = getLastRowOfData(spreadsheetTabFirstColumn, spreadsheetFirstRowNumber);
  const spreadsheetTotalRowCount = spreadsheetLastRowOfData - spreadsheetFirstRowNumber + 1 // Add one back to account for the first row itself

  Logger.log(`********** Collecting data from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(`First row of data for sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetFirstRowNumber)
  Logger.log(`Actual last row of sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetLastRowOfData);
  Logger.log(`Total record count for (${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetTotalRowCount);
  Logger.log(`********** End data collection from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(``);

  const collectedData = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTotalRowCount, spreadsheetTotalColumnCount).getValues();

  return collectedData;

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/*This function gets the last row of the indicated column.  This is a useful tool for finding the last value of a table that might also be running array formulas, which create the illusion that the data extends to the bottom of the sheet*/

/*const lastRowOfData = getLastRowOfData(spreadsheetSelectedColumn, spreadsheetFirstRowNumber);*/

/*Requires the Origin Column (as values) and a # for the first row of data.*/

function getLastRowOfData(spreadsheetSelectedColumn, spreadsheetFirstRowNumber) {

  /* Configuration */
  let spreadSheetLastRow = 0; // Initialize a variable to store the result

  /* Iterate backwards until a value is found */
  //Logger.log(`Reviewing data limits from row ${spreadsheetSelectedColumn.length - 1}`);
  for (let i = spreadsheetSelectedColumn.length - 1; i >= 0; i--) {
    if (spreadsheetSelectedColumn[i][0] !== "") {
      spreadSheetLastRow = i + spreadsheetFirstRowNumber;
      break; // Found a non-empty cell, exit the loop
    }
  }

  /* Return value */
  if (spreadSheetLastRow > 0) {
    return (spreadSheetLastRow);
  } else {
    return (spreadsheetFirstRowNumber);
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function findHeader(sourceId, sourceTab, sourceSheetName, searchField, targetHeader) {

  Logger.log(`********** Searching for headers in sheet(${sourceSheetName}), tab (${sourceTab}) **********`);
  const ss = SpreadsheetApp.openById(sourceId);
  const sheet = ss.getSheetByName(sourceTab);

  /*Get Header Row Number by searching for value*/
  const data = sheet.getRange(searchField).getValues();

  for (i = 0; i < data.length; i++) {

    if (data[i] == targetHeader) {

      i++;
      Logger.log(`Found header in row ${i}`);
      Logger.log(`********** End search for headers in sheet(${sourceSheetName}), tab (${sourceTab}) **********`);
      Logger.log(``);
      return (i);

    } 
  }

  if (i = data.length) {

    Logger.log(`Did not find "${targetHeader}" in data field "${searchField}"`);
    Logger.log(`********** End search for headers in sheet(${sourceSheetName}), tab (${sourceTab}) **********`);
    Logger.log(``);
    return (0);

  }
}



