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

  const ss = openSpreadsheetSafe(sheetId, spreadsheetName);
  const spreadsheetTab = getSheetSafe(ss, spreadsheetTabName, spreadsheetName);
  
  const lastRow = spreadsheetTab.getLastRow();
  const maxAvailableRows = lastRow - spreadsheetFirstRowNumber + 1;
  
  if (maxAvailableRows <= 0) {
    Logger.log(`[INFO] No data rows found starting from row ${spreadsheetFirstRowNumber} in ${spreadsheetName} (${spreadsheetTabName}).`);
    return [];
  }

  const spreadsheetTabFirstColumn = getRangeValuesSafe(
    spreadsheetTab, 
    spreadsheetFirstRowNumber, 
    spreadsheetFirstColumnNumber, 
    maxAvailableRows, 
    1, 
    `${spreadsheetName} -> ${spreadsheetTabName} (first column)`
  );
  
  const spreadsheetLastRowOfData = getLastRowOfData(spreadsheetTabFirstColumn, spreadsheetFirstRowNumber);
  const spreadsheetTotalRowCount = spreadsheetLastRowOfData - spreadsheetFirstRowNumber + 1; // Add one back to account for the first row itself
  
  if (spreadsheetTotalRowCount <= 0) {
    Logger.log(`[INFO] Calculated total row count is non-positive for ${spreadsheetName} (${spreadsheetTabName}).`);
    return [];
  }
  
  const spreadsheetTotalColumnCount = spreadsheetTab.getLastColumn() - spreadsheetFirstColumnNumber + 1; // Add one back to account for the first column itself
  
  Logger.log(`********** Collecting data from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(`First row of data for sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetFirstRowNumber)
  Logger.log(`Actual last row of sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetLastRowOfData);
  Logger.log(`Total record count for (${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetTotalRowCount);
  Logger.log(`********** End data collection from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(``);

  const collectedData = getRangeValuesSafe(
    spreadsheetTab, 
    spreadsheetFirstRowNumber, 
    spreadsheetFirstColumnNumber, 
    spreadsheetTotalRowCount, 
    spreadsheetTotalColumnCount, 
    `${spreadsheetName} -> ${spreadsheetTabName}`
  );

  return collectedData;

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function collectDataByDefinedRange(sheetId, spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetName, spreadsheetTabName, spreadsheetTotalColumnCount) {

  const ss = openSpreadsheetSafe(sheetId, spreadsheetName);
  const spreadsheetTab = getSheetSafe(ss, spreadsheetTabName, spreadsheetName);
  
  const lastRow = spreadsheetTab.getLastRow();
  const maxAvailableRows = lastRow - spreadsheetFirstRowNumber + 1;
  
  if (maxAvailableRows <= 0) {
    Logger.log(`[INFO] No data rows found starting from row ${spreadsheetFirstRowNumber} in ${spreadsheetName} (${spreadsheetTabName}).`);
    return [];
  }

  const spreadsheetTabFirstColumn = getRangeValuesSafe(
    spreadsheetTab, 
    spreadsheetFirstRowNumber, 
    spreadsheetFirstColumnNumber, 
    maxAvailableRows, 
    1, 
    `${spreadsheetName} -> ${spreadsheetTabName} (first column)`
  );
  
  const spreadsheetLastRowOfData = getLastRowOfData(spreadsheetTabFirstColumn, spreadsheetFirstRowNumber);
  const spreadsheetTotalRowCount = spreadsheetLastRowOfData - spreadsheetFirstRowNumber + 1; // Add one back to account for the first row itself

  if (spreadsheetTotalRowCount <= 0) {
    Logger.log(`[INFO] Calculated total row count is non-positive for ${spreadsheetName} (${spreadsheetTabName}).`);
    return [];
  }

  Logger.log(`********** Collecting data from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(`First row of data for sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetFirstRowNumber)
  Logger.log(`Actual last row of sheet(${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetLastRowOfData);
  Logger.log(`Total record count for (${spreadsheetName}), tab (${spreadsheetTabName}): ` + spreadsheetTotalRowCount);
  Logger.log(`********** End data collection from sheet(${spreadsheetName}), tab (${spreadsheetTabName}) **********`)
  Logger.log(``);

  const collectedData = getRangeValuesSafe(
    spreadsheetTab, 
    spreadsheetFirstRowNumber, 
    spreadsheetFirstColumnNumber, 
    spreadsheetTotalRowCount, 
    spreadsheetTotalColumnCount, 
    `${spreadsheetName} -> ${spreadsheetTabName}`
  );

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
  const ss = openSpreadsheetSafe(sourceId, sourceSheetName);
  const sheet = getSheetSafe(ss, sourceTab, sourceSheetName);

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

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Robust safe wrappers for logging, performance auditing, and try-catch safety

function openSpreadsheetSafe(id, description) {
  const start = Date.now();
  Logger.log(`[DEBUG] Opening spreadsheet: "${description}" (ID: ${id})...`);
  try {
    const ss = SpreadsheetApp.openById(id);
    Logger.log(`[DEBUG] Successfully opened "${description}" in ${Date.now() - start}ms`);
    return ss;
  } catch (e) {
    Logger.log(`[ERROR] Failed to open spreadsheet "${description}" (ID: ${id}): ${e.toString()}`);
    throw e;
  }
}

function getSheetSafe(ss, tabName, description) {
  Logger.log(`[DEBUG] Fetching tab "${tabName}" from "${description}"...`);
  try {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      throw new Error(`Tab "${tabName}" not found in spreadsheet "${description}"`);
    }
    return sheet;
  } catch (e) {
    Logger.log(`[ERROR] getSheetSafe failed for tab "${tabName}" from "${description}": ${e.toString()}`);
    throw e;
  }
}

function getRangeValuesSafe(sheet, startRow, startCol, numRows, numCols, description) {
  const start = Date.now();
  Logger.log(`[DEBUG] Reading values from "${description}" (Range: StartRow=${startRow}, StartCol=${startCol}, Rows=${numRows}, Cols=${numCols})...`);
  try {
    if (numRows <= 0 || numCols <= 0) {
      throw new Error(`Invalid range size: Rows=${numRows}, Cols=${numCols}`);
    }
    const values = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
    Logger.log(`[DEBUG] Successfully read ${values.length} rows from "${description}" in ${Date.now() - start}ms`);
    return values;
  } catch (e) {
    Logger.log(`[ERROR] Failed to read range values from "${description}": ${e.toString()}`);
    throw e;
  }
}




