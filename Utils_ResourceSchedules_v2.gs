function archiveProductionResourceSchedules_v2() {

  /* Date Configuration 
  Resource Archival will use it's own thresholds.  The goal is to capture the most accurate snapshot possible BEFORE the project data is potentially removed from the board.  The user is capable of transforming historical data in unintended ways and I don't want to pick up all that static.  One the otherhand, project based archival will not capture people unassigned to a project, assigned to an install, or assigned placeholder project.  That makes archival by date the most efficient way to do this, with 'yesterday' being the most accurate picture.  The further we get away from yesterday, the more likely we will pick up static.  ONLY looking for yesterday though seems a little aggressive.  An issue with the automation could potentially cause multiple days to get missed.  So we're going to set a custom threshold designed to review three days worth of data.*/

  const currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0);
  const thresholdDate = new Date(currentDate.getTime() - 4 * 24 * 60 * 60 * 1000); // Looking for records from the past 3 days
  thresholdDate.setHours(0, 0, 0, 0);


  /*Get active resource data from last 3 days*/
  const sourceSheetId = `1bZmaAAXqfWUTC80I1F2_VVEN7jAKynRV_BywNGF0kK8`;
  const sourceSheetName = `Fabrication Schedule Utility_v202`;
  const sourceTabName = `Production Staff Schedule Export`;
  const sourceTabFirstRowNumber = 2;
  const sourceTabFirstColumnNumber = 1;

  Logger.log(`********** Starting archival for sheet(${sourceSheetName}), tab (${sourceTabName}) **********`);

  const collectedActiveResources = collectData(sourceSheetId, sourceTabFirstRowNumber, sourceTabFirstColumnNumber, sourceSheetName, sourceTabName);
  const activeResourcesToReview = [];

  collectedActiveResources.forEach((record) => {
    let resourceDate = new Date(record[0]);

    if (resourceDate > thresholdDate && resourceDate < currentDate && resourceDate != `` && resourceDate instanceof Date && !isNaN(resourceDate)) { activeResourcesToReview.push(record) }
  });

  /*Get archived resource keys from last 3 days*/
  const targetSheetId = `1XyLOuaQ-bxwz2TI-ORutxp3RHo6h6srHMaiaj3c49lM`;
  const targetSheetName = `Pdata_Collected Schedules`;
  const targetTabName = `Production Staff Schedule_Past`;
  const targetTabFirstRowNumber = 4;
  const targetTabFirstColumnNumber = 1;

  const collectedArchivedResources = collectData(targetSheetId, targetTabFirstRowNumber, targetTabFirstColumnNumber, targetSheetName, targetTabName);
  const archivedResourceKeys = [];

  collectedArchivedResources.forEach((record) => {
    let resourceDate = new Date(record[0]);

    if (resourceDate > thresholdDate && resourceDate < currentDate && resourceDate != `` && resourceDate instanceof Date && !isNaN(resourceDate)) { archivedResourceKeys.push(record[10]) }
  });


  const dataToArchive = [];
  const archivalTimestamp = new Date();

  activeResourcesToReview.forEach((record) => {

    if (!archivedResourceKeys.includes(record[10])) {

      record.push(archivalTimestamp);
      dataToArchive.push(record);

    }
  });

  if (dataToArchive.length > 0) {

    Logger.log(`${activeResourcesToReview.length} resource records reviewed.`);
    Logger.log(`${archivedResourceKeys.length} resource keys in date range.`);
    Logger.log(`${dataToArchive.length} resource records schedule for archive.`);
    Logger.log(``);
    Logger.log(`Resources scheduled for archive: `)

    dataToArchive.forEach((record) => {

      Logger.log(`${record[0]}_${record[1]}_${record[2]}`);

    });

  } else {

    Logger.log(`${activeResourcesToReview.length} resource records reviewed.`);
    Logger.log(`${archivedResourceKeys.length} resource keys in date range.`);
    Logger.log(`${dataToArchive.length} resource records schedule for archive.`);

  }

  /*Append to Archive*/
  if (dataToArchive.length > 0) {

    const targetSheet = SpreadsheetApp.openById(targetSheetId);
    const targetTab = targetSheet.getSheetByName(targetTabName);
    const targetTabLastRow = targetTab.getLastRow();


    targetTab.getRange(targetTabLastRow + 1, targetTabFirstColumnNumber, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
    Logger.log(``);
    Logger.log(`Appended ${dataToArchive.length} rows of data to ${targetSheetName} (${targetTabName})`);
  }

  Logger.log(`********** Ending archival for sheet(${sourceSheetName}), tab (${sourceTabName}) **********`);
  Logger.log(``);

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function archivePreProductionResourceSchedules_v2() {

  /* Date Configuration 
  Resource Archival will use it's own thresholds.  The goal is to capture the most accurate snapshot possible BEFORE the project data is potentially removed from the board.  The user is capable of transforming historical data in unintended ways and I don't want to pick up all that static.  One the otherhand, project based archival will not capture people unassigned to a project, assigned to an install, or assigned placeholder project.  That makes archival by date the most efficient way to do this, with 'yesterday' being the most accurate picture.  The further we get away from yesterday, the more likely we will pick up static.  ONLY looking for yesterday though seems a little aggressive.  An issue with the automation could potentially cause multiple days to get missed.  So we're going to set a custom threshold designed to review three days worth of data.*/

  const currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0);
  const thresholdDate = new Date(currentDate.getTime() - 4 * 24 * 60 * 60 * 1000); // Looking for records from the past 3 days
  thresholdDate.setHours(0, 0, 0, 0);


  /*Get active resource data from last 14 days*/
  const sourceSheetId = `1tObvoR4ojbAE6a5cRbdGyyy3LVy5ITSKTD0tg3cQkeI`;
  const sourceSheetName = `Pre-Production Schedule Utility_v202`;
  const sourceTabName = `PreProduction Staff Schedule Export`;
  const sourceTabFirstRowNumber = 2;
  const sourceTabFirstColumnNumber = 1;

  Logger.log(`********** Starting archival for sheet(${sourceSheetName}), tab (${sourceTabName}) **********`);

  const collectedActiveResources = collectData(sourceSheetId, sourceTabFirstRowNumber, sourceTabFirstColumnNumber, sourceSheetName, sourceTabName);
  const activeResourcesToReview = [];

  collectedActiveResources.forEach((record) => {
    let resourceDate = new Date(record[0]);

    if (resourceDate > thresholdDate && resourceDate < currentDate && resourceDate != `` && resourceDate instanceof Date && !isNaN(resourceDate)) { activeResourcesToReview.push(record) }
  });

  /*Get archived resource keys from last 14 days*/
  const targetSheetId = `1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI`;
  const targetSheetName = `PData_Project Archive_Current`;
  const targetTabName = `PreProduction Staff Schedule`;
  const targetTabFirstRowNumber = 2;
  const targetTabFirstColumnNumber = 1;

  const collectedArchivedResources = collectData(targetSheetId, targetTabFirstRowNumber, targetTabFirstColumnNumber, targetSheetName, targetTabName);
  const archivedResourceKeys = [];

  collectedArchivedResources.forEach((record) => {
    let resourceDate = new Date(record[0]);

    if (resourceDate > thresholdDate && resourceDate < currentDate && resourceDate != `` && resourceDate instanceof Date && !isNaN(resourceDate)) { archivedResourceKeys.push(record[10]) }
  });


  const dataToRetain = [];
  const dataToArchive = [];
  const archivalTimestamp = new Date();

  activeResourcesToReview.forEach((record) => {

    if (archivedResourceKeys.includes(record[10])) {

      dataToRetain.push(record);

    } else {

      record.push(archivalTimestamp);
      dataToArchive.push(record);

    }
  });

  if (dataToArchive.length > 0) {

    Logger.log(`${activeResourcesToReview.length} resource records reviewed.`);
    Logger.log(`${archivedResourceKeys.length} resource keys in date range.`);
    Logger.log(`${dataToArchive.length} resource records schedule for archive.`);
    Logger.log(``);
    Logger.log(`Resources scheduled for archive: `)

    dataToArchive.forEach((record) => {

      Logger.log(`${record[0]}_${record[1]}_${record[2]}`);

    });

  } else {

    Logger.log(`${activeResourcesToReview.length} resource records reviewed.`);
    Logger.log(`${archivedResourceKeys.length} resource keys in date range.`);
    Logger.log(`${dataToArchive.length} resource records schedule for archive.`);

  }

  /*Append to Archive*/
  if (dataToArchive.length > 0) {

    const targetSheet = SpreadsheetApp.openById(targetSheetId);
    const targetTab = targetSheet.getSheetByName(targetTabName);
    const targetTabLastRow = targetTab.getLastRow();


    targetTab.getRange(targetTabLastRow + 1, targetTabFirstColumnNumber, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
    Logger.log(``);
    Logger.log(`Appended ${dataToArchive.length} rows of data to ${targetSheetName} (${targetTabName})`);
  }

  Logger.log(`********** Ending archival for sheet(${sourceSheetName}), tab (${sourceTabName}) **********`);
  Logger.log(``);

}
