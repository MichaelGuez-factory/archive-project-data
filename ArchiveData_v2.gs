function archiveData_v2() {

  /* Date Configuration */
  const currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0);
  const thresholdDate = new Date(currentDate.getTime() - 14 * 24 * 60 * 60 * 1000); // 14 days ago
  thresholdDate.setHours(0, 0, 0, 0);
  const yesterday = new Date(currentDate.getTime() - 1 * 24 * 60 * 60 * 1000); // 1 day ago
  yesterday.setHours(0, 0, 0, 0);
  const currentDayMidnight = new Date(currentDate.getTime() - 0 * 24 * 60 * 60 * 1000); // 1 day ago
  currentDayMidnight.setHours(0, 0, 0, 0);
  Logger.log(`********** Date Configuration **********`);
  Logger.log(`currentDay is: ` + currentDate);
  Logger.log(`Threshold date is: ` + thresholdDate);
  Logger.log(`********** End Date Configuration **********`);
  Logger.log(``);

  /*Get Project Data*/
  const activeProjectsId = `1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8`;
  const spreadsheetFirstRowNumber = 2;
  const spreadsheetFirstColumnNumber = 1;
  const spreadsheetName = `Pdata_Active Projects`;
  const spreadsheetTabName = `Projects`;

  const ss = SpreadsheetApp.openById(activeProjectsId);
  const spreadsheetTab = ss.getSheetByName(spreadsheetTabName);
  const spreadsheetTabFirstColumn = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTab.getLastRow(), 1).getValues();
  const spreadsheetLastRowOfData = getLastRowOfData(spreadsheetTabFirstColumn, spreadsheetFirstRowNumber);
  const spreadsheetTotalRowCount = spreadsheetLastRowOfData - spreadsheetFirstRowNumber + 1 // Add one back to account for the first row itself
  const spreadsheetTotalColumnCount = spreadsheetTab.getLastColumn() - spreadsheetFirstColumnNumber;
  const activeProjects = spreadsheetTab.getRange(spreadsheetFirstRowNumber, spreadsheetFirstColumnNumber, spreadsheetTotalRowCount, spreadsheetTotalColumnCount).getValues();
  const projectsToArchive = [];
  const activeProjectList = [];
  activeProjects.forEach((record) => {

    if (record[10] < thresholdDate) {

      //Logger.log(record);
      projectsToArchive.push(record);

    }

  });

  activeProjects.forEach((record) => {

    if (record[10] >= thresholdDate) {

      //Logger.log(record);

      activeProjectList.push(record[11]);

    }

  });

  Logger.log(`********** Current Active Projects **********`)
  activeProjectList.forEach((record) => {

    Logger.log(`${record}`)

  });


  Logger.log(`********** Compiling archival schedule **********`)
  projectsToArchive.forEach((record) => {

    Logger.log(`${record[11]} was last updated on ${record[10]}`)

  });
  Logger.log(`********** End compiling archival schedule **********`)
  Logger.log(``);

  /*Map by ScoroId*/
  const scoroIdsToArchive = getListOfScoroIdsForArchive(thresholdDate, projectsToArchive);
  Logger.log(`********** Project Archive by Scoro Id**********`);
  Logger.log(scoroIdsToArchive);
  Logger.log(`********** End Project Archive by Scoro Id **********`);
  Logger.log(``);

  /*Map by Project Reference*/
  const projectReferencesToArchive = getListOfprojectReferencesForArchive(thresholdDate, projectsToArchive);
  Logger.log(`********** Project Archive by Project Reference **********`);
  Logger.log(projectReferencesToArchive);
  Logger.log(`********** End Project Archive by Project Reference **********`);
  Logger.log(``);


  /*File Back-Up*/
  //fileBackUp_v2();

  /*Archive Qotes_v2*/
  archiveQuotes_v2(scoroIdsToArchive);

  /*Archive Routing (works - 08/06/25)*/
  archiveRouting_v2(scoroIdsToArchive);

  /*Archive Items (works - 08/06/25)*/
  archiveItems_v2(scoroIdsToArchive);

  /*Archive Installs (works - 08/06/25)
  Installs are archived 14 days past there install date, UNLESS the project is still active.  No installs associated with an active project should be archived as of 12/31/2025.*/
  archiveInstalls_v2(thresholdDate,activeProjectList); 

  /*Archive Resource Schedules (works - 08/07/25)
  Resource Schedules are archived by 'yesterday', keeping a running list in the archive.  This allows us to capture data that might not be associated with an active project, as well as to capture assignements to installs and strike events.  Since the schedule vanishes when the project record is deleted there is only a need to maintain the archive.  We do not need to remove these from the project data.*/
  archiveProductionResourceSchedules_v2();
  /*archivePreProductionResourceSchedules_v2();*/ /*Deprecate*/

  /*Archive Expenses (works - 08/07/25)*/
  archiveTFNYCExpenses_v2(projectReferencesToArchive);

  /*Archive Projects (works - 08/07/25)*/
  archiveProjects_v2(scoroIdsToArchive);

  /*Remove Quotes*/
  removeQuotes_v2(scoroIdsToArchive);

  /*Remove Routing (works - 08/06/25)*/
  removeRouting_v2(scoroIdsToArchive);

  /*Remove Expenses (works - 08/08/25)*/
  removeExpenses_v2(projectReferencesToArchive);

  /*Remove Installs (works - 08/06/25)*/
  removeInstalls_v2(thresholdDate, activeProjectList);

  /*Remove Projects (works - 08/07/25)*/
  removeProjects_v2(scoroIdsToArchive);

}



