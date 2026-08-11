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

  const ss = openSpreadsheetSafe(activeProjectsId, spreadsheetName);
  const spreadsheetTab = getSheetSafe(ss, spreadsheetTabName, spreadsheetName);
  
  const lastRow = spreadsheetTab.getLastRow();
  const maxAvailableRows = lastRow - spreadsheetFirstRowNumber + 1;
  
  if (maxAvailableRows <= 0) {
    Logger.log(`[INFO] No active projects rows found.`);
    return;
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
    Logger.log(`[INFO] Calculated active projects total row count is non-positive.`);
    return;
  }
  
  const spreadsheetTotalColumnCount = spreadsheetTab.getLastColumn() - spreadsheetFirstColumnNumber + 1;
  
  const activeProjects = getRangeValuesSafe(
    spreadsheetTab, 
    spreadsheetFirstRowNumber, 
    spreadsheetFirstColumnNumber, 
    spreadsheetTotalRowCount, 
    spreadsheetTotalColumnCount, 
    `${spreadsheetName} -> ${spreadsheetTabName}`
  );
  
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

  // Archive and removal steps orchestrated sequentially inside try-catch safeguards
  const steps = [
    { name: "Archive Quotes", action: () => archiveQuotes_v2(scoroIdsToArchive) },
    { name: "Archive Routing", action: () => archiveRouting_v2(scoroIdsToArchive) },
    { name: "Archive Items", action: () => archiveItems_v2(scoroIdsToArchive) },
    { name: "Archive Installs", action: () => archiveInstalls_v2(thresholdDate, activeProjectList) },
    { name: "Archive Resource Schedules", action: () => archiveProductionResourceSchedules_v2() },
    { name: "Archive Expenses", action: () => archiveTFNYCExpenses_v2(projectReferencesToArchive) },
    { name: "Archive Projects", action: () => archiveProjects_v2(scoroIdsToArchive) },
    { name: "Remove Quotes", action: () => removeQuotes_v2(scoroIdsToArchive) },
    { name: "Remove Routing", action: () => removeRouting_v2(scoroIdsToArchive) },
    { name: "Remove Expenses", action: () => removeExpenses_v2(projectReferencesToArchive) },
    { name: "Remove Installs", action: () => removeInstalls_v2(thresholdDate, activeProjectList) },
    { name: "Remove Projects", action: () => removeProjects_v2(scoroIdsToArchive) }
  ];

  steps.forEach(step => {
    const start = Date.now();
    Logger.log(`\n========== Starting Step: ${step.name} ==========`);
    try {
      step.action();
      Logger.log(`========== Completed Step: ${step.name} in ${Date.now() - start}ms ==========\n`);
    } catch (e) {
      Logger.log(`[CRITICAL ERROR] Step "${step.name}" failed: ${e.toString()}\n${e.stack}`);
      Logger.log(`========== Failed Step: ${step.name} after ${Date.now() - start}ms ==========\n`);
    }
  });

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Diagnostic Tool to verify sheet IDs, tabs, connectivity, permissions, and grid bounds

function runDiagnostics() {
  Logger.log("=== STARTING WORKFLOW DIAGNOSTICS ===");
  
  const sources = [
    { id: "1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8", name: "Pdata_Active Projects", tabs: ["Projects", "Quote Export", "Quotes", "Routing Data Export_2", "Routing", "Item Data Export", "Project Export"] },
    { id: "18o-aMpo4cLMh-g6P0iyqjLFyfACzXAGZVI7di5g-8t4", name: "TFNYC Install Schedule Utility", tabs: ["Install Schedule"] },
    { id: "1bZmaAAXqfWUTC80I1F2_VVEN7jAKynRV_BywNGF0kK8", name: "Fabrication Schedule Utility_v202", tabs: ["Production Staff Schedule Export"] },
    { id: "1XyLOuaQ-bxwz2TI-ORutxp3RHo6h6srHMaiaj3c49lM", name: "Pdata_Collected Schedules", tabs: ["Production Staff Schedule_Past"] },
    { id: "1znzFr87-_EqkxfvribcZNimweO9r99tj2Q_RU_nIKxc", name: "Pdata_CollectedProjectBudgets", tabs: ["Active Project Budget Export"] },
    { id: "1CnoKvQLvCOiOjwqnBOdHN8EZ6N2yzLE5QwYB7g8nVkU", name: "PM Project Utility_v11P1", tabs: ["Project Materials  & Services Tracking"] },
    { id: "1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI", name: "Pdata_Project Archive_Current", tabs: ["Quotes", "Routing", "Items", "Installs", "Archived Project Expenses", "Projects", "PreProduction Staff Schedule"] }
  ];

  sources.forEach(src => {
    Logger.log(`\n--- Diagnosing: ${src.name} (ID: ${src.id}) ---`);
    try {
      const ss = openSpreadsheetSafe(src.id, src.name);
      Logger.log(`[PASS] Opened spreadsheet successfully.`);
      src.tabs.forEach(tabName => {
        try {
          const sheet = ss.getSheetByName(tabName);
          if (!sheet) {
            Logger.log(`  [FAIL] Tab "${tabName}" NOT FOUND!`);
          } else {
            Logger.log(`  [PASS] Tab "${tabName}" found. Dimensions: MaxRows=${sheet.getMaxRows()}, LastRow=${sheet.getLastRow()}, MaxCols=${sheet.getMaxColumns()}, LastCol=${sheet.getLastColumn()}`);
          }
        } catch (tabErr) {
          Logger.log(`  [FAIL] Error accessing tab "${tabName}": ${tabErr.toString()}`);
        }
      });
    } catch (ssErr) {
      Logger.log(`[FAIL] Could not open spreadsheet: ${ssErr.toString()}`);
    }
  });

  Logger.log("\n=== DIAGNOSTICS COMPLETE ===");
}




