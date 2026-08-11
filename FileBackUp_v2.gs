function fileBackUp_v2() {

  /* Back-Up Data Sources */
  Logger.log(`Backing up the data files ....`);

  copySpreadsheetToSharedDrive(`PData_Scoro Project Update_v2`, `1_w_xznqERtazm0W7Gs0TXDvKaB0h0GNyPUSscx6610Q`);
  copySpreadsheetToSharedDrive(`PData_Active Projects`, `1OSgoq-4TDjud6WbSAfK1LPBT4FO4xB7sbwmDE4Mt5B8`);
  copySpreadsheetToSharedDrive(`PData_Project Archive_Current`, `1E7RJn0eo2ZQ7nul1v3iS7DH57esvP_GhsO7tEOBwDWI`);
  copySpreadsheetToSharedDrive(`PData_Project Archive_2024`, `1EWLspx--rvzzF_TY2N_kKnoFiPBqE9zFXmLca5SDpwo`);
  copySpreadsheetToSharedDrive(`Pdata_CollectedProjectBudgets`, `1znzFr87-_EqkxfvribcZNimweO9r99tj2Q_RU_nIKxc`);
  copySpreadsheetToSharedDrive(`Pdata_CollectedProjectResourceSchedules`, `1XyLOuaQ-bxwz2TI-ORutxp3RHo6h6srHMaiaj3c49lM`);
  copySpreadsheetToSharedDrive(`Install Schedule Utility_v100`, `18o-aMpo4cLMh-g6P0iyqjLFyfACzXAGZVI7di5g-8t4`);
  copySpreadsheetToSharedDrive(`Fabrication Schedule Utility_v202`, `1bZmaAAXqfWUTC80I1F2_VVEN7jAKynRV_BywNGF0kK8`);
  copySpreadsheetToSharedDrive(`Production Timeline_V2`, `1LkGr3dcbUGnUiJ7xtuZHwQnteEJFjg9hEzq11dojX-I`);
  copySpreadsheetToSharedDrive(`FINAL - Weekly Hot Bids Report`, `1c_OxqcR8-648vQJR0cdReybQF8HRzcS3wrDkKmj4gos`);
  copySpreadsheetToSharedDrive(`021_TFNYC Materials Resource Utility_v2`, `1yMuZiUh86hURmD6IgwtoV2iq4xEk5PA0n24NuX_LVbk`);
  copySpreadsheetToSharedDrive(`020_TFNYC Labor & Equipment Resource Utility_v2`, `1NQ1mZPUk6jrqO1_1x4ODKfJDomuZeQRjN_QCh0o7D-M`);
  copySpreadsheetToSharedDrive(`010_TFNYC Resource Codes Utility_v3_WIP`, `1mMAmZhL9WCdPzVm5hYDf85k10sgMDiI5gwgwxYzXCa0`);
  copySpreadsheetToSharedDrive(`022_TFNYC Vendor & Services Resource Utility_v2`, `1957DrbnmAO-7UiGLfXB4NrYtXrlQclhBHB4GvaVx5Cg`);
  copySpreadsheetToSharedDrive(`023_TFNYC Admin Resource Utility_v2`, `1elHBC2tp6XfcMrj3eHno3vqpEWbC3cKuyfXMncaj-fE`);
  copySpreadsheetToSharedDrive(`Assembly Calculator_v1001`, `1PKy0wkgjKUkSOsOLURXtm7eaNvoY75ehCizfweJ7-eI`);
  copySpreadsheetToSharedDrive(`030_Staff and Equipment Resources_v2`, `1ieX6wxrrQoEU3QlbFw73yXs0vj1O4W1ovhwxVtFGxBg`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P0_Admin`, `1lUm1JSgKqQtH-T9vVrQS8vhalfvWpSOpeQa9hNzDnvY`);
  copySpreadsheetToSharedDrive(`PM Project Utility_v11P1_Open`, `1CnoKvQLvCOiOjwqnBOdHN8EZ6N2yzLE5QwYB7g8nVkU`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P2_Andi`, `1iaPicE1HF6sRbJpqc7ZAvz9o7G2Fex1DMbRQDoxm7Kw`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P3_Colin`, `1M2Zq9mKKA352FPULdlGbymNiFDM-7B_OsPeVFDgmutI`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P4_MarkK`, `1K85wQha34bmrJORIfMSmvjRFaPZUBT7X9oocY1OWZk8`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P5_Simon`, `1CFQsVWk9dvbqWNZTFGlzO13a1w3Kt8gPsje7W0lgJbk`);
  //copySpreadsheetToSharedDrive(`PM Project Utility_v11P6_Open`, `1-qivcoWo1PkvSRCoduGcOUUQSN3LL0yn6I0ZN3pFvr8`);
  copySpreadsheetToSharedDrive(`Timesheet`, `1xxsuiWc6jA1YopAGp_Mp4InLtFtrGmTPcpsVLGb2gP8`);


  Logger.log(`Back up complete.`)
  Logger.log(``);

  /* Manage Back-Up Folder */
  Logger.log(`Managing files in Back-Up Folder (removing files over 17 days old)... `);
  deleteOldFiles(`0AKkY5ZQNkiFaUk9PVA`, `1v10Yd4jJV-lcs68lYN7Z58LESur-hDPn`, 17);
  Logger.log(`File management complete.`)
  Logger.log(``);

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* This function will review the files in the Back-Up Folder.  It is using the file names to determine older copies that can be removed.  

/* const manageBackUpFolder = deleteOldFiles (driveID, folderID, numDaysToKeep); 

/* Requires Drive Id, Folder Id, and Number of Days to Keep (#) */

function deleteOldFiles(driveID, folderID, numDaysToKeep) {

  const folder = DriveApp.getFolderById(folderID);
  const folderName = folder.getName();
  const files = folder.getFiles();

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - numDaysToKeep);

  let filesRemoved = 0;
  let fileCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    fileCount++;

    try {
      // Extract the date part from the filename (before the first underscore)
      const dateString = fileName.substring(0, fileName.indexOf('_'));
      const fileDate = new Date(dateString);

      // Check if the file is older than the cutoff date
      if (fileDate < cutoffDate) {
        file.setTrashed(true); // Move to trash (safer than immediate delete)
        Logger.log(`Removed file: ${fileName}`);
        filesRemoved++;
      }
    } catch (error) {
      // Handle cases where the filename doesn't have the expected format
      Logger.log(`Skipping file due to filename format: ${fileName}.  Error: ${error}`);
    }
  }
  Logger.log(`Reviewed ${fileCount} files in the ${folderName} folder.`);
  Logger.log(`Removed ${filesRemoved} files that were more than ${numDaysToKeep} day old.`);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function copySpreadsheetToSharedDrive(newFileName, sourceFileId) {
  // The ID of the shared drive folder where the file will be saved.
  const targetFolderId = '1v10Yd4jJV-lcs68lYN7Z58LESur-hDPn';

  try {
    // Get the source file and target folder objects.
    const sourceFile = DriveApp.getFileById(sourceFileId);
    const targetFolder = DriveApp.getFolderById(targetFolderId);

    // Get the current time and format it for the filename.
    // Example format: 2025-08-12_08-30-00
    const timestamp = new Date().toISOString();

    // Construct the final filename.
    const finalName = `${timestamp}_${newFileName}`;

    // Copy the file to the target folder with the new name.
    sourceFile.makeCopy(finalName, targetFolder);

    Logger.log(`Successfully copied file: ${finalName}`);

  } catch (e) {
    // Log any errors that occur.
    Logger.log(`Error copying file: ${e.toString()}`);
  }
}