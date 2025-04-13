/**
 * Example script showing how to use the Google Drive Metadata Extractor
 * with a smaller dataset for testing purposes.
 * 
 * This script processes only the first 10 files you own in Google Drive.
 */

function generateTestMetadataCSV() {
  console.log('Starting test metadata collection process...');
  
  // Create a new spreadsheet
  console.log('Creating new spreadsheet...');
  const ss = SpreadsheetApp.create('Google Drive Metadata Test');
  const sheet = ss.getActiveSheet();
  
  // Set headers
  console.log('Setting up headers...');
  const headers = [
    'Source',
    'File Name',
    'File Type',
    'Owner',
    'Created Time',
    'Modified Time',
    'File Size (bytes)',
    'File ID',
    'Link'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get current user's email
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  console.log(`Processing files owned by: ${currentUserEmail}`);
  
  // Get Drive files - using a query to only get files owned by the current user
  console.log('Starting to process Drive files...');
  
  // Use the searchFiles method with a query to only get files owned by the current user
  let driveFiles = DriveApp.searchFiles('"me" in owners');
  
  let count = 0;
  const maxFiles = 10; // Limit to 10 files for testing
  
  // Variables for progress tracking
  let lastProgressUpdate = new Date().getTime();
  const progressInterval = 3000; // 3 seconds in milliseconds
  
  // Process files with a limit
  console.log(`Processing up to ${maxFiles} files for testing...`);
  
  // Collect file data
  const fileDataArray = [];
  
  while (driveFiles.hasNext() && count < maxFiles) {
    const file = driveFiles.next();
    
    // Create a direct link to the file
    const fileUrl = file.getUrl();
    
    const fileData = [
      'Drive',
      file.getName(),
      file.getMimeType(),
      currentUserEmail,
      file.getDateCreated(),
      file.getLastUpdated(),
      file.getSize(),
      file.getId(),
      fileUrl
    ];
    
    fileDataArray.push(fileData);
    count++;
    
    // Check if it's time to update progress (every 3 seconds)
    const currentTime = new Date().getTime();
    if (currentTime - lastProgressUpdate >= progressInterval) {
      console.log(`Progress: ${count}/${maxFiles} files processed (${Math.round(count/maxFiles*100)}%)`);
      lastProgressUpdate = currentTime;
    }
  }
  
  // Write all data at once
  console.log('Writing data to spreadsheet...');
  if (fileDataArray.length > 0) {
    sheet.getRange(2, 1, fileDataArray.length, headers.length).setValues(fileDataArray);
  }
  
  // Format the sheet
  console.log('Formatting spreadsheet...');
  sheet.autoResizeColumns(1, headers.length);
  
  // Create a CSV file
  console.log('Creating CSV file...');
  const csvContent = convertToCSV(sheet);
  const csvFile = DriveApp.createFile('google_files_metadata_test.csv', csvContent, MimeType.CSV);
  
  // Add a link to the CSV file in the spreadsheet
  console.log('Adding CSV download link to spreadsheet...');
  sheet.getRange(1, headers.length + 1).setValue('CSV File Link');
  sheet.getRange(2, headers.length + 1).setFormula('=HYPERLINK("' + csvFile.getUrl() + '", "Click here to download CSV")');
  
  console.log(`Test completed. Processed ${count} files owned by you.`);
  console.log(`Spreadsheet URL: ${ss.getUrl()}`);
  console.log(`CSV File URL: ${csvFile.getUrl()}`);
  
  // Try to show UI alert, but don't fail if it's not available
  try {
    SpreadsheetApp.getUi().alert('Test completed: ' + count + ' files processed. Metadata has been saved to a CSV file. Check the link in cell ' + 
      sheet.getRange(2, headers.length + 1).getA1Notation());
  } catch (e) {
    console.log('Note: UI alert could not be shown, but the process completed successfully.');
    console.log('Please check the spreadsheet and CSV file URLs above in the execution log.');
  }
}

/**
 * Simple CSV conversion function for the test script
 */
function convertToCSV(sheet) {
  console.log('Converting spreadsheet data to CSV format...');
  const data = sheet.getDataRange().getValues();
  let csv = '';
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < row.length; j++) {
      let cell = row[j];
      // Handle special characters and commas
      if (typeof cell === 'string') {
        cell = cell.replace(/"/g, '""');
        if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
          cell = '"' + cell + '"';
        }
      }
      csv += cell;
      if (j < row.length - 1) csv += ',';
    }
    csv += '\n';
  }
  
  console.log('CSV conversion completed');
  return csv;
}

// Add menu item to run the test script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Metadata Tools')
    .addItem('Generate Test CSV (10 files)', 'generateTestMetadataCSV')
    .addToUi();
} 