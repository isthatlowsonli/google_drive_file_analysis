function generateMetadataCSV() {
  console.log('Starting metadata collection process...');
  
  // Create a new spreadsheet
  console.log('Creating new spreadsheet...');
  const ss = SpreadsheetApp.create('Google Drive and Photos Metadata');
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
  
  // Count total files owned by the current user
  console.log('Counting total files owned by you...');
  const totalFiles = countOwnedFiles();
  console.log(`Total files owned by you: ${totalFiles}`);
  
  // Get Drive files - using a query to only get files owned by the current user
  console.log('Starting to process Drive files...');
  
  // Use the searchFiles method with a query to only get files owned by the current user
  let driveFiles = DriveApp.searchFiles('"me" in owners');
  
  let count = 0;
  
  // Variables for progress tracking
  let lastProgressUpdate = new Date().getTime();
  const progressInterval = 3000; // 3 seconds in milliseconds
  
  // Process files
  console.log(`Processing all files (this may take a while)...`);
  
  // Process in batches to avoid memory issues
  const batchSize = 5000;
  let currentBatch = [];
  let batchNumber = 1;
  let row = 2;
  
  while (driveFiles.hasNext()) {
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
    
    currentBatch.push(fileData);
    count++;
    
    // Check if it's time to update progress (every 3 seconds)
    const currentTime = new Date().getTime();
    if (currentTime - lastProgressUpdate >= progressInterval) {
      console.log(`Progress: ${count} files processed`);
      lastProgressUpdate = currentTime;
    }
    
    // Write batch to spreadsheet when it reaches the batch size
    if (currentBatch.length >= batchSize) {
      console.log(`Writing batch ${batchNumber} to spreadsheet (${currentBatch.length} files)...`);
      sheet.getRange(row, 1, currentBatch.length, headers.length).setValues(currentBatch);
      row += currentBatch.length;
      currentBatch = [];
      batchNumber++;
    }
  }
  
  // Write the final batch if there are any remaining files
  if (currentBatch.length > 0) {
    console.log(`Writing final batch to spreadsheet (${currentBatch.length} files)...`);
    sheet.getRange(row, 1, currentBatch.length, headers.length).setValues(currentBatch);
  }
  
  // Format the sheet
  console.log('Formatting spreadsheet...');
  sheet.autoResizeColumns(1, headers.length);
  
  // Create a CSV file directly from the data instead of converting the sheet
  console.log('Creating CSV file...');
  const csvFile = createCSVDirectly(sheet, headers);
  
  // Add a link to the CSV file in the spreadsheet
  console.log('Adding CSV download link to spreadsheet...');
  sheet.getRange(1, headers.length + 1).setValue('CSV File Link');
  sheet.getRange(2, headers.length + 1).setFormula('=HYPERLINK("' + csvFile.getUrl() + '", "Click here to download CSV")');
  
  console.log(`Process completed. Processed ${count} files owned by you.`);
  console.log(`Spreadsheet URL: ${ss.getUrl()}`);
  console.log(`CSV File URL: ${csvFile.getUrl()}`);
  
  // Try to show UI alert, but don't fail if it's not available
  try {
    SpreadsheetApp.getUi().alert('Process completed: ' + count + ' files processed. Metadata has been saved to a CSV file. Check the link in cell ' + 
      sheet.getRange(2, headers.length + 1).getA1Notation());
  } catch (e) {
    console.log('Note: UI alert could not be shown, but the process completed successfully.');
    console.log('Please check the spreadsheet and CSV file URLs above in the execution log.');
  }
}

/**
 * Counts the total number of files owned by the current user
 * @return {number} The total number of files owned by the current user
 */
function countOwnedFiles() {
  var count = 0;
  var files = DriveApp.searchFiles('"me" in owners');
  while (files.hasNext()) {
    files.next();
    count++;
  }
  Logger.log('Total files owned by me: ' + count);
  return count;
}

/**
 * Creates a CSV file directly from the spreadsheet data in smaller chunks
 * to avoid timeout issues with large datasets
 */
function createCSVDirectly(sheet, headers) {
  console.log('Creating CSV file directly from data...');
  
  // Get the data range
  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();
  
  // Process in smaller chunks to avoid timeout
  const chunkSize = 5000;
  let csvContent = '';
  
  // Add headers
  for (let i = 0; i < headers.length; i++) {
    csvContent += headers[i];
    if (i < headers.length - 1) csvContent += ',';
  }
  csvContent += '\n';
  
  // Process data in chunks
  for (let i = 1; i < numRows; i += chunkSize) {
    const endRow = Math.min(i + chunkSize, numRows);
    console.log(`Processing CSV rows ${i} to ${endRow-1} of ${numRows-1}...`);
    
    const rows = sheet.getRange(i + 1, 1, endRow - i, numCols).getValues();
    
    for (let j = 0; j < rows.length; j++) {
      const row = rows[j];
      for (let k = 0; k < row.length; k++) {
        let cell = row[k];
        // Handle special characters and commas
        if (typeof cell === 'string') {
          cell = cell.replace(/"/g, '""');
          if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
            cell = '"' + cell + '"';
          }
        }
        csvContent += cell;
        if (k < row.length - 1) csvContent += ',';
      }
      csvContent += '\n';
    }
  }
  
  console.log('CSV content created, saving to file...');
  return DriveApp.createFile('google_files_metadata.csv', csvContent, MimeType.CSV);
}

// Add menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Metadata Tools')
    .addItem('Generate Metadata CSV', 'generateMetadataCSV')
    .addItem('Count My Files', 'countOwnedFiles')
    .addToUi();
} 