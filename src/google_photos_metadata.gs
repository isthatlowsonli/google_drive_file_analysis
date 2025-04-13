/**
 * Script to extract metadata from Google Photos
 * 
 * IMPORTANT: Before running this script, you MUST enable the Drive Advanced Service:
 * 1. In the Apps Script editor, click on "Services" (+ icon)
 * 2. Find "Drive API" in the list
 * 3. Click "Add" to enable it
 * 4. Save the script
 * 5. Try running the script again
 */

function generatePhotosMetadataCSV() {
  console.log('Starting Google Photos metadata collection process...');
  
  // Check if Drive Advanced Service is enabled
  try {
    // Test Drive API access
    Drive.Files.list({
      maxResults: 1
    });
  } catch (e) {
    const errorMessage = 'ERROR: Drive Advanced Service is not enabled.\n\n' +
      'To enable it:\n' +
      '1. In the Apps Script editor, click on "Services" (+ icon)\n' +
      '2. Find "Drive API" in the list\n' +
      '3. Click "Add" to enable it\n' +
      '4. Save the script\n' +
      '5. Try running the script again';
    
    console.error(errorMessage);
    try {
      SpreadsheetApp.getUi().alert(errorMessage);
    } catch (e) {
      // If we can't show the UI alert, the error is already logged
    }
    return;
  }
  
  // Create a new spreadsheet
  console.log('Creating new spreadsheet...');
  const ss = SpreadsheetApp.create('Google Photos Metadata');
  const sheet = ss.getActiveSheet();
  
  // Set headers
  console.log('Setting up headers...');
  const headers = [
    'Source',
    'File Name',
    'File Type',
    'Created Time',
    'Modified Time',
    'File Size (bytes)',
    'File ID',
    'Link',
    'Camera Make',
    'Camera Model',
    'Time Taken',
    'Location',
    'Width',
    'Height',
    'Image Format'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get current user's email
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  console.log(`Processing photos for: ${currentUserEmail}`);
  
  // Use a different approach to find photos
  // First, try to find the Google Photos folder
  console.log('Searching for Google Photos...');
  
  // Try to find photos using a more general query
  const query = "mimeType contains 'image/' and trashed = false and '" + currentUserEmail + "' in owners";
  let photoFiles;
  
  try {
    photoFiles = Drive.Files.list({
      q: query,
      fields: 'files(id,name,mimeType,createdTime,modifiedTime,webContentLink,size,imageMediaMetadata),nextPageToken',
      maxResults: 1000
    });
  } catch (e) {
    const errorMessage = 'Error accessing photos: ' + e.message;
    console.error(errorMessage);
    try {
      SpreadsheetApp.getUi().alert(errorMessage);
    } catch (e) {
      // If we can't show the UI alert, the error is already logged
    }
    return;
  }
  
  let count = 0;
  let pageToken;
  
  // Variables for progress tracking
  let lastProgressUpdate = new Date().getTime();
  const progressInterval = 3000; // 3 seconds in milliseconds
  
  // Process files in batches
  console.log('Processing photos (this may take a while)...');
  const batchSize = 1000;
  let currentBatch = [];
  let batchNumber = 1;
  let row = 2;
  
  do {
    if (pageToken) {
      try {
        photoFiles = Drive.Files.list({
          q: query,
          fields: 'files(id,name,mimeType,createdTime,modifiedTime,webContentLink,size,imageMediaMetadata),nextPageToken',
          maxResults: 1000,
          pageToken: pageToken
        });
      } catch (e) {
        console.error('Error fetching next page: ' + e.message);
        break;
      }
    }
    
    if (!photoFiles.files || photoFiles.files.length === 0) {
      console.log('No photos found.');
      break;
    }
    
    for (let i = 0; i < photoFiles.files.length; i++) {
      const file = photoFiles.files[i];
      const metadata = file.imageMediaMetadata || {};
      
      // Get location info if available
      let location = 'N/A';
      if (metadata.location) {
        location = `${metadata.location.latitude}, ${metadata.location.longitude}`;
      }
      
      const fileData = [
        'Photos',
        file.name,
        file.mimeType,
        file.createdTime,
        file.modifiedTime,
        file.size,
        file.id,
        file.webContentLink,
        metadata.cameraMake || 'N/A',
        metadata.cameraModel || 'N/A',
        metadata.time || 'N/A',
        location,
        metadata.width || 'N/A',
        metadata.height || 'N/A',
        metadata.imageFormat || 'N/A'
      ];
      
      currentBatch.push(fileData);
      count++;
      
      // Check if it's time to update progress (every 3 seconds)
      const currentTime = new Date().getTime();
      if (currentTime - lastProgressUpdate >= progressInterval) {
        console.log(`Progress: ${count} photos processed`);
        lastProgressUpdate = currentTime;
      }
      
      // Write batch to spreadsheet when it reaches the batch size
      if (currentBatch.length >= batchSize) {
        console.log(`Writing batch ${batchNumber} to spreadsheet (${currentBatch.length} photos)...`);
        sheet.getRange(row, 1, currentBatch.length, headers.length).setValues(currentBatch);
        row += currentBatch.length;
        currentBatch = [];
        batchNumber++;
      }
    }
    
    pageToken = photoFiles.nextPageToken;
  } while (pageToken);
  
  // Write the final batch if there are any remaining files
  if (currentBatch.length > 0) {
    console.log(`Writing final batch to spreadsheet (${currentBatch.length} photos)...`);
    sheet.getRange(row, 1, currentBatch.length, headers.length).setValues(currentBatch);
  }
  
  if (count === 0) {
    const message = 'No photos were found in your Google Drive.';
    console.log(message);
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {
      // If we can't show the UI alert, the message is already logged
    }
    return;
  }
  
  // Format the sheet
  console.log('Formatting spreadsheet...');
  sheet.autoResizeColumns(1, headers.length);
  
  // Create a CSV file
  console.log('Creating CSV file...');
  const csvFile = createCSVDirectly(sheet, headers);
  
  // Add a link to the CSV file in the spreadsheet
  console.log('Adding CSV download link to spreadsheet...');
  sheet.getRange(1, headers.length + 1).setValue('CSV File Link');
  sheet.getRange(2, headers.length + 1).setFormula('=HYPERLINK("' + csvFile.getUrl() + '", "Click here to download CSV")');
  
  console.log(`Process completed. Processed ${count} photos.`);
  console.log(`Spreadsheet URL: ${ss.getUrl()}`);
  console.log(`CSV File URL: ${csvFile.getUrl()}`);
  
  // Try to show UI alert, but don't fail if it's not available
  try {
    SpreadsheetApp.getUi().alert('Process completed: ' + count + ' photos processed. Metadata has been saved to a CSV file. Check the link in cell ' + 
      sheet.getRange(2, headers.length + 1).getA1Notation());
  } catch (e) {
    console.log('Note: UI alert could not be shown, but the process completed successfully.');
    console.log('Please check the spreadsheet and CSV file URLs above in the execution log.');
  }
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
  return DriveApp.createFile('google_photos_metadata.csv', csvContent, MimeType.CSV);
}

// Add menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Metadata Tools')
    .addItem('Generate Photos Metadata CSV', 'generatePhotosMetadataCSV')
    .addToUi();
} 