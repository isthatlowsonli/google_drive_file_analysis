# Google Drive and Photos Metadata Extractor

A Google Apps Script tool to extract metadata from files in your Google Drive and Google Photos, saving the information to CSV files.

## Features

- Extracts metadata from all files you own in Google Drive
- Extracts detailed metadata from Google Photos including camera information and location data
- Processes files in batches to handle large datasets efficiently
- Creates formatted spreadsheets with all metadata
- Generates downloadable CSV files
- Provides progress updates during processing
- Includes direct links to original files
- Includes a test mode for processing a small sample of files

## Metadata Fields

### Google Drive Files
The script extracts the following metadata for each Drive file:
- Source (Drive)
- File Name
- File Type (MIME type)
- Owner
- Created Time
- Modified Time
- File Size (in bytes)
- File ID
- Link (direct URL to the file)

### Google Photos
For photos, the script extracts additional metadata:
- All Drive metadata fields
- Camera Make
- Camera Model
- Time Taken
- Location (latitude, longitude)
- Image Dimensions (width, height)
- Image Format

## Setup Instructions

1. Go to [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Copy the contents of the following files into the script editor:
   - `google_drive_metadata.gs` for Drive file metadata
   - `google_photos_metadata.gs` for Photos metadata
   - `example.gs` for testing functionality
4. Enable the Drive Advanced Service:
   - Click on "Services" (+ icon)
   - Find "Drive API" in the list
   - Click "Add" to enable it
5. Save the project with a name like "Google Drive and Photos Metadata Extractor"
6. Refresh the page to see the new menu items

## Usage

1. Open Google Drive
2. You'll see a new menu item called "Metadata Tools"
3. Choose from the following options:
   - "Generate Metadata CSV" to process all Drive files
   - "Generate Photos Metadata CSV" to process all photos
   - "Generate Test CSV (10 files)" to test with a small sample
4. The script will:
   - Count all files you own
   - Process them in batches
   - Create a new spreadsheet with the metadata
   - Generate a CSV file
   - Add a download link to the CSV in the spreadsheet

## Performance Considerations

- The script processes files in batches to avoid memory issues
  - Drive files: 5,000 files per batch
  - Photos: 1,000 photos per batch
- For users with many files (10,000+), the process may take several minutes
- Progress updates are shown every 3 seconds
- The script only processes files you own, not shared files

## Troubleshooting

If you encounter any issues:

- Check the execution logs for error messages
- Make sure you have permission to access the files
- Verify that the Drive Advanced Service is enabled
- Try running the test function first to verify access
- If the script times out, try running it again - it will continue from where it left off

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. See CONTRIBUTING.md for guidelines. 