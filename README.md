# Excel VBA CSV Exporter

This Excel VBA macro automates the process of exporting data from an active worksheet to a CSV file. The macro identifies the last row and column with data, copies the data into a new workbook, and then saves it as a CSV file with a timestamp.

## Features

- **Automatic Data Range Detection**: The macro identifies the last used row and column with actual data (ignoring empty cells) in the active worksheet.
- **CSV File Generation**: The selected data is copied to a new workbook and saved as a CSV file.
- **Timestamped Filename**: The CSV file is saved with a filename that includes the current worksheet name and a timestamp (in `DD-MMM-YYYY_HH-MM-SS` format).
- **Non-Destructive Process**: The original workbook remains unchanged during the process.

## How It Works

1. **Data Range Detection**: 
   - The macro scans the active worksheet to find the last row and column with data.
   - The data range is defined based on these detected boundaries.

2. **Data Export**:
   - The identified range is copied to a new temporary workbook.
   - The temporary workbook is then saved as a CSV file in the same directory as the original workbook (or the default file path if the workbook is unsaved).

3. **Notification**:
   - Upon completion, the user is notified with a message box displaying the path of the saved CSV file.

## Usage

To use this macro, you can either:
- Directly run it from the VBA editor (`Alt + F11`) in Excel.
- Assign it to a button on your worksheet for easy access.

Ensure that your workbook is saved before running the macro to avoid issues with file paths.
