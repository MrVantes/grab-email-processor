# Grab Email Processor

This repository contains a script designed to process emails from a specific sender (such as Grab transaction emails), extract relevant data (such as transaction details and locations), and log the data into a Google Sheets spreadsheet for further tracking and analysis.

## Iterations

- For a single sheet (use Mapping on column) = script-grab.js
- For multiple sheet (use Mapping on SheetName) = mscript-grab.js
- Here is the map to be changed for now its month 12 -> column 11 and so on.

```
const map = new Map([
  [12, 11],
  [1, 20],
  [2, 29],
  [3, 38],
  [4, 47],
]);
```

## Purpose

The script automatically reads Grab email receipts, extracts information such as the total amount paid, location details, and the type of service used (e.g., "Work" or "Leisure"), and stores this data into a Google Spreadsheet. It is particularly useful for those who wish to track Grab transactions in a more automated way without manually logging each transaction.

## Key Features

- **Email Processing**: The script searches for emails from a specified sender (e.g., Grab).
- **Data Extraction**: Extracts key information such as the total amount paid, location, and service type (e.g., "Work" or "Leisure").
- **Logging Data**: Inserts the extracted data into specific rows and columns of a Google Sheet.
- **Timezone Handling**: Adjusts the processing for a specific timezone.
- **Morning and Afternoon Differentiation**: Categorizes transactions into morning or afternoon, based on the time the email was sent.

## Setup Instructions

### The Expected Table

<div align="center">
  <img src="https://github.com/user-attachments/assets/5ae69294-f2a0-4c2d-a717-4635ad5354f0" alt="Expected Table">
</div>

- PIC: Person in Charge

### Prerequisites

- A **Google Cloud** account with access to Google Apps Script.
- A **Google Spreadsheet** where the data will be logged.
- Access to your **Grab receipt emails**.

### Steps to Use

1. **Create a Google Spreadsheet**:
   - Create a new Google Spreadsheet where the data will be logged.
   - Obtain the **spreadsheet ID** (the long string in the URL) and **sheet name** (e.g., 'Grab Calculator').

2. **Open Google Apps Script**:
   - Open [Google Apps Script](https://script.google.com/) and create a new project.
   - Paste the script code into the Apps Script editor.

3. **Update Script with Your Information**:
   - Replace `"your-email@example.com"` with your actual email address.
   - Replace the placeholder `'YOUR_SPREADSHEET_ID'` with the actual ID of your Google Spreadsheet.
   - Replace the placeholder `'YOUR_SHEET_NAME'` with the name of the sheet where data should be logged.

4. **Set Permissions**:
   - Authorize the script to access your Gmail and Google Sheets by following the prompts in Google Apps Script.

5. **Run the Script**:
   - Manually run the script or set a trigger to run the function `getEmailsFromSpecificSenderToday` at a regular interval (e.g., daily).

6. **View the Logged Data**:
   - After running the script, you should see the processed Grab email data in the specified Google Sheet.

- **Timezone**: You can change the timezone by modifying the `timezone` variable in the script.
- **Location Mapping**: Update the `simplifyText` function with additional locations or services if needed.

## Contributions

Feel free to open issues or submit pull requests if you have suggestions for improvements or find bugs. This project is open for contributions to make the email processing and logging even more efficient!

## License

This project is licensed under the MIT License.
