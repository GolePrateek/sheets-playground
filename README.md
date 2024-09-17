## sheets-playground

This repository provides a playground for interacting with Google Sheets using the Google Sheets API. It demonstrates how to read, write, and manipulate data in spreadsheets, including basic operations like:

- **Reading data:**  Retrieving data from a specified range of cells.
- **Writing data:** Updating the values of cells within a spreadsheet.
- **Clearing data:** Removing all data from a specified range of cells.
- **Highlighting rows:** Applying background colors to specific rows based on a defined condition.
- **Clearing formatting:** Removing all custom formatting applied to cells.

### Prerequisites

Before using the code, you'll need to:

1. **Enable the Google Sheets API:**
    - Go to the Google Cloud Platform Console ([https://console.cloud.google.com/](https://console.cloud.google.com/)).
    - Navigate to "APIs & Services" > "Library" and search for "Google Sheets API."
    - Enable the API.

2. **Create API credentials:**
    - In the Google Cloud Platform Console, navigate to "APIs & Services" > "Credentials" and create an API key.
    - Download the credentials file as a JSON file ("credentials.json").
    - **Important:** Replace the placeholder "credentials.json" in the code with the actual file path to your credentials file.

### Usage

1. **Set up the project:**
    - Clone this repository: `git clone https://github.com/your-username/sheets-playground.git`
    - Navigate to the project directory: `cd sheets-playground`

2. **Install dependencies:**
    - `npm install`

3. **Run the code:**
    - Execute the desired script:
      - `node highlight.js` (To highlight rows based on a specific condition)
      - `node script.js` (To update row color)

### Example

The `highlight.js` script demonstrates how to read a spreadsheet and highlight rows based on a value in the 'Gender' column:

1. **Read data:** It retrieves all rows from the spreadsheet.
2. **Filter rows:** It filters the rows to find those where the 'Gender' value is "Male."
3. **Get row indices:** It extracts the indices of the filtered rows.
4. **Clear formatting:** It removes any existing custom formatting from the spreadsheet.
5. **Highlight rows:** It applies a background color to the filtered rows.

### Customization

You can customize the code to:

- **Change spreadsheet ID:** Replace `'1CicXW85QnB71syn-HgzWj-UypJenWR6pdIRTvnUFaX0'` with your actual spreadsheet ID.
- **Modify the range:** Specify a different range of cells to be processed.
- **Adjust the highlighting logic:** Change the condition in the `filter()` function to target different rows.
- **Apply different formatting:** Experiment with different background colors, text styles, or other formatting options.

### Notes

- The provided scripts require `credentials.json` to authenticate with the Google Sheets API.
- This project is a basic example and can be extended to implement more complex operations.
- The code assumes that the spreadsheet has at least one sheet, and the `sheetId` is set to 0, which refers to the first sheet. Adjust this if your spreadsheet has multiple sheets.
- Refer to the Google Sheets API documentation for more advanced operations and possibilities: [https://developers.google.com/sheets/api/](https://developers.google.com/sheets/api/)

This playground provides a starting point for working with the Google Sheets API. You can adapt and expand the code to fit your specific needs and create custom applications that automate tasks and streamline your spreadsheet workflows.
