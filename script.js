// integrations/googleapi/gsheets.js
const { GoogleAuth } = require('google-auth-library');
const { google } = require('googleapis');

class GSheets {
  constructor() {
    this.auth = new GoogleAuth({
      keyFile: 'credentials.json',
      scopes: 'https://www.googleapis.com/auth/spreadsheets',
    });
    this.service = google.sheets({ version: 'v4', auth: this.auth });
  }

  async getValues(spreadsheetId, range) {
    try {
      const response = await this.service.spreadsheets.values.get({
        spreadsheetId,
        range,
      });
      return response.data.values;
    } catch (error) {
      console.error('The API returned an error:', error.message);
      throw new Error(error.message);
    }
  }

  async updateValues(spreadsheetId, range, values) {
    try {
      const resource = { values };
      const response = await this.service.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'RAW',
        resource,
      });
      return response.data;
    } catch (error) {
      console.error('The API returned an error:', error.message);
      throw new Error(error.message);
    }
  }

  async clear(spreadsheetId, range) {
    try {
      const response = await this.service.spreadsheets.values.clear({
        spreadsheetId,
        range,
      });
      return response.data;
    } catch (error) {
      console.error('The API returned an error:', error.message);
      throw new Error(error.message);
    }
  }

  async readAndHighlight(spreadsheetId, range) {
    try {
      const rows = await this.getValues(spreadsheetId, range);

      const rowsToHighlight = rows.filter(row => {
        // Your calculation or logic here
        const value = parseInt(row[0]); // Assuming column A contains numerical values
        return value > 50; // Example: Highlight rows where value in column A is greater than 50
      });

      const requests = rowsToHighlight.map((row, index) => ({
        repeatCell: {
          range: {
            sheetId: 0, // Adjust the sheet ID if needed
            startRowIndex: index + 1,
            endRowIndex: index + 2,
            startColumnIndex: 0,
            endColumnIndex: 1, // Assuming you want to highlight column A
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 1,
                green: 0,
                blue: 0,
                alpha: 1,
              },
            },
          },
          fields: 'userEnteredFormat(backgroundColor)'
        },
      }));

      await this.service.spreadsheets.batchUpdate({
        spreadsheetId,
        requests,
      });

      console.log('Rows highlighted successfully');
    } catch (error) {
      console.error('Error occurred:', error.message);
    }
  }

  async updateRowColor() {
    const auth = new google.auth.GoogleAuth({
      scopes: [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/spreadsheets',
      ],
    });

    const authClient = await auth.getClient();
    google.options({ auth: authClient });

    const spreadsheetId = 'your-spreadsheet-id';
    const range = 'A1:F1'; // Change this to the row you want to update
    const requestBody = {
      requests: [
        {
          repeatCell: {
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 1.0,
                  green: 0.0,
                  blue: 0.0,
                },
              },
            },
            range: {
              sheetId: 0,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 6,
            },
            fields: 'userEnteredFormat.backgroundColor',
          },
        },
      ],
    };

    sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody,
    }, (err, res) => {
      if (err) return console.error('The API returned an error: ' + err);
      console.log(res.data);
    });
  }

}

module.exports = GSheets;
