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
        const value = row[1]; // Assuming column B contains the gender
        return value.toLowerCase() === 'male'; // Example: Highlight rows where value is male
      });

      console.log('Rows to highlight:', rowsToHighlight);

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

      /*
      const requestBody = {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: 0,
                startRowIndex: 1,
                endRowIndex: 2,
                startColumnIndex: 0
              },
              cell: {
                userEnteredFormat: {
                  backgroundColorStyle: {
                    rgbColor:{
                      red: 0.5,
                      green: 1,
                      blue: 1,
                      alpha: 1,
                    }
                  },
                },
              },
              fields: 'userEnteredFormat.backgroundColorStyle',
            },
          },
        ],
      };
      */

      await this.service.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

      console.log('Rows highlighted successfully');
    } catch (error) {
      console.error('Error occurred:', error.message);
    }
  }

  async highlightRows(spreadsheetId, rowIndices, color = { red: 0.5, green: 1, blue: 0.5, alpha: 1 }) {
    try {
      const requests = rowIndices.map((index) => ({
        repeatCell: {
          range: {
            sheetId: 0,
            startRowIndex: index,
            endRowIndex: index + 1,
            startColumnIndex: 0
          },
          cell: {
            userEnteredFormat: {
              backgroundColorStyle: {
                rgbColor: color
              },
            },
          },
          fields: 'userEnteredFormat.backgroundColorStyle',
        },
      }));

      await this.service.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

      console.log('Rows highlighted successfully');
    } catch (error) {
      console.error('Error occurred:', error.message);
    }
  }

  async clearFormatting(spreadsheetId, skipHeaders = true) {
    try {
      const requests = [
        {
          updateCells: {
            range: {
              sheetId: 0,
              startRowIndex: skipHeaders ? 1 : 0,
              startColumnIndex: 0
            },
            fields: 'userEnteredFormat',
          },
        },
      ];

      await this.service.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });

      console.log('Formatting cleared successfully');
    } catch (error) {
      console.error('Error occurred:', error.message);
    }
  }
}

module.exports = GSheets;
