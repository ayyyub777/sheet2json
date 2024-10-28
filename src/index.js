const { google } = require("googleapis");
const { OAuth2Client } = require("google-auth-library");

class Sheet2Json {
  constructor(accessToken) {
    this.oauth2Client = new google.auth.OAuth2();
    this.oauth2Client.setCredentials({ access_token: accessToken });
  }

  get sheetsClient() {
    return google.sheets({ version: "v4", auth: this.oauth2Client });
  }

  /**
   * Fetches spreadsheet data and converts it to JSON with column headers as keys
   * @param {string} spreadsheetId - The ID of the Google Spreadsheet
   * @returns {Promise<Object>} Object containing JSON data for each sheet
   * @throws {Error} If spreadsheet access fails or data is invalid
   */
  async get(spreadsheetId) {
    try {
      const rawData = await this.getRawSpreadsheetData(spreadsheetId);
      return this.convertToJson(rawData);
    } catch (error) {
      throw new Error(`Failed to get JSON data: ${error.message}`);
    }
  }

  /**
   * Fetches raw spreadsheet data
   * @private
   */
  async getRawSpreadsheetData(spreadsheetId) {
    try {
      // Get all sheet names
      const metadata = await this.sheetsClient.spreadsheets.get({
        spreadsheetId,
      });

      const sheetNames =
        metadata.data.sheets?.map((sheet) => sheet.properties?.title) || [];

      // Fetch data from all sheets in parallel
      const allData = {};
      await Promise.all(
        sheetNames.map(async (sheetName) => {
          const response = await this.sheetsClient.spreadsheets.values.get({
            spreadsheetId,
            range: sheetName,
          });
          allData[sheetName] = response.data.values || [];
        })
      );

      return allData;
    } catch (error) {
      throw new Error(`Failed to get spreadsheet data: ${error.message}`);
    }
  }

  /**
   * Converts raw spreadsheet data to JSON format
   * @private
   */
  convertToJson(spreadsheetData) {
    if (
      !spreadsheetData ||
      typeof spreadsheetData !== "object" ||
      Object.keys(spreadsheetData).length === 0
    ) {
      return {};
    }

    const result = {};

    for (const sheet in spreadsheetData) {
      const data = spreadsheetData[sheet];
      if (Array.isArray(data) && data.length >= 2) {
        const [headers, ...rows] = data;

        // Validate headers are strings and unique
        const validHeaders = this.validateHeaders(headers);
        if (!validHeaders) {
          throw new Error(`Invalid or duplicate headers in sheet: ${sheet}`);
        }

        result[sheet] = rows.map((row) => {
          return row.reduce((obj, value, index) => {
            if (typeof headers[index] === "string") {
              obj[headers[index]] = this.parseValue(value);
            }
            return obj;
          }, {});
        });
      } else {
        result[sheet] = [];
      }
    }

    return result;
  }

  /**
   * Validates that headers are unique strings
   * @private
   */
  validateHeaders(headers) {
    if (!Array.isArray(headers)) return false;

    const seen = new Set();
    for (const header of headers) {
      if (typeof header !== "string" || header.trim() === "") return false;
      if (seen.has(header)) return false;
      seen.add(header);
    }

    return true;
  }

  /**
   * Parses cell values into appropriate types
   * @private
   */
  parseValue(value) {
    if (value === null || value === undefined) return null;

    // Try to parse numbers
    if (typeof value === "string") {
      if (/^\d+$/.test(value)) return parseInt(value, 10);
      if (/^\d*\.\d+$/.test(value)) return parseFloat(value);
      if (value.toLowerCase() === "true") return true;
      if (value.toLowerCase() === "false") return false;
    }

    return value;
  }
}

module.exports = Sheet2Json;
