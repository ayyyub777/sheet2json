import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";

export interface SheetData
  extends Array<Array<string | number | boolean | null>> {}

export interface AllSheetData {
  [sheetName: string]: SheetData;
}

export interface JsonSheetData {
  [sheetName: string]: Record<string, string | number | boolean | null>[];
}

class sheet2json {
  private oauth2Client: OAuth2Client;

  constructor(accessToken: string) {
    this.oauth2Client = new google.auth.OAuth2();
    this.oauth2Client.setCredentials({ access_token: accessToken });
  }

  private get sheetsClient() {
    return google.sheets({ version: "v4", auth: this.oauth2Client });
  }

  /**
   * Fetches spreadsheet data and converts it to JSON with column headers as keys
   * @param spreadsheetId - The ID of the Google Spreadsheet
   * @returns Object containing JSON data for each sheet
   * @throws {Error} If spreadsheet access fails or data is invalid
   */
  async get(spreadsheetId: string): Promise<JsonSheetData> {
    try {
      const rawData = await this.getRawSpreadsheetData(spreadsheetId);
      return this.convertToJson(rawData);
    } catch (error) {
      throw new Error(`Failed to get JSON data: ${(error as Error).message}`);
    }
  }

  /**
   * Fetches raw spreadsheet data
   * @private
   */
  private async getRawSpreadsheetData(
    spreadsheetId: string
  ): Promise<AllSheetData> {
    try {
      // Get all sheet names
      const metadata = await this.sheetsClient.spreadsheets.get({
        spreadsheetId,
      });

      const sheetNames =
        metadata.data.sheets?.map(
          (sheet) => sheet.properties?.title as string
        ) || [];

      // Fetch data from all sheets in parallel
      const allData: AllSheetData = {};
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
      throw new Error(
        `Failed to get spreadsheet data: ${(error as Error).message}`
      );
    }
  }

  /**
   * Converts raw spreadsheet data to JSON format
   * @private
   */
  private convertToJson(spreadsheetData: AllSheetData): JsonSheetData {
    if (
      !spreadsheetData ||
      typeof spreadsheetData !== "object" ||
      Object.keys(spreadsheetData).length === 0
    ) {
      return {};
    }

    const result: JsonSheetData = {};

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
            if (headers[index] === "string") {
              obj[headers[index]] = this.parseValue(value);
            }
            return obj;
          }, {} as Record<string, string | number | boolean | null>);
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
  private validateHeaders(
    headers: Array<string | number | boolean | null>
  ): boolean {
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
  private parseValue(
    value: string | number | boolean | null
  ): string | number | boolean | null {
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

export default sheet2json;
