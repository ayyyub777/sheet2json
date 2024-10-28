# sheet2json

`sheet2json` is a Node.js package that allows you to easily convert data from Google Sheets into JSON format. With this package, you can fetch spreadsheet data, where column headers are used as keys in the resulting JSON objects, making it simple to work with structured data from Google Sheets.

## Features

- Fetch data from multiple sheets within a single Google Spreadsheet.
- Convert spreadsheet data into JSON format with column headers as keys.
- Handle various data types, including strings, numbers, booleans, and null values.

## Installation

To install `sheet2json`, you can use npm:

```bash
npm install sheet2json
```

## Usage

First, ensure you have a valid access token for the Google Sheets API. You can obtain this token by following the authentication steps in the Google Sheets API documentation.

Here is an example of how to use the `sheet2json` package:

```javascript
const sheet2json = require("sheet2json");

const sheet = new sheet2json(
  "YOUR_ACCESS_TOKEN_HERE"
);

async function main() {
  try {
    const data = await sheet.get("YOUR_SPREADSHEET_ID_HERE");
    console.log(data);
  } catch (error) {
    console.error("Error fetching data:", error);
  }
}

main();
```

### Parameters

- **accessToken**: Your Google API access token, which grants permission to access your Google Sheets data.

### Method

- `get(spreadsheetId: string)`: Fetches spreadsheet data and converts it to JSON format.
  - **spreadsheetId**: The ID of the Google Spreadsheet you want to access.

### Returns

The `get` method returns a promise that resolves to an object containing JSON data for each sheet in the specified spreadsheet. Each key corresponds to a sheet name, and its value is an array of objects representing rows in that sheet, with column headers as keys.

## Example

Assuming you have a Google Spreadsheet with the following structure:

| Name  | Age | Active |
|-------|-----|--------|
| Alice | 30  | true   |
| Bob   | 25  | false  |

Using the above code, the output would look like this:

```json
{
  "Sheet1": [
    { "Name": "Alice", "Age": 30, "Active": true },
    { "Name": "Bob", "Age": 25, "Active": false }
  ]
}
```

## Error Handling

If there is an error accessing the spreadsheet or if the data is invalid, an error will be thrown with a descriptive message. Make sure to handle these errors appropriately in your application.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue to discuss improvements or new features.
