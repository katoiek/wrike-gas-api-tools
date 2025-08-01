function flattenObject(obj, prefix = '') {
  let result = {};

  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      const value = obj[key];

      if (typeof value === 'object' && !Array.isArray(value)) {
        const nested = flattenObject(value, `${prefix}${key}.`);
        result = { ...result, ...nested };
      } else {
        result[`${prefix}${key}`] = value;
      }
    }
  }

  return result;
}

function clearGetAllCustomFieldsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GetAllCustomFields');
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 12).clearContent();
  }
}

function getAllCustomFields() {
  clearGetAllCustomFieldsSheet();

  var method = '/customfields';
  var apiEndpoint = scriptProperties.getProperty('api_url') + method;

  const response = UrlFetchApp.fetch(
    apiEndpoint,
    {
      headers: {
        Authorization: 'Bearer ' + scriptProperties.getProperty('token'),
      },
    }
  );
  const myJson = JSON.parse(response.getContentText());
  const data = myJson.data;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('GetAllCustomFields');

  const headers = ['id', 'accountId', 'title', 'type', 'spaceId', 'sharedIds', 'scope',
                   'currency', 'aggregation', 'decimalPlaces', 'useThousandsSeparator', 'readOnly'];

  headers.forEach((header, index) => {
    resultSheet.getRange(1, index + 1).setValue(header);
  });

  data.forEach((item, index) => {
    const flatItem = flattenObject(item);
    headers.forEach((header, colIndex) => {
      let value = flatItem[header] || null;
      // spaceIdがnullまたは空の場合は「アカウント」と表示
      if (header === 'spaceId' && (!value || value === null)) {
        value = 'アカウント';
      }
      resultSheet.getRange(index + 2, colIndex + 1).setValue(value);
    });
  });
}
