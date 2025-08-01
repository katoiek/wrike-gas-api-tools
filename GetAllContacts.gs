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

function clearGetAllContactsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GetAllContacts');
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 12).clearContent();
  }
}

function GetAllContacts() {
  clearGetAllContactsSheet();

  var method = '/contacts';
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
  const resultSheet = ss.getSheetByName('GetAllContacts');

  const headers = ['id', 'firstName', 'lastName', 'type', 'profiles', 'avatarUrl', 'timezone',
                   'locale', 'deleted', 'title', 'primaryEmail'];

  headers.forEach((header, index) => {
    resultSheet.getRange(1, index + 1).setValue(header);
  });

  data.forEach((item, index) => {
    const flatItem = flattenObject(item);
    headers.forEach((header, colIndex) => {
      resultSheet.getRange(index + 2, colIndex + 1).setValue(flatItem[header] || null);
    });
  });
}
