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

  // spaceIdでデータをグループ化
  const groupedData = {};
  data.forEach(item => {
    const flatItem = flattenObject(item);
    const spaceId = flatItem.spaceId || 'undefined';

    if (!groupedData[spaceId]) {
      groupedData[spaceId] = [];
    }
    groupedData[spaceId].push(flatItem);
  });

  let currentRow = 2;

  // spaceIdごとにデータを出力
  Object.keys(groupedData).sort().forEach(spaceId => {
    // spaceIdのヘッダー行を追加
    resultSheet.getRange(currentRow, 1).setValue(`${spaceId}`);
    resultSheet.getRange(currentRow, 1, 1, headers.length).setBackground('#E6F3FF');
    resultSheet.getRange(currentRow, 1, 1, headers.length).setFontWeight('bold');
    currentRow++;

    // このspaceIdのデータを出力
    groupedData[spaceId].forEach(flatItem => {
      headers.forEach((header, colIndex) => {
        resultSheet.getRange(currentRow, colIndex + 1).setValue(flatItem[header] || null);
      });
      currentRow++;
    });

    // 空行を追加してグループを区切る
    currentRow++;
  });
}
