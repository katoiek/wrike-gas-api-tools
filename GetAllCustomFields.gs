/**
 * Flatten nested objects into a single level object / ネストされたオブジェクトを単一レベルのオブジェクトに平坦化する
 * @param {Object} obj - Object to flatten / 平坦化するオブジェクト
 * @param {string} prefix - Prefix for keys / キーのプレフィックス
 * @return {Object} Flattened object / 平坦化されたオブジェクト
 */
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

/**
 * Clear the GetAllCustomFields sheet / GetAllCustomFieldsシートをクリアする
 */
function clearGetAllCustomFieldsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GetAllCustomFields');
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 12).clearContent();  // Clear 12 columns / 12列分をクリア
  }
}

/**
 * Get all custom fields from Wrike API and write to spreadsheet / Wrike APIから全カスタムフィールドを取得してスプレッドシートに書き込む
 */
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

  // Set title row / タイトル行を設定
  headers.forEach((header, index) => {
    resultSheet.getRange(1, index + 1).setValue(header);
  });

  data.forEach((item, index) => {
    const flatItem = flattenObject(item);
    headers.forEach((header, colIndex) => {
      let value = flatItem[header] || null;
      // Display "Account" if spaceId is null or empty / spaceIdがnullまたは空の場合は「アカウント」と表示
      if (header === 'spaceId' && (!value || value === null)) {
        value = 'アカウント';
      }
      resultSheet.getRange(index + 2, colIndex + 1).setValue(value);
    });
  });
}
