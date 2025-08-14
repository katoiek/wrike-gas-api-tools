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
 * Clear the GetSpaceList sheet / GetSpaceListシートをクリアする
 */
function clearGetSpaceListSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GetSpaceList'); // Get sheet named 'GetSpaceList' / 'GetSpaceList'という名前のシートを取得
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 12).clearContent();  // Clear 12 columns / 12列分をクリア
  }
}

/**
 * Get space list from Wrike API and write to spreadsheet / Wrike APIからスペース一覧を取得してスプレッドシートに書き込む
 */
function getSpaceList() {
  clearGetSpaceListSheet();

  var method = '/spaces';
  var parameters =''
  var apiEndpoint = scriptProperties.getProperty('api_url') + method + parameters;

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
  const resultSheet = ss.getSheetByName('GetSpaceList'); // Get sheet named 'GetSpaceList' / 'GetSpaceList'という名前のシートを取得

  const headers = ['id', 'title', 'accessType', 'archived', 'defaultProjectWorkflowId', 'defaultTaskWorkflowId'];
  // const headers = ['id', 'accountId', 'title', 'createdDate', 'updatedDate', 'description', 'sharedIds',
  //                  'parentIds', 'childIds', 'scope', 'permalink', 'workflowId'];

  // Set title row / タイトル行を設定
  headers.forEach((header, index) => {
    // Columns start from 1, not 0, so add 1 to index / 列は0から始まらず1から始まるため、indexに1を足します
    resultSheet.getRange(1, index + 1).setValue(header);
  });

  data.forEach((item, index) => {
    const flatItem = flattenObject(item);
    headers.forEach((header, colIndex) => {
      // Row number adds 2 to array index considering title row (JavaScript starts from 0, skip title row) / 行番号はタイトル行を考慮して配列のindexに2を足します(JavaScriptは0から始まり、タイトル行をスキップするため)
      // Column number starts from 1, not 0, so add 1 to colIndex / 列番号は0から始まらず1から始まるため、colIndexに1を足します
      resultSheet.getRange(index + 2, colIndex + 1).setValue(flatItem[header] || null);
    });
  });
}
