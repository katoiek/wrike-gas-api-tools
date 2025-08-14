/**
 * Flatten nested objects into a single level object / ネストされたオブジェクトを単一レベルのオブジェクトに平坦化する関数
 *
 * @param {Object} obj - Object to flatten / フラット化するオブジェクト
 * @param {string} prefix - Prefix for keys (for recursive calls) / キーのプレフィックス（再帰呼び出し用）
 * @return {Object} Flattened object / フラット化されたオブジェクト
 */
function flattenObject(obj, prefix = '') {
  let result = {};

  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const value = obj[key];

      if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        // Process nested objects recursively / ネストされたオブジェクトを再帰的に処理
        const nested = flattenObject(value, `${prefix}${key}.`);
        result = { ...result, ...nested };
      } else {
        // For primitive values or arrays / プリミティブ値または配列の場合
        result[`${prefix}${key}`] = value;
      }
    }
  }
  return result;
}

/**
 * Get my information from Wrike API and display in spreadsheet / Wrike APIから自分の情報を取得してスプレッドシートに表示する関数
 */
function getInformationAboutMe() {
  try {
    // Build API endpoint / APIエンドポイントの構築
    const method = '/contacts';
    const parameters = '?me';
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API URLまたはトークンが設定されていません。Wrike認証を実行してください。');
    }

    const apiEndpoint = apiUrl + method + parameters;

    // Execute API request / APIリクエストの実行
    const response = UrlFetchApp.fetch(
      apiEndpoint,
      {
        headers: {
          Authorization: 'Bearer ' + token,
        },
        muteHttpExceptions: true
      }
    );

    // Check response status code / レスポンスのステータスコードをチェック
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`API呼び出しエラー: ステータスコード ${responseCode}, レスポンス: ${response.getContentText()}`);
    }

    // Parse response / レスポンスの解析
    const responseText = response.getContentText();
    const myJson = JSON.parse(responseText);
    const data = myJson.data;

    if (!data || data.length === 0) {
      throw new Error('APIからデータが返されませんでした。');
    }

    // Flatten data / データのフラット化
    const flattenedData = flattenObject(data);

    // Write to spreadsheet / スプレッドシートへの書き込み
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('GetInfoAboutMe');

    if (!resultSheet) {
      throw new Error("'GetInfoAboutMe'シートが見つかりません。シートを作成してください。");
    }

    // Clear previous data / 以前のデータをクリアする
    resultSheet.clear();

    // Set titles (A1 and B1) / タイトルを設定（A1とB1に設定）
    resultSheet.getRange(1, 1).setValue('Key');
    resultSheet.getRange(1, 2).setValue('Value');

    // Format title row / タイトル行の書式設定
    resultSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f3f3f3');

    // Write data / データの書き込み
    let row = 2;
    for (const key in flattenedData) {
      if (Object.prototype.hasOwnProperty.call(flattenedData, key)) {
        // Remove leading "0." from key (for array data) / キーから先頭の "0." を除去（配列データの場合）
        let columnKey = key.startsWith('0.') ? key.slice(2) : key;

        // Set key in column A, value in column B / A列にキー、B列に値を設定
        resultSheet.getRange(row, 1).setValue(columnKey);
        resultSheet.getRange(row, 2).setValue(flattenedData[key]);
        row++;
      }
    }

    // Auto-resize columns / 列の幅を自動調整
    resultSheet.autoResizeColumns(1, 2);

    // Display success message / 成功メッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast('ユーザー情報の取得が完了しました。', '処理完了');

  } catch (error) {
    // Error handling / エラーハンドリング
    console.error('エラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラーが発生しました: ' + error.message,
      'エラー',
      10
    );
  }
}
