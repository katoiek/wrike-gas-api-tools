/**
 * ネストされたオブジェクトをフラット化する関数
 *
 * @param {Object} obj - フラット化するオブジェクト
 * @param {string} prefix - キーのプレフィックス（再帰呼び出し用）
 * @return {Object} フラット化されたオブジェクト
 */
function flattenObject(obj, prefix = '') {
  let result = {};

  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const value = obj[key];

      if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        // ネストされたオブジェクトを再帰的に処理
        const nested = flattenObject(value, `${prefix}${key}.`);
        result = { ...result, ...nested };
      } else {
        // プリミティブ値または配列の場合
        result[`${prefix}${key}`] = value;
      }
    }
  }
  return result;
}

/**
 * Wrike APIから自分の情報を取得してスプレッドシートに表示する関数
 */
function getInformationAboutMe() {
  try {
    // APIエンドポイントの構築
    const method = '/contacts';
    const parameters = '?me';
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API URLまたはトークンが設定されていません。Wrike認証を実行してください。');
    }

    const apiEndpoint = apiUrl + method + parameters;

    // APIリクエストの実行
    const response = UrlFetchApp.fetch(
      apiEndpoint,
      {
        headers: {
          Authorization: 'Bearer ' + token,
        },
        muteHttpExceptions: true
      }
    );

    // レスポンスのステータスコードをチェック
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`API呼び出しエラー: ステータスコード ${responseCode}, レスポンス: ${response.getContentText()}`);
    }

    // レスポンスの解析
    const responseText = response.getContentText();
    const myJson = JSON.parse(responseText);
    const data = myJson.data;

    if (!data || data.length === 0) {
      throw new Error('APIからデータが返されませんでした。');
    }

    // データのフラット化
    const flattenedData = flattenObject(data);

    // スプレッドシートへの書き込み
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('GetInfoAboutMe');

    if (!resultSheet) {
      throw new Error("'GetInfoAboutMe'シートが見つかりません。シートを作成してください。");
    }

    // 以前のデータをクリアする
    resultSheet.clear();

    // タイトルを設定（A1とB1に設定）
    resultSheet.getRange(1, 1).setValue('Key');
    resultSheet.getRange(1, 2).setValue('Value');

    // タイトル行の書式設定
    resultSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f3f3f3');

    // データの書き込み
    let row = 2;
    for (const key in flattenedData) {
      if (Object.prototype.hasOwnProperty.call(flattenedData, key)) {
        // キーから先頭の "0." を除去（配列データの場合）
        let columnKey = key.startsWith('0.') ? key.slice(2) : key;

        // A列にキー、B列に値を設定
        resultSheet.getRange(row, 1).setValue(columnKey);
        resultSheet.getRange(row, 2).setValue(flattenedData[key]);
        row++;
      }
    }

    // 列の幅を自動調整
    resultSheet.autoResizeColumns(1, 2);

    // 成功メッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast('ユーザー情報の取得が完了しました。', '処理完了');

  } catch (error) {
    // エラーハンドリング
    console.error('エラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラーが発生しました: ' + error.message,
      'エラー',
      10
    );
  }
}
