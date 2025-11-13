/**
 * Convert user IDs from API v2 to v4 format / ユーザーIDをAPI v2からv4形式に変換する
 */
function getUserIdv4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Userlist = ss.getSheetByName('Userlist');
  const useridsRange = Userlist.getRange('B2:B' + Userlist.getLastRow());
  const userids = useridsRange.getValues();
  const method = '/ids';
  const parameters = '?type=ApiV2User&ids=';

  // Object to hold mapping between original and converted IDs / 元のIDと変換後のIDのマッピングを保持するオブジェクト
  // Key: original ID (B column value), Value: converted ID / キー: 元のID (B列の値)、値: 変換後のID
  const idMapping = {};

  // Process in small batches (e.g., 10 at a time) / バッチサイズを小さくして処理する（例: 10件ずつ）
  // This makes it easier to identify correspondence within batches even if API response order is not guaranteed / これにより、APIの応答順序が保証されなくても、バッチ内での対応関係を特定しやすくなる
  const BATCH_SIZE = 10;

  for (let i = 0; i < userids.length; i += BATCH_SIZE) {
      const idBatch = userids.slice(i, i + BATCH_SIZE).map(idArray => idArray[0]).filter(id => id !== '');

      if (idBatch.length > 0) {
          const apiEndpoint = scriptProperties.getProperty('api_url') + method + parameters + '[' + idBatch.join(',') + ']';

          try {
              const response = UrlFetchApp.fetch(apiEndpoint, {
                  headers: {
                      Authorization: 'Bearer ' + scriptProperties.getProperty('token'),
                  },
              });

              if (response.getResponseCode() === 200) {
                  const myJson = JSON.parse(response.getContentText());
                  const data = myJson.data;

                  Logger.log('API Response for batch starting at index ' + i + ': ' + JSON.stringify(data));

                  // Identify converted IDs corresponding to original IDs from API response / APIレスポンスから元のIDに対応する変換後のIDを特定
                  for (const item of data) {
                      if (item.hasOwnProperty('id')) {
                          // Find the optimal method to identify the original ID / 元のIDを特定するための最適な方法を探す
                          let originalId = null;

                          // 1. First, check if API response has a field indicating the original ID / 1. まず、APIレスポンスに元のIDを示すフィールドがあるか確認
                          if (item.hasOwnProperty('apiV2Id')) {
                              originalId = item.apiV2Id;
                          }
                          // 2. Next, try exact ID matching / 2. 次に、IDの完全一致を試みる
                          else {
                              for (const batchId of idBatch) {
                                  // Exact match / 完全一致
                                  if (item.id === batchId) {
                                      originalId = batchId;
                                      break;
                                  }
                              }
                          }
                          // 3. Finally, try partial matching / 3. 最後に、部分一致を試みる
                          if (!originalId) {
                              for (const batchId of idBatch) {
                                  // Partial match (with stricter conditions) / 部分一致（より厳密な条件で）
                                  if (item.id.includes(batchId) || batchId.includes(item.id)) {
                                      originalId = batchId;
                                      break;
                                  }
                              }
                          }

                          // If original ID is identified, add to mapping / 元のIDが特定できた場合、マッピングに追加
                          if (originalId) {
                              idMapping[originalId] = item.id;
                              Logger.log(`Mapped: ${originalId} -> ${item.id}`);
                          } else {
                              Logger.log(`Failed to map ID: ${item.id}`);
                          }
                      }
                  }
              } else {
                  Logger.log('Error: Received response code ' + response.getResponseCode() + ' for batch starting at index ' + i);
              }
          } catch (e) {
              Logger.log('Exception occurred for batch starting at index ' + i + ': ' + e.message);
          }

          // Wait between batch processing (rate limiting countermeasure) / バッチ処理の間隔を空ける（レート制限対策）
          Utilities.sleep(500);
      }
  }

  // Write results while maintaining original order / 元の順序を保持したまま結果を書き込む
  if (Object.keys(idMapping).length > 0) {
      // Create result array based on original order / 元の順序に基づいて結果配列を作成
      const results = [];
      for (let i = 0; i < userids.length; i++) {
          const originalId = userids[i][0];
          if (originalId !== '' && idMapping[originalId] !== undefined) {
              results.push([idMapping[originalId]]);
          } else {
              results.push(['']); // Set empty value if conversion failed / 変換できなかった場合は空の値を設定
          }
      }

      // Write results to spreadsheet / 結果をスプレッドシートに書き込む
      Userlist.getRange(2, 1, results.length, 1).setValues(results);
      Logger.log('Successfully wrote ' + results.length + ' IDs to the spreadsheet.');
      Logger.log('Mapped ' + Object.keys(idMapping).length + ' IDs out of ' + userids.filter(id => id[0] !== '').length + ' non-empty IDs.');
  } else {
      Logger.log('No data to write to the spreadsheet.');
  }
}

