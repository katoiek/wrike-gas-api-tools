/**
 * BacklogのデータをWrikeのタスクとしてインポートするスクリプト
 *
 * 前提条件:
 * - BacklogDataという名前のスプレッドシートが存在する
 * - B1セルにWrikeフォルダのパーマリンクが設定されている
 * - scriptPropertiesに'api_url'と'token'が設定されている
 */

/**
 * メイン関数：BacklogデータをWrikeにインポート
 */
function ImportBacklogToWrike() {
  try {
    // スプレッドシートの取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    if (!backlogSheet) {
      throw new Error('BacklogDataシートが見つかりません');
    }

    // B1セルからパーマリンクを取得してフォルダIDを抽出
    const permalink = backlogSheet.getRange('B1').getValue();
    const folderId = extractFolderIdFromPermalink(permalink);

    if (!folderId) {
      throw new Error('B1セルのパーマリンクからフォルダIDを抽出できませんでした');
    }

    // データの取得（ヘッダー行とデータラベル行を除く）
    const lastRow = backlogSheet.getLastRow();
    const lastCol = backlogSheet.getLastColumn();

    if (lastRow <= 2) {
      throw new Error('インポートするデータがありません');
    }

    // データラベル行（2行目）を取得してコメント列のインデックスを特定
    const headers = backlogSheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const titleIndex = headers.indexOf('件名');
    const descriptionIndex = headers.indexOf('詳細');
    const startDateIndex = headers.indexOf('開始日');
    const dueDateIndex = headers.indexOf('期限日');

    if (titleIndex === -1) {
      throw new Error('件名列が見つかりません');
    }

    // コメント列のインデックスを取得
    const commentIndices = [];
    for (let i = 1; i <= 200; i++) {
      const commentIndex = headers.indexOf(`コメント${i}`);
      if (commentIndex !== -1) {
        commentIndices.push(commentIndex);
      }
    }

    // データ行を取得（3行目から開始）
    const dataRange = backlogSheet.getRange(3, 1, lastRow - 2, lastCol);
    const data = dataRange.getValues();

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    // 各行をWrikeタスクとして作成
    for (let i = 0; i < data.length; i++) {
      try {
        const row = data[i];
        const title = row[titleIndex] || '';
        let description = row[descriptionIndex] || '';

        // タイトルが空の場合はスキップ
        if (!title.trim()) {
          console.log(`行 ${i + 3}: タイトルが空のためスキップしました`);
          continue;
        }

        // コメントを収集して連結
        const comments = [];
        commentIndices.forEach((index, i) => {
          const comment = row[index];
          if (comment && comment.toString().trim()) {
            const commentNumber = i + 1;
            const formattedComment = `<strong>【コメント${commentNumber}】</strong><br />${comment.toString().trim()}`;
            comments.push(formattedComment);
          }
        });

        // コメントがある場合は詳細の最後に追加
        if (comments.length > 0) {
          const commentsText = comments.join('<br /><br />');
          if (description.trim()) {
            description += '<br /><br />' + commentsText;
          } else {
            description = commentsText;
          }
        }

        // \\nを<br />に置き換え
        description = description.replace(/\\\\n/g, '<br />');

        // 開始日と期限日を取得
        const startDate = startDateIndex !== -1 ? row[startDateIndex] : undefined;
        const dueDate = dueDateIndex !== -1 ? row[dueDateIndex] : undefined;

        // Wrikeタスクを作成
        const taskData = createWrikeTask(folderId, title, description, startDate, dueDate);

        if (taskData) {
          successCount++;
          console.log(`行 ${i + 3}: タスク作成成功 - ${title}`);
        } else {
          errorCount++;
          errors.push(`行 ${i + 3}: タスク作成失敗 - ${title}`);
        }

        // API制限を考慮して少し待機
        Utilities.sleep(100);

      } catch (error) {
        errorCount++;
        errors.push(`行 ${i + 3}: エラー - ${error.message}`);
        console.error(`行 ${i + 3}でエラー:`, error);
      }
    }

    // 結果をログ出力
    console.log(`インポート完了: 成功 ${successCount}件, 失敗 ${errorCount}件`);
    if (errors.length > 0) {
      console.log('エラー詳細:');
      errors.forEach(error => console.log(error));
    }

    // 結果をアラートで表示
    const message = `インポート完了\n成功: ${successCount}件\n失敗: ${errorCount}件`;
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    console.error('インポート処理でエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * パーマリンクからフォルダIDを取得（Wrike APIを使用）
 * @param {string} permalink - Wrikeフォルダのパーマリンク
 * @return {string|null} フォルダID
 */
function extractFolderIdFromPermalink(permalink) {
  if (!permalink || typeof permalink !== 'string') {
    return null;
  }

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API URLまたはトークンが設定されていません');
    }

    // パーマリンクをURLエンコード
    const encodedPermalink = encodeURIComponent(permalink);
    const endpoint = `${apiUrl}/folders?permalink=${encodedPermalink}`;

    // API リクエスト
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseText);

      // データが存在し、最初のフォルダのIDを返す
      if (responseData.data && responseData.data.length > 0) {
        return responseData.data[0].id;
      } else {
        console.error('パーマリンクに対応するフォルダが見つかりませんでした');
        return null;
      }
    } else {
      console.error(`フォルダID取得API エラー (${responseCode}):`, responseText);
      return null;
    }

  } catch (error) {
    console.error('フォルダID取得エラー:', error);
    return null;
  }
}

/**
 * 日付文字列を変換（YYYY/MM/DD → YYYY-MM-DD）
 * @param {string|Date} dateValue - 日付値
 * @return {string|null} YYYY-MM-DD形式の日付文字列、無効な場合はnull
 */
function convertDateFormat(dateValue) {
  if (!dateValue) {
    return null;
  }

  let dateStr = '';
  
  if (dateValue instanceof Date) {
    const year = dateValue.getFullYear();
    const month = String(dateValue.getMonth() + 1).padStart(2, '0');
    const day = String(dateValue.getDate()).padStart(2, '0');
    dateStr = `${year}-${month}-${day}`;
  } else {
    dateStr = dateValue.toString().trim();
    
    if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(dateStr)) {
      const parts = dateStr.split('/');
      const year = parts[0];
      const month = String(parts[1]).padStart(2, '0');
      const day = String(parts[2]).padStart(2, '0');
      dateStr = `${year}-${month}-${day}`;
    } else {
      return null;
    }
  }
  
  return dateStr;
}

/**
 * Wrikeタスクを作成
 * @param {string} folderId - 作成先フォルダID
 * @param {string} title - タスクタイトル
 * @param {string} description - タスク詳細
 * @param {string|Date} startDate - 開始日（YYYY/MM/DD形式またはDate、オプション）
 * @param {string|Date} dueDate - 期限日（YYYY/MM/DD形式またはDate、オプション）
 * @return {Object|null} 作成されたタスクデータ
 */
function createWrikeTask(folderId, title, description, startDate = undefined, dueDate = undefined) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API URLまたはトークンが設定されていません');
    }

    const endpoint = `${apiUrl}/folders/${folderId}/tasks`;

    // リクエストペイロード
    const payload = {
      title: title,
      description: description
    };

    // 開始日と期限日を変換して追加
    const convertedStartDate = convertDateFormat(startDate);
    const convertedDueDate = convertDateFormat(dueDate);

    if (convertedStartDate && convertedDueDate) {
      payload.dates = {
        start: convertedStartDate,
        due: convertedDueDate
      };
    } else if (convertedStartDate && !convertedDueDate) {
      payload.dates = {
        start: convertedStartDate,
        due: convertedStartDate
      };
    } else if (!convertedStartDate && convertedDueDate) {
      payload.dates = {
        due: convertedDueDate
      };
    }

    // API リクエスト
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseText);
      return responseData.data && responseData.data[0] ? responseData.data[0] : responseData;
    } else {
      console.error(`API エラー (${responseCode}):`, responseText);
      throw new Error(`API エラー: ${responseCode} - ${responseText}`);
    }

  } catch (error) {
    console.error('Wrikeタスク作成エラー:', error);
    throw error;
  }
}

/**
 * テスト用関数：単一タスクの作成テスト
 */
function testCreateSingleTask() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    if (!backlogSheet) {
      throw new Error('BacklogDataシートが見つかりません');
    }

    const permalink = backlogSheet.getRange('B1').getValue();
    const folderId = extractFolderIdFromPermalink(permalink);

    if (!folderId) {
      throw new Error('フォルダIDを抽出できませんでした');
    }

    const testTitle = 'テストタスク - ' + new Date().toLocaleString();
    const testDescription = 'これはテスト用のタスクです。<br />改行テストも含まれています。';

    const result = createWrikeTask(folderId, testTitle, testDescription);

    if (result) {
      console.log('テストタスク作成成功:', result);
      SpreadsheetApp.getUi().alert('テストタスクの作成に成功しました');
    }

  } catch (error) {
    console.error('テストタスク作成エラー:', error);
    SpreadsheetApp.getUi().alert(`テストタスク作成エラー: ${error.message}`);
  }
}

/**
 * 設定確認用関数
 */
function checkConfiguration() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    let message = '設定確認結果:\n';
    message += `API URL: ${apiUrl ? '設定済み' : '未設定'}\n`;
    message += `Token: ${token ? '設定済み' : '未設定'}\n`;
    message += `BacklogDataシート: ${backlogSheet ? '存在' : '存在しない'}\n`;

    if (backlogSheet) {
      const permalink = backlogSheet.getRange('B1').getValue();
      const folderId = extractFolderIdFromPermalink(permalink);
      message += `B1セルのパーマリンク: ${permalink ? '設定済み' : '未設定'}\n`;
      message += `抽出されたフォルダID: ${folderId || '抽出失敗'}`;
    }

    console.log(message);
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    console.error('設定確認エラー:', error);
    SpreadsheetApp.getUi().alert(`設定確認エラー: ${error.message}`);
  }
}
