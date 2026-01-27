/**
 * Script to import Backlog data as Wrike tasks / BacklogのデータをWrikeのタスクとしてインポートするスクリプト
 *
 * Prerequisites: / 前提条件:
 * - A spreadsheet named BacklogData exists / BacklogDataという名前のスプレッドシートが存在する
 * - Wrike folder permalink is set in cell B1 / B1セルにWrikeフォルダのパーマリンクが設定されている
 * - 'api_url' and 'token' are set in scriptProperties / scriptPropertiesに'api_url'と'token'が設定されている
 */

/**
 * Main function: Import Backlog data to Wrike / メイン関数：BacklogデータをWrikeにインポート
 */
function ImportBacklogToWrike() {
  try {
    // Get spreadsheet / スプレッドシートの取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    if (!backlogSheet) {
      throw new Error('BacklogData sheet not found / BacklogDataシートが見つかりません');
    }

    // Retrieve permalink from cell B1 and extract folder ID / B1セルからパーマリンクを取得してフォルダIDを抽出
    const permalink = backlogSheet.getRange('B1').getValue();
    const folderId = extractFolderIdFromPermalink(permalink);

    if (!folderId) {
      throw new Error('Failed to extract folder ID from permalink in cell B1 / B1セルのパーマリンクからフォルダIDを抽出できませんでした');
    }

    // Get data (excluding header row and data label row) / データの取得（ヘッダー行とデータラベル行を除く）
    const lastRow = backlogSheet.getLastRow();
    const lastCol = backlogSheet.getLastColumn();

    if (lastRow <= 2) {
      throw new Error('No data to import / インポートするデータがありません');
    }

    // Get data label row (2nd row) and identify index of comment columns / データラベル行（2行目）を取得してコメント列のインデックスを特定
    const headers = backlogSheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const titleIndex = headers.indexOf('件名');
    const descriptionIndex = headers.indexOf('詳細');
    const startDateIndex = headers.indexOf('開始日');
    const dueDateIndex = headers.indexOf('期限日');

    if (titleIndex === -1) {
      throw new Error('Subject column not found / 件名列が見つかりません');
    }

    // Get indices for comment columns / コメント列のインデックスを取得
    const commentIndices = [];
    for (let i = 1; i <= 200; i++) {
      const commentIndex = headers.indexOf(`コメント${i}`);
      if (commentIndex !== -1) {
        commentIndices.push(commentIndex);
      }
    }

    // Get data rows (starting from 3rd row) / データ行を取得（3行目から開始）
    const dataRange = backlogSheet.getRange(3, 1, lastRow - 2, lastCol);
    const data = dataRange.getValues();

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    // Create each row as a Wrike task / 各行をWrikeタスクとして作成
    for (let i = 0; i < data.length; i++) {
      try {
        const row = data[i];
        const title = row[titleIndex] || '';
        let description = row[descriptionIndex] || '';

        // Skip if title is empty / タイトルが空の場合はスキップ
        if (!title.trim()) {
          console.log(`Row / 行 ${i + 3}: Skipped because title is empty / タイトルが空のためスキップしました`);
          continue;
        }

        // Collect and concatenate comments / コメントを収集して連結
        const comments = [];
        commentIndices.forEach((index, i) => {
          const comment = row[index];
          if (comment && comment.toString().trim()) {
            const commentNumber = i + 1;
            const formattedComment = `<strong>【Comment / コメント ${commentNumber}】</strong><br />${comment.toString().trim()}`;
            comments.push(formattedComment);
          }
        });

        // Add comments to the end of description if present / コメントがある場合は詳細の最後に追加
        if (comments.length > 0) {
          const commentsText = comments.join('<br /><br />');
          if (description.trim()) {
            description += '<br /><br />' + commentsText;
          } else {
            description = commentsText;
          }
        }

        // Replace \\n with <br /> / \\nを<br />に置き換え
        description = description.replace(/\\\\n/g, '<br />');

        // Get start date and due date / 開始日と期限日を取得
        const startDate = startDateIndex !== -1 ? row[startDateIndex] : undefined;
        const dueDate = dueDateIndex !== -1 ? row[dueDateIndex] : undefined;

        // Create Wrike task / Wrikeタスクを作成
        const taskData = createWrikeTask(folderId, title, description, startDate, dueDate);

        if (taskData) {
          successCount++;
          console.log(`Row / 行 ${i + 3}: Task creation successful / タスク作成成功 - ${title}`);
        } else {
          errorCount++;
          errors.push(`Row / 行 ${i + 3}: Task creation failed / タスク作成失敗 - ${title}`);
        }

        // Wait a bit considering API limits / API制限を考慮して少し待機
        Utilities.sleep(100);

      } catch (error) {
        errorCount++;
        errors.push(`Row / 行 ${i + 3}: Error / エラー - ${error.message}`);
        console.error(`Error on row / 行 ${i + 3}でエラー:`, error);
      }
    }

    // Log results / 結果をログ出力
    console.log(`Import complete: / インポート完了: Success / 成功 ${successCount} items, Failure / 失敗 ${errorCount} items`);
    if (errors.length > 0) {
      console.log('Error Details: / エラー詳細:');
      errors.forEach(error => console.log(error));
    }

    // Display results in alert / 結果をアラートで表示
    const message = `Import complete / インポート完了\nSuccess / 成功: ${successCount} items\nFailure / 失敗: ${errorCount} items`;
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    console.error('An error occurred during import process: / インポート処理でエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert(`An error occurred: / エラーが発生しました: ${error.message}`);
  }
}

/**
 * Get folder ID from permalink (using Wrike API) / パーマリンクからフォルダIDを取得（Wrike APIを使用）
 * @param {string} permalink - Permalink of Wrike folder / Wrikeフォルダのパーマリンク
 * @return {string|null} Folder ID / フォルダID
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
      throw new Error('API URL or token not set / API URLまたはトークンが設定されていません');
    }

    // URL encode the permalink / パーマリンクをURLエンコード
    const encodedPermalink = encodeURIComponent(permalink);
    const endpoint = `${apiUrl}/folders?permalink=${encodedPermalink}`;

    // API Request / API リクエスト
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

      // Return ID of the first folder if data exists / データが存在し、最初のフォルダのIDを返す
      if (responseData.data && responseData.data.length > 0) {
        return responseData.data[0].id;
      } else {
        console.error('Folder corresponding to permalink not found / パーマリンクに対応するフォルダが見つかりませんでした');
        return null;
      }
    } else {
      console.error(`Folder ID retrieval API error / フォルダID取得API エラー (${responseCode}):`, responseText);
      return null;
    }

  } catch (error) {
    console.error('Folder ID retrieval error: / フォルダID取得エラー:', error);
    return null;
  }
}

/**
 * Convert date string (YYYY/MM/DD → YYYY-MM-DD) / 日付文字列を変換（YYYY/MM/DD → YYYY-MM-DD）
 * @param {string|Date} dateValue - Date value / 日付値
 * @return {string|null} Date string in YYYY-MM-DD format, or null if invalid / YYYY-MM-DD形式の日付文字列、無効な場合はnull
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
 * Create Wrike task / Wrikeタスクを作成
 * @param {string} folderId - Destination folder ID / 作成先フォルダID
 * @param {string} title - Task title / タスクタイトル
 * @param {string} description - Task description / タスク詳細
 * @param {string|Date} startDate - Start date (YYYY/MM/DD format or Date, optional) / 開始日（YYYY/MM/DD形式またはDate、オプション）
 * @param {string|Date} dueDate - Due date (YYYY/MM/DD format or Date, optional) / 期限日（YYYY/MM/DD形式またはDate、オプション）
 * @return {Object|null} Created task data / 作成されたタスクデータ
 */
function createWrikeTask(folderId, title, description, startDate = undefined, dueDate = undefined) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API URL or token not set / API URLまたはトークンが設定されていません');
    }

    const endpoint = `${apiUrl}/folders/${folderId}/tasks`;

    // Request payload / リクエストペイロード
    const payload = {
      title: title,
      description: description
    };

    // Convert and add start and due dates / 開始日と期限日を変換して追加
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

    // API Request / API リクエスト
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
      console.error(`API Error / API エラー (${responseCode}):`, responseText);
      throw new Error(`API Error: / API エラー: ${responseCode} - ${responseText}`);
    }

  } catch (error) {
    console.error('Wrike task creation error: / Wrikeタスク作成エラー:', error);
    throw error;
  }
}

/**
 * Test function: Test creation of a single task / テスト用関数：単一タスクの作成テスト
 */
function testCreateSingleTask() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    if (!backlogSheet) {
      throw new Error('BacklogData sheet not found / BacklogDataシートが見つかりません');
    }

    const permalink = backlogSheet.getRange('B1').getValue();
    const folderId = extractFolderIdFromPermalink(permalink);

    if (!folderId) {
      throw new Error('Failed to extract folder ID / フォルダIDを抽出できませんでした');
    }

    const testTitle = 'Test Task / テストタスク - ' + new Date().toLocaleString();
    const testDescription = 'This is a test task. / これはテスト用のタスクです。<br />Includes line break test. / 改行テストも含まれています。';

    const result = createWrikeTask(folderId, testTitle, testDescription);

    if (result) {
      console.log('Test task creation successful: / テストタスク作成成功:', result);
      SpreadsheetApp.getUi().alert('Test task creation successful / テストタスクの作成に成功しました');
    }

  } catch (error) {
    console.error('Test task creation error: / テストタスク作成エラー:', error);
    SpreadsheetApp.getUi().alert(`Test task creation error: / テストタスク作成エラー: ${error.message}`);
  }
}

/**
 * Function for checking configuration / 設定確認用関数
 */
function checkConfiguration() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backlogSheet = ss.getSheetByName('BacklogData');

    let message = 'Configuration check results: / 設定確認結果:\n';
    message += `API URL: ${apiUrl ? 'Set / 設定済み' : 'Not set / 未設定'}\n`;
    message += `Token: ${token ? 'Set / 設定済み' : 'Not set / 未設定'}\n`;
    message += `BacklogData sheet: ${backlogSheet ? 'Exists / 存在' : 'Does not exist / 存在しない'}\n`;

    if (backlogSheet) {
      const permalink = backlogSheet.getRange('B1').getValue();
      const folderId = extractFolderIdFromPermalink(permalink);
      message += `Cell B1 permalink: / B1セルのパーマリンク: ${permalink ? 'Set / 設定済み' : 'Not set / 未設定'}\n`;
      message += `Extracted folder ID: / 抽出されたフォルダID: ${folderId || 'Extraction failed / 抽出失敗'}`;
    }

    console.log(message);
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    console.error('Configuration check error: / 設定確認エラー:', error);
    SpreadsheetApp.getUi().alert(`Configuration check error: / 設定確認エラー: ${error.message}`);
  }
}
