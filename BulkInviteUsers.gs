/**
 * Feature for bulk inviting users / ユーザーを一括で招待するための機能
 * Uses Wrike API v4 invitation endpoint / Wrike API v4の招待エンドポイントを使用
 *
 * Retrieves user information from BulkUserInvite sheet and sends invitations in bulk / BulkUserInviteシートからユーザー情報を取得し、一括で招待を送信します
 */

/**
 * Function to bulk invite users / ユーザーを一括招待する関数
 */
function BulkInviteUsers() {
  try {
    // Get user information from spreadsheet / スプレッドシートからユーザー情報を取得
    const userList = getUsersFromSpreadsheet();
    if (userList.length === 0) {
      throw new Error('No users to invite were found. Please enter data into the BulkUserInvite sheet. / 招待するユーザーが見つかりませんでした。BulkUserInviteシートにデータを入力してください。');
    }

    Logger.log(`Inviting ${userList.length} users. / ${userList.length}人のユーザーを招待します。`);

    // Get account ID / アカウントIDを取得
    const accountId = scriptProperties.getProperty('account_Id');
    if (!accountId) {
      throw new Error('Account ID is not set. Please check the parameters sheet. / アカウントIDが設定されていません。パラメータシートを確認してください。');
    }

    // Send invitations / 招待を送信
    const results = sendInvitations(userList, accountId);

    // Write results to spreadsheet / 結果をスプレッドシートに書き込む
    writeResultsToSpreadsheet(results);

    // Display success message / 成功メッセージを表示
    const successCount = results.filter(r => r.success).length;
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Sent invitations to ${successCount}/${userList.length} users. Please check the result column of the BulkUserInvite sheet for details. / ${successCount}/${userList.length}人のユーザーに招待を送信しました。詳細はBulkUserInviteシートの結果列を確認してください。`,
      'Invitation process complete / 招待処理完了'
    );

  } catch (error) {
    // Error handling / エラーハンドリング
    console.error('An error occurred while inviting users: / ユーザー招待中にエラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'An error occurred: / エラーが発生しました: ' + error.message,
      'Error / エラー',
      10
    );
  }
}

/**
 * Function to retrieve user information from spreadsheet / スプレッドシートからユーザー情報を取得する関数
 *
 * @return {Array} Array of user information / ユーザー情報の配列
 */
function getUsersFromSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BulkUserInvite');

    if (!sheet) {
      throw new Error("'BulkUserInvite' sheet was not found. Please create the sheet. / 'BulkUserInvite'シートが見つかりません。シートを作成してください。");
    }

    // Get data range (A2:D100) - skip header row / データ範囲を取得（A2:D100）- ヘッダー行をスキップ
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
    const values = dataRange.getValues();

    // Filter only valid user information / 有効なユーザー情報のみをフィルタリング
    const users = values
      .filter(row => row[0] && row[1] && row[3]) // Only rows with email, first name, and role entered / メールアドレス、名、役割が入力されている行のみ
      .map(row => {
        // Convert from display name of role to value to be sent to API / 役割の表示名からAPIに送信する値に変換
        let roleValue = null;
        const roleDisplay = row[3].toString().trim();

        if (roleDisplay.startsWith('User')) {
          roleValue = 'User';
        } else if (roleDisplay.startsWith('Collaborator')) {
          roleValue = 'Collaborator';
        } else if (roleDisplay.startsWith('External')) {
          roleValue = 'External';
        }

        // Return null if it's not a valid role (will be filtered out later) / 有効な役割でない場合はnullを返す（後でフィルタリングされる）
        if (!roleValue) {
          return null;
        }

        return {
          email: row[0].toString().trim(),
          firstName: row[1].toString().trim(),
          lastName: row[2] ? row[2].toString().trim() : '',
          role: roleValue,
          roleDisplay: roleDisplay
        };
      })
      .filter(user => user !== null); // Exclude null items / nullの項目を除外

    return users;

  } catch (error) {
    Logger.log(`An error occurred while retrieving user information from the spreadsheet: / スプレッドシートからのユーザー情報取得中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Function to send invitations / 招待を送信する関数
 *
 * @param {Array} users - Array of user information / ユーザー情報の配列
 * @param {string} accountId - Account ID / アカウントID
 * @return {Array} Array of invitation results / 招待結果の配列
 */
function sendInvitations(users, accountId) {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  if (!apiUrl || !token) {
    throw new Error('API URL or token is not set. Please perform Wrike authentication. / API URLまたはトークンが設定されていません。Wrike認証を実行してください。');
  }

  // Invitation endpoint / 招待エンドポイント
  const invitationsEndpoint = `${apiUrl}/invitations`;

  // Array to store invitation results / 招待結果を格納する配列
  const results = [];

  // Send invitation to each user / 各ユーザーに対して招待を送信
  for (const user of users) {
    try {
      // Request options / リクエストオプション
      const requestOptions = getRequestOptions(token, 'POST', {
        email: user.email,
        firstName: user.firstName,
        lastName: user.lastName,
        role: user.role,
        accountId: accountId
      });

      Logger.log(`User invitation request: / ユーザー招待リクエスト: ${user.email}, ${user.firstName} ${user.lastName}, Role: / 役割: ${user.roleDisplay} (API Value: / API値: ${user.role})`);

      // Execute API request / APIリクエストを実行
      const response = UrlFetchApp.fetch(invitationsEndpoint, requestOptions);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        // Success / 成功
        const responseData = JSON.parse(responseText);
        Logger.log(`User invitation success: / ユーザー招待成功: ${user.email}, Invitation ID: / 招待ID: ${responseData.data[0].id}`);

        results.push({
          email: user.email,
          success: true,
          message: 'Invitation sent successfully / 招待送信成功',
          invitationId: responseData.data[0].id,
          status: responseData.data[0].status
        });
      } else {
        // Error / エラー
        Logger.log(`User invitation error: / ユーザー招待エラー: ${user.email}, Code: / コード: ${responseCode}, Response: / レスポンス: ${responseText}`);

        let errorMessage = 'Failed to send invitation / 招待送信失敗';
        try {
          const errorData = JSON.parse(responseText);
          errorMessage += `: ${errorData.error || errorData.errorDescription || 'Unknown error / 不明なエラー'}`;
        } catch (e) {
          errorMessage += `: ${responseText || 'Failed to parse response / レスポンスの解析に失敗しました'}`;
        }

        results.push({
          email: user.email,
          success: false,
          message: errorMessage
        });
      }

      // Space out API requests (rate limit countermeasure) / APIリクエストの間隔を空ける（レート制限対策）
      Utilities.sleep(200);

    } catch (error) {
      // Exception occurred / 例外発生
      Logger.log(`User invitation exception: / ユーザー招待例外: ${user.email}, Error: / エラー: ${error.message}`);

      results.push({
        email: user.email,
        success: false,
        message: `Error occurred during invitation process: / 招待処理中にエラーが発生: ${error.message}`
      });
    }
  }

  return results;
}

/**
 * Function to write invitation results to spreadsheet / 招待結果をスプレッドシートに書き込む関数
 *
 * @param {Array} results - Array of invitation results / 招待結果の配列
 */
function writeResultsToSpreadsheet(results) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BulkUserInvite');

    if (!sheet) {
      throw new Error("'BulkUserInvite' sheet was not found. / 'BulkUserInvite'シートが見つかりません。");
    }

    // Get data range / データ範囲を取得
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();

    // Write results / 結果を書き込む
    for (const result of results) {
      // Search for email address row / メールアドレスの行を検索
      for (let i = 0; i < emails.length; i++) {
        if (emails[i][0].toString().trim() === result.email) {
          // Write result (Column E for result, Column F for detailed message) / 結果を書き込む（E列に結果、F列に詳細メッセージ）
          sheet.getRange(i + 2, 5).setValue(result.success ? 'Success / 成功' : 'Failure / 失敗');
          sheet.getRange(i + 2, 6).setValue(result.message);

          // Write invitation ID and status (only if successful) / 招待IDと状態を書き込む（成功した場合のみ）
          if (result.success) {
            sheet.getRange(i + 2, 7).setValue(result.invitationId);
            sheet.getRange(i + 2, 8).setValue(result.status);
          }

          break;
        }
      }
    }

    // Auto-resize columns / 列の幅を自動調整
    sheet.autoResizeColumns(5, 4);

  } catch (error) {
    Logger.log(`An error occurred while writing results: / 結果の書き込み中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Function to retrieve common options for Wrike API requests / Wrike APIリクエストのための共通オプションを取得する関数
 *
 * @param {string} token - API token / APIトークン
 * @param {string} method - HTTP method / HTTPメソッド
 * @param {Object} [payload] - Request body (optional) / リクエストボディ（オプション）
 * @return {Object} Request options / リクエストオプション
 */
function getRequestOptions(token, method = 'GET', payload = null) {
  const options = {
    method: method,
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = typeof payload === 'string' ? payload : JSON.stringify(payload);
  }

  return options;
}

/**
 * Function to initialize BulkUserInvite sheet / BulkUserInviteシートを初期化する関数
 */
function initBulkUserInviteSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('BulkUserInvite');

    // Create sheet if it does not exist / シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet('BulkUserInvite');
    }

    // Clear sheet / シートをクリア
    sheet.clear();

    // Set header row / ヘッダー行を設定
    const headers = [
      'Email / メールアドレス*', 'First Name / 名*', 'Last Name / 姓', 'Role / 役割*', 'Result / 結果', 'Details / 詳細', 'Invitation ID / 招待ID', 'Status / 状態'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Set data validation for role (dropdown list) / 役割のデータ検証（ドロップダウンリスト）を設定
    const roleValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['User(正規ユーザー)', 'Collaborator(コラボレーター)', 'External(外部ユーザー)'], true)
      .setAllowInvalid(false)
      .build();

    // Apply data validation to Column D (Role) / D列（役割）にデータ検証を適用
    sheet.getRange(2, 4, 100).setDataValidation(roleValidation);

    // Set default value / デフォルト値を設定
    sheet.getRange(2, 4, 100).setValue('Collaborator(コラボレーター)');

    // Adjust column widths / 列の幅を調整
    sheet.setColumnWidth(1, 250); // Email / メールアドレス
    sheet.setColumnWidth(2, 150); // First Name / 名
    sheet.setColumnWidth(3, 150); // Last Name / 姓
    sheet.setColumnWidth(4, 150); // Role / 役割
    sheet.setColumnWidth(5, 100); // Result / 結果
    sheet.setColumnWidth(6, 250); // Details / 詳細
    sheet.setColumnWidth(7, 200); // Invitation ID / 招待ID
    sheet.setColumnWidth(8, 120); // Status / 状態

    // Add description / 説明を追加
    sheet.getRange(1, 10).setValue('Usage: / 使用方法:');
    sheet.getRange(2, 10).setValue('1. Enter email address and first name (Required) / 1. メールアドレスと名前を入力してください（必須）');
    sheet.getRange(3, 10).setValue('2. Last name is optional / 2. 姓は任意です');
    sheet.getRange(4, 10).setValue('3. Select role from dropdown (Required) / 3. 役割はドロップダウンから選択してください（必須）');
    sheet.getRange(5, 10).setValue('4. Rows with empty roles will be skipped / 4. 役割が空の行はスキップされます');
    sheet.getRange(6, 10).setValue('5. After entering, execute the "Bulk User Invite" menu / 5. 入力後、「ユーザー一括招待」メニューを実行してください');

    // Display success message / 成功メッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'BulkUserInvite sheet has been initialized. Please enter user information. / BulkUserInviteシートが初期化されました。ユーザー情報を入力してください。',
      'Sheet initialization complete / シート初期化完了'
    );

  } catch (error) {
    console.error('An error occurred during sheet initialization: / シート初期化中にエラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'An error occurred: / エラーが発生しました: ' + error.message,
      'Error / エラー',
      10
    );
  }
}
