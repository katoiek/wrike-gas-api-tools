/**
 * ユーザーを一括で招待するための機能
 * Wrike API v4の招待エンドポイントを使用
 *
 * BulkUserInviteシートからユーザー情報を取得し、一括で招待を送信します
 */

/**
 * ユーザーを一括招待する関数
 */
function BulkInviteUsers() {
  try {
    // スプレッドシートからユーザー情報を取得
    const userList = getUsersFromSpreadsheet();
    if (userList.length === 0) {
      throw new Error('招待するユーザーが見つかりませんでした。BulkUserInviteシートにデータを入力してください。');
    }

    Logger.log(`${userList.length}人のユーザーを招待します。`);

    // アカウントIDを取得
    const accountId = scriptProperties.getProperty('account_Id');
    if (!accountId) {
      throw new Error('アカウントIDが設定されていません。パラメータシートを確認してください。');
    }

    // 招待を送信
    const results = sendInvitations(userList, accountId);

    // 結果をスプレッドシートに書き込む
    writeResultsToSpreadsheet(results);

    // 成功メッセージを表示
    const successCount = results.filter(r => r.success).length;
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${successCount}/${userList.length}人のユーザーに招待を送信しました。詳細はBulkUserInviteシートの結果列を確認してください。`,
      '招待処理完了'
    );

  } catch (error) {
    // エラーハンドリング
    console.error('ユーザー招待中にエラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラーが発生しました: ' + error.message,
      'エラー',
      10
    );
  }
}

/**
 * スプレッドシートからユーザー情報を取得する関数
 *
 * @return {Array} ユーザー情報の配列
 */
function getUsersFromSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BulkUserInvite');

    if (!sheet) {
      throw new Error("'BulkUserInvite'シートが見つかりません。シートを作成してください。");
    }

    // データ範囲を取得（A2:D100）- ヘッダー行をスキップ
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
    const values = dataRange.getValues();

    // 有効なユーザー情報のみをフィルタリング
    const users = values
      .filter(row => row[0] && row[1] && row[3]) // メールアドレス、名、役割が入力されている行のみ
      .map(row => {
        // 役割の表示名からAPIに送信する値に変換
        let roleValue = null;
        const roleDisplay = row[3].toString().trim();

        if (roleDisplay.startsWith('User')) {
          roleValue = 'User';
        } else if (roleDisplay.startsWith('Collaborator')) {
          roleValue = 'Collaborator';
        } else if (roleDisplay.startsWith('External')) {
          roleValue = 'External';
        }

        // 有効な役割でない場合はnullを返す（後でフィルタリングされる）
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
      .filter(user => user !== null); // nullの項目を除外

    return users;

  } catch (error) {
    Logger.log(`スプレッドシートからのユーザー情報取得中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * 招待を送信する関数
 *
 * @param {Array} users - ユーザー情報の配列
 * @param {string} accountId - アカウントID
 * @return {Array} 招待結果の配列
 */
function sendInvitations(users, accountId) {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  if (!apiUrl || !token) {
    throw new Error('API URLまたはトークンが設定されていません。Wrike認証を実行してください。');
  }

  // 招待エンドポイント
  const invitationsEndpoint = `${apiUrl}/invitations`;

  // 招待結果を格納する配列
  const results = [];

  // 各ユーザーに対して招待を送信
  for (const user of users) {
    try {
      // リクエストオプション
      const requestOptions = getRequestOptions(token, 'POST', {
        email: user.email,
        firstName: user.firstName,
        lastName: user.lastName,
        role: user.role,
        accountId: accountId
      });

      Logger.log(`ユーザー招待リクエスト: ${user.email}, ${user.firstName} ${user.lastName}, 役割: ${user.roleDisplay} (API値: ${user.role})`);

      // APIリクエストを実行
      const response = UrlFetchApp.fetch(invitationsEndpoint, requestOptions);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        // 成功
        const responseData = JSON.parse(responseText);
        Logger.log(`ユーザー招待成功: ${user.email}, 招待ID: ${responseData.data[0].id}`);

        results.push({
          email: user.email,
          success: true,
          message: '招待送信成功',
          invitationId: responseData.data[0].id,
          status: responseData.data[0].status
        });
      } else {
        // エラー
        Logger.log(`ユーザー招待エラー: ${user.email}, コード: ${responseCode}, レスポンス: ${responseText}`);

        let errorMessage = '招待送信失敗';
        try {
          const errorData = JSON.parse(responseText);
          errorMessage += `: ${errorData.error || errorData.errorDescription || '不明なエラー'}`;
        } catch (e) {
          errorMessage += `: ${responseText || 'レスポンスの解析に失敗しました'}`;
        }

        results.push({
          email: user.email,
          success: false,
          message: errorMessage
        });
      }

      // APIリクエストの間隔を空ける（レート制限対策）
      Utilities.sleep(200);

    } catch (error) {
      // 例外発生
      Logger.log(`ユーザー招待例外: ${user.email}, エラー: ${error.message}`);

      results.push({
        email: user.email,
        success: false,
        message: `招待処理中にエラーが発生: ${error.message}`
      });
    }
  }

  return results;
}

/**
 * 招待結果をスプレッドシートに書き込む関数
 *
 * @param {Array} results - 招待結果の配列
 */
function writeResultsToSpreadsheet(results) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BulkUserInvite');

    if (!sheet) {
      throw new Error("'BulkUserInvite'シートが見つかりません。");
    }

    // データ範囲を取得
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues();

    // 結果を書き込む
    for (const result of results) {
      // メールアドレスの行を検索
      for (let i = 0; i < emails.length; i++) {
        if (emails[i][0].toString().trim() === result.email) {
          // 結果を書き込む（E列に結果、F列に詳細メッセージ）
          sheet.getRange(i + 2, 5).setValue(result.success ? '成功' : '失敗');
          sheet.getRange(i + 2, 6).setValue(result.message);

          // 招待IDと状態を書き込む（成功した場合のみ）
          if (result.success) {
            sheet.getRange(i + 2, 7).setValue(result.invitationId);
            sheet.getRange(i + 2, 8).setValue(result.status);
          }

          break;
        }
      }
    }

    // 列の幅を自動調整
    sheet.autoResizeColumns(5, 4);

  } catch (error) {
    Logger.log(`結果の書き込み中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Wrike APIリクエストのための共通オプションを取得する関数
 *
 * @param {string} token - APIトークン
 * @param {string} method - HTTPメソッド
 * @param {Object} [payload] - リクエストボディ（オプション）
 * @return {Object} リクエストオプション
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
 * BulkUserInviteシートを初期化する関数
 */
function initBulkUserInviteSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('BulkUserInvite');

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet('BulkUserInvite');
    }

    // シートをクリア
    sheet.clear();

    // ヘッダー行を設定
    const headers = [
      'メールアドレス*', '名*', '姓', '役割*', '結果', '詳細', '招待ID', '状態'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // 役割のデータ検証（ドロップダウンリスト）を設定
    const roleValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['User(正規ユーザー)', 'Collaborator(コラボレーター)', 'External(外部ユーザー)'], true)
      .setAllowInvalid(false)
      .build();

    // D列（役割）にデータ検証を適用
    sheet.getRange(2, 4, 100).setDataValidation(roleValidation);

    // デフォルト値を設定
    sheet.getRange(2, 4, 100).setValue('Collaborator(コラボレーター)');

    // 列の幅を調整
    sheet.setColumnWidth(1, 250); // メールアドレス
    sheet.setColumnWidth(2, 150); // 名
    sheet.setColumnWidth(3, 150); // 姓
    sheet.setColumnWidth(4, 120); // 役割
    sheet.setColumnWidth(5, 80);  // 結果
    sheet.setColumnWidth(6, 250); // 詳細
    sheet.setColumnWidth(7, 200); // 招待ID
    sheet.setColumnWidth(8, 120); // 状態

    // 説明を追加
    sheet.getRange(1, 10).setValue('使用方法:');
    sheet.getRange(2, 10).setValue('1. メールアドレスと名前を入力してください（必須）');
    sheet.getRange(3, 10).setValue('2. 姓は任意です');
    sheet.getRange(4, 10).setValue('3. 役割はドロップダウンから選択してください（必須）');
    sheet.getRange(5, 10).setValue('4. 役割が空の行はスキップされます');
    sheet.getRange(6, 10).setValue('5. 入力後、「ユーザー一括招待」メニューを実行してください');

    // 成功メッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'BulkUserInviteシートが初期化されました。ユーザー情報を入力してください。',
      'シート初期化完了'
    );

  } catch (error) {
    console.error('シート初期化中にエラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラーが発生しました: ' + error.message,
      'エラー',
      10
    );
  }
}
