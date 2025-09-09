/**
 * ユーザーを一括でViewerに変換し、指定したグループに追加する関数
 *
 * この関数は以下の処理を行います：
 * 1. Userlist2シートからメールアドレスを取得してUser ID v4を取得・更新
 * 2. 各ユーザーのロールをViewerに変更
 * 3. 更新の成功・失敗に関わらず、すべてのユーザーを指定したグループに追加
 */
function UpdateUsertoViewer() {
  // 定数の定義
  const GROUP_ID = 'KX7XTCJ7'; // 追加先のグループID（必要に応じて変更してください）
  const USER_ROLE = 'Viewer'; // 設定するユーザーロール
  const VERSION = '2025-01-27 10:00'; // スクリプトのバージョン

  // スクリプトの実行開始をログに記録
  Logger.log(`Running UpdateUsertoViewer script - Version: ${VERSION}`);
  Logger.log(`Target Group ID: ${GROUP_ID}`);

  try {
    // まず、メールアドレスからUser ID v4を取得・更新
    Logger.log('Step 1: Updating User ID v4 from email addresses...');
    const updateResult = updateUserIdFromEmails();

    if (!updateResult.success) {
      Logger.log(`Failed to update User IDs: ${updateResult.message}`);
      return;
    }

    Logger.log(`Successfully updated ${updateResult.updatedCount} User IDs.`);

    // スプレッドシートからユーザーIDを取得
    const userIds = getUserIdsFromUserlist2();
    if (userIds.length === 0) {
      Logger.log('No valid user IDs found in the Userlist2 spreadsheet.');
      return;
    }

    Logger.log(`Found ${userIds.length} user IDs in spreadsheet: ${userIds.join(', ')}`);

    // ユーザーをViewerに変更
    const updatedUsers = updateUsersToViewer(userIds, USER_ROLE);
    Logger.log(`Successfully updated ${updatedUsers.length} users to Viewer user type.`);

    // 更新の成功・失敗に関わらず、すべてのユーザーをグループに追加
    Logger.log(`Attempting to add all ${userIds.length} users to group ${GROUP_ID}, regardless of update success.`);
    const addedToGroup = addUsersToGroup(userIds, GROUP_ID);
    Logger.log(`Successfully added ${addedToGroup ? userIds.length : 0} users to group ${GROUP_ID}.`);

  } catch (error) {
    Logger.log(`Main process error: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Userlist2シートのメールアドレスからUser ID v4を取得して更新する関数
 *
 * @return {Object} 更新結果 {success: boolean, message: string, updatedCount: number}
 */
function updateUserIdFromEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userlist2Sheet = ss.getSheetByName('Userlist2');

    if (!userlist2Sheet) {
      return {
        success: false,
        message: "'Userlist2' sheet not found in the spreadsheet.",
        updatedCount: 0
      };
    }

    // B列のメールアドレスを取得（1行目をスキップ）
    const emailRange = userlist2Sheet.getRange('B2:B' + userlist2Sheet.getLastRow());
    const emailValues = emailRange.getValues();

    // 空でないメールアドレスのみを抽出
    const emails = emailValues
      .map(row => row[0])
      .filter(email => email !== '' && typeof email === 'string' && email.includes('@'));

    if (emails.length === 0) {
      return {
        success: false,
        message: "No valid email addresses found in column B.",
        updatedCount: 0
      };
    }

    Logger.log(`Found ${emails.length} email addresses: ${emails.join(', ')}`);

    // バッチサイズを設定（APIの制限に応じて調整）
    const BATCH_SIZE = 50;
    const userIdMapping = {};
    let totalUpdated = 0;

    // メールアドレスをバッチごとに処理
    for (let i = 0; i < emails.length; i += BATCH_SIZE) {
      const emailBatch = emails.slice(i, i + BATCH_SIZE);

      try {
        const batchResult = getUserIdsByEmails(emailBatch);

        // 結果をマッピングに追加
        Object.assign(userIdMapping, batchResult);
        totalUpdated += Object.keys(batchResult).length;

        Logger.log(`Batch ${Math.floor(i / BATCH_SIZE) + 1}: Retrieved ${Object.keys(batchResult).length} User IDs`);

        // APIリクエストの間隔を空ける（レート制限対策）
        if (i + BATCH_SIZE < emails.length) {
          Utilities.sleep(500);
        }

      } catch (error) {
        Logger.log(`Error processing batch starting at index ${i}: ${error.message}`);
      }
    }

    // 結果をスプレッドシートに書き込み
    if (Object.keys(userIdMapping).length > 0) {
      writeUserIdsToSpreadsheet(userlist2Sheet, emailValues, userIdMapping);

      return {
        success: true,
        message: `Successfully retrieved User IDs for ${totalUpdated} email addresses.`,
        updatedCount: totalUpdated
      };
    } else {
      return {
        success: false,
        message: "No User IDs could be retrieved from the provided email addresses.",
        updatedCount: 0
      };
    }

  } catch (error) {
    Logger.log(`Error in updateUserIdFromEmails: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`,
      updatedCount: 0
    };
  }
}

/**
 * メールアドレスの配列からUser IDを取得する関数
 *
 * @param {Array} emails - メールアドレスの配列
 * @return {Object} メールアドレスとUser IDのマッピング
 */
function getUserIdsByEmails(emails) {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  if (!apiUrl || !token) {
    throw new Error('API URL or token not found in script properties. Please authenticate first.');
  }

  // emailsパラメータを配列形式で構築
  const emailsParam = '[' + emails.map(email => encodeURIComponent(email)).join(',') + ']';
  const apiEndpoint = `${apiUrl}/contacts?emails=${emailsParam}`;

  Logger.log(`API Endpoint: ${apiEndpoint}`);

  const requestOptions = {
    method: 'GET',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(apiEndpoint, requestOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`API Response Code: ${responseCode}`);
    Logger.log(`API Response: ${responseText}`);

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseText);
      const contacts = jsonResponse.data || [];

      const userIdMapping = {};

      // レスポンスからメールアドレスとUser IDのマッピングを作成
      contacts.forEach(contact => {
        if (contact.profiles && contact.profiles.length > 0 && contact.id) {
          const profile = contact.profiles[0]; // 最初のプロファイルを使用
          if (profile.email) {
            userIdMapping[profile.email] = contact.id; // contact.idを使用
            Logger.log(`Mapped: ${profile.email} -> ${contact.id}`);
          }
        }
      });

      return userIdMapping;

    } else {
      throw new Error(`API request failed with code ${responseCode}: ${responseText}`);
    }

  } catch (error) {
    Logger.log(`Exception in getUserIdsByEmails: ${error.message}`);
    throw error;
  }
}

/**
 * 取得したUser IDをスプレッドシートに書き込む関数
 *
 * @param {Sheet} sheet - 対象のシート
 * @param {Array} emailValues - 元のメールアドレスの配列
 * @param {Object} userIdMapping - メールアドレスとUser IDのマッピング
 */
function writeUserIdsToSpreadsheet(sheet, emailValues, userIdMapping) {
  try {
    // A列に書き込むためのデータを準備
    const userIdResults = [];

    for (let i = 0; i < emailValues.length; i++) {
      const email = emailValues[i][0];
      if (email !== '' && userIdMapping[email]) {
        userIdResults.push([userIdMapping[email]]);
      } else {
        userIdResults.push(['']); // マッピングが見つからない場合は空文字
      }
    }

    // A列（2行目から）に結果を書き込み
    if (userIdResults.length > 0) {
      sheet.getRange(2, 1, userIdResults.length, 1).setValues(userIdResults);
      Logger.log(`Successfully wrote ${userIdResults.length} User IDs to column A.`);
    }

  } catch (error) {
    Logger.log(`Error writing User IDs to spreadsheet: ${error.message}`);
    throw error;
  }
}

/**
 * Userlist2シートからユーザーIDを取得する関数
 *
 * @return {Array} 有効なユーザーIDの配列
 */
function getUserIdsFromUserlist2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userlist2Sheet = ss.getSheetByName('Userlist2');

    if (!userlist2Sheet) {
      throw new Error("'Userlist2' sheet not found in the spreadsheet.");
    }

    // A列のUser ID v4を配列として取得（1行目をスキップ）
    const userIdRange = userlist2Sheet.getRange('A2:A' + userlist2Sheet.getLastRow());
    const userIdValues = userIdRange.getValues();

    // 空でないユーザーIDのみをフラット化して返す
    return userIdValues
      .map(row => row[0])
      .filter(id => id !== '');

  } catch (error) {
    Logger.log(`Error getting user IDs from Userlist2 spreadsheet: ${error.message}`);
    return [];
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
 * ユーザーのユーザータイプをViewerに変更する関数
 * 新しいuserTypeIdパラメータを使用（profileパラメータは非推奨）
 *
 * @param {Array} userIds - ユーザーIDの配列
 * @param {string} role - 設定するロール（後方互換性のため保持、実際は使用されない）
 * @return {Array} 更新に成功したユーザーIDの配列
 */
function updateUsersToViewer(userIds, role = 'Viewer') {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  // ViewerのユーザータイプID
  const VIEWER_USER_TYPE_ID = 'IEAC2GFMNH77777Y';

  // リクエストのペイロード（新しいuserTypeIdパラメータを使用）
  const payload = {
    userTypeId: VIEWER_USER_TYPE_ID
  };

  // 更新結果を追跡する配列
  const successfulUpdates = [];
  const permissionErrors = [];
  const otherErrors = [];

  // 各ユーザーを順番に処理
  for (const userId of userIds) {
    try {
      const apiEndpoint = `${apiUrl}/users/${userId}`;
      Logger.log(`Attempting to update user with endpoint: ${apiEndpoint}`);

      const requestOptions = getRequestOptions(token, 'PUT', payload);
      const response = UrlFetchApp.fetch(apiEndpoint, requestOptions);

      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        successfulUpdates.push(userId);
        Logger.log(`Successfully updated user ${userId} to Viewer user type (${VIEWER_USER_TYPE_ID}).`);
      } else {
        // エラーの種類を分類
        if (responseCode === 403 && responseText.includes('not_allowed')) {
          permissionErrors.push(userId);
          Logger.log(`Permission error updating user ${userId}: Code ${responseCode}, Response: ${responseText}`);
        } else {
          otherErrors.push(userId);
          Logger.log(`Error updating user ${userId}: Code ${responseCode}, Response: ${responseText}`);
        }
      }

      // APIリクエストの間隔を空ける（レート制限対策）
      Utilities.sleep(100);

    } catch (error) {
      Logger.log(`Exception occurred while updating user ${userId}: ${error.message}`);
      otherErrors.push(userId);
    }
  }

  // エラーの統計情報をログに出力
  if (permissionErrors.length > 0) {
    Logger.log(`${permissionErrors.length} users could not be updated due to permission/license limitations.`);
  }

  if (otherErrors.length > 0) {
    Logger.log(`${otherErrors.length} users could not be updated due to other errors.`);
  }

  return successfulUpdates;
}

/**
 * ユーザーをグループに追加する関数
 *
 * @param {Array} userIds - ユーザーIDの配列
 * @param {string} groupId - グループID
 * @return {boolean} 追加に成功したかどうか
 */
function addUsersToGroup(userIds, groupId) {
  if (userIds.length === 0) {
    Logger.log('No users to add to group.');
    return false;
  }

  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    Logger.log(`API URL from properties: ${apiUrl}`);

    // URLパラメータ方式でユーザーをグループに追加
    const userIdsParam = userIds.join(',');
    const groupApiEndpoint = `${apiUrl}/groups/${groupId}?addMembers=[${userIdsParam}]`;

    Logger.log(`Attempting to add users to group with endpoint: ${groupApiEndpoint}`);

    const requestOptions = getRequestOptions(token, 'PUT');
    const groupResponse = UrlFetchApp.fetch(groupApiEndpoint, requestOptions);

    const responseCode = groupResponse.getResponseCode();
    const responseText = groupResponse.getContentText();

    Logger.log(`Group API response code: ${responseCode}`);
    Logger.log(`Group API response: ${responseText}`);

    if (responseCode === 200) {
      Logger.log(`Successfully added users to group ${groupId} using primary method.`);
      return true;
    } else {
      Logger.log(`Error adding users to group: Code ${responseCode}, Response: ${responseText}`);

      // 別の方法を試す
      if (responseCode !== 200) {
        Logger.log(`Trying alternative method for adding users to group ${groupId}...`);
        return tryAlternativeGroupAddMethod(userIds, groupId);
      }

      return false;
    }

  } catch (error) {
    Logger.log(`Exception occurred while adding users to group: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return false;
  }
}

/**
 * 代替方法でユーザーをグループに追加する関数
 *
 * @param {Array} userIds - ユーザーIDの配列
 * @param {string} groupId - グループID
 * @return {boolean} 追加に成功したかどうか
 */
function tryAlternativeGroupAddMethod(userIds, groupId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    // 代替エンドポイント（JSONボディ方式）
    const alternativeEndpoint = `${apiUrl}/groups/${groupId}`;
    Logger.log(`Trying alternative method with endpoint: ${alternativeEndpoint}`);

    // 代替ペイロード
    const payload = {
      addMembers: userIds
    };

    const requestOptions = getRequestOptions(token, 'PUT', payload);
    Logger.log(`Alternative payload: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(alternativeEndpoint, requestOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Alternative method response code: ${responseCode}`);
    Logger.log(`Alternative method response: ${responseText}`);

    if (responseCode === 200) {
      Logger.log(`Successfully added users to group ${groupId} using alternative method.`);
      return true;
    } else {
      Logger.log(`Alternative method failed: Code ${responseCode}, Response: ${responseText}`);
      return false;
    }

  } catch (error) {
    Logger.log(`Exception in alternative method: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return false;
  }
}

/**
 * メールアドレスからUser ID v4を取得する単独実行関数
 * （テスト用・個別実行用）
 */
function getUserIdFromEmailsOnly() {
  Logger.log('Starting getUserIdFromEmailsOnly function...');

  const result = updateUserIdFromEmails();

  if (result.success) {
    Logger.log(`✅ Success: ${result.message}`);
    Logger.log(`Updated ${result.updatedCount} User IDs.`);
  } else {
    Logger.log(`❌ Failed: ${result.message}`);
  }

  return result;
}
