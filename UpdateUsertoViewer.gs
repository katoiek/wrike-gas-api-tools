/**
 * Function to bulk convert users to Viewers and add them to a specified group / ユーザーを一括でViewerに変換し、指定したグループに追加する関数
 *
 * This function performs the following processes: / この関数は以下の処理を行います：
 * 1. Get email addresses from Userlist2 sheet and retrieve/update User ID v4 / Userlist2シートからメールアドレスを取得してUser ID v4を取得・更新
 * 2. Change each user's role to Viewer / 各ユーザーのロールをViewerに変更
 * 3. Add all users to the specified group regardless of success/failure of the update / 更新の成功・失敗に関わらず、すべてのユーザーを指定したグループに追加
 */
function UpdateUsertoViewer() {
  // Constant definitions / 定数の定義
  const GROUP_ID = 'KX7XTCJ7'; // Target Group ID (change as necessary) / 追加先のグループID（必要に応じて変更してください）
  const USER_ROLE = 'Viewer'; // User role to set / 設定するユーザーロール
  const VERSION = '2025-01-27 10:00'; // Script version / スクリプトのバージョン

  // Log the start of script execution / スクリプトの実行開始をログに記録
  Logger.log(`Running UpdateUsertoViewer script - Version: ${VERSION}`);
  Logger.log(`Target Group ID: ${GROUP_ID}`);

  try {
    // Step 1: Update User ID v4 from email addresses / まず、メールアドレスからUser ID v4を取得・更新
    Logger.log('Step 1: Updating User ID v4 from email addresses...');
    const updateResult = updateUserIdFromEmails();

    if (!updateResult.success) {
      Logger.log(`Failed to update User IDs: ${updateResult.message}`);
      return;
    }

    Logger.log(`Successfully updated ${updateResult.updatedCount} User IDs.`);

    // Get user IDs from the spreadsheet / スプレッドシートからユーザーIDを取得
    const userIds = getUserIdsFromUserlist2();
    if (userIds.length === 0) {
      Logger.log('No valid user IDs found in the Userlist2 spreadsheet.');
      return;
    }

    Logger.log(`Found ${userIds.length} user IDs in spreadsheet: / スプレッドシートで${userIds.length}個のユーザーIDが見つかりました: ${userIds.join(', ')}`);

    // Change users to Viewer role / ユーザーをViewerに変更
    const updatedUsers = updateUsersToViewer(userIds, USER_ROLE);
    Logger.log(`Successfully updated ${updatedUsers.length} users to Viewer user type. / ${updatedUsers.length}人のユーザーのユーザータイプをViewerに正常に更新しました。`);

    // Add all users to the group regardless of success/failure of the update / 更新の成功・失敗に関わらず、すべてのユーザーをグループに追加
    Logger.log(`Attempting to add all ${userIds.length} users to group ${GROUP_ID}, regardless of update success. / 更新の成功に関わらず、全${userIds.length}名のユーザーをグループ${GROUP_ID}に追加を試みます。`);
    const addedToGroup = addUsersToGroup(userIds, GROUP_ID);
    Logger.log(`Successfully added ${addedToGroup ? userIds.length : 0} users to group ${GROUP_ID}. / ユーザーをグループ${GROUP_ID}に正常に追加しました：${addedToGroup ? userIds.length : 0}名。`);

  } catch (error) {
    Logger.log(`Main process error: / メイン処理エラー: ${error.message}`);
    Logger.log(`Stack trace: / スタックトレース: ${error.stack}`);
  }
}

/**
 * Function to retrieve/update User ID v4 from email addresses in Userlist2 sheet / Userlist2シートのメールアドレスからUser ID v4を取得して更新する関数
 *
 * @return {Object} Update result {success: boolean, message: string, updatedCount: number} / 更新結果
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

    // Get email addresses from column B (skip row 1) / B列のメールアドレスを取得（1行目をスキップ）
    const emailRange = userlist2Sheet.getRange('B2:B' + userlist2Sheet.getLastRow());
    const emailValues = emailRange.getValues();

    // Extract only non-empty and valid email addresses / 空でないメールアドレスのみを抽出
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

    Logger.log(`Found ${emails.length} email addresses: / ${emails.length}個のメールアドレスが見つかりました: ${emails.join(', ')}`);

    // Set batch size (adjust based on API limits) / バッチサイズを設定（APIの制限に応じて調整）
    const BATCH_SIZE = 50;
    const userIdMapping = {};
    let totalUpdated = 0;

    // Process email addresses in batches / メールアドレスをバッチごとに処理
    for (let i = 0; i < emails.length; i += BATCH_SIZE) {
      const emailBatch = emails.slice(i, i + BATCH_SIZE);

      try {
        const batchResult = getUserIdsByEmails(emailBatch);

        // Add results to mapping / 結果をマッピングに追加
        Object.assign(userIdMapping, batchResult);
        totalUpdated += Object.keys(batchResult).length;

        Logger.log(`Batch ${Math.floor(i / BATCH_SIZE) + 1}: Retrieved ${Object.keys(batchResult).length} User IDs`);

        // Space out API requests (rate limit countermeasure) / APIリクエストの間隔を空ける（レート制限対策）
        if (i + BATCH_SIZE < emails.length) {
          Utilities.sleep(500);
        }

      } catch (error) {
        Logger.log(`Error processing batch starting at index ${i}: ${error.message}`);
      }
    }

    // Write results to spreadsheet / 結果をスプレッドシートに書き込み
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
 * Function to retrieve User IDs from an array of email addresses / メールアドレスの配列からUser IDを取得する関数
 *
 * @param {Array} emails - Array of email addresses / メールアドレスの配列
 * @return {Object} Mapping between email addresses and User IDs / メールアドレスとUser IDのマッピング
 */
function getUserIdsByEmails(emails) {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  if (!apiUrl || !token) {
    throw new Error('API URL or token not found in script properties. Please authenticate first.');
  }

  // Build emails parameter in array format / emailsパラメータを配列形式で構築
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

      // Create mapping between email address and User ID from response / レスポンスからメールアドレスとUser IDのマッピングを作成
      contacts.forEach(contact => {
        if (contact.profiles && contact.profiles.length > 0 && contact.id) {
          const profile = contact.profiles[0]; // Use first profile / 最初のプロファイルを使用
          if (profile.email) {
            userIdMapping[profile.email] = contact.id; // Use contact.id / contact.idを使用
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
 * Function to write retrieved User IDs to the spreadsheet / 取得したUser IDをスプレッドシートに書き込む関数
 *
 * @param {Sheet} sheet - Target sheet / 対象のシート
 * @param {Array} emailValues - Array of original email addresses / 元のメールアドレスの配列
 * @param {Object} userIdMapping - Mapping between email addresses and User IDs / メールアドレスとUser IDのマッピング
 */
function writeUserIdsToSpreadsheet(sheet, emailValues, userIdMapping) {
  try {
    // Prepare data to be written to column A / A列に書き込むためのデータを準備
    const userIdResults = [];

    for (let i = 0; i < emailValues.length; i++) {
      const email = emailValues[i][0];
      if (email !== '' && userIdMapping[email]) {
        userIdResults.push([userIdMapping[email]]);
      } else {
        userIdResults.push(['']); // Empty string if mapping not found / マッピングが見つからない場合は空文字
      }
    }

    // Write results to column A (starting from row 2) / A列（2行目から）に結果を書き込み
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
 * Function to retrieve user IDs from Userlist2 sheet / Userlist2シートからユーザーIDを取得する関数
 *
 * @return {Array} Array of valid user IDs / 有効なユーザーIDの配列
 */
function getUserIdsFromUserlist2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userlist2Sheet = ss.getSheetByName('Userlist2');

    if (!userlist2Sheet) {
      throw new Error("'Userlist2' sheet not found in the spreadsheet.");
    }

    // Get User ID v4 from column A (skip row 1) / A列のUser ID v4を配列として取得（1行目をスキップ）
    const userIdRange = userlist2Sheet.getRange('A2:A' + userlist2Sheet.getLastRow());
    const userIdValues = userIdRange.getValues();

    // Flatten and return only non-empty user IDs / 空でないユーザーIDのみをフラット化して返す
    return userIdValues
      .map(row => row[0])
      .filter(id => id !== '');

  } catch (error) {
    Logger.log(`Error getting user IDs from Userlist2 spreadsheet: ${error.message}`);
    return [];
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
 * Function to change user type to Viewer / ユーザーのユーザータイプをViewerに変更する関数
 * Uses new userTypeId parameter (profile parameter is deprecated) / 新しいuserTypeIdパラメータを使用（profileパラメータは非推奨）
 *
 * @param {Array} userIds - Array of user IDs / ユーザーIDの配列
 * @param {string} role - Role to set (retained for backward compatibility, not actually used) / 設定するロール（後方互換性のため保持、実際は使用されない）
 * @return {Array} Array of user IDs successfully updated / 更新に成功したユーザーIDの配列
 */
function updateUsersToViewer(userIds, role = 'Viewer') {
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  // Viewer user type ID / ViewerのユーザータイプID
  const VIEWER_USER_TYPE_ID = 'IEAC2GFMNH77777Y';

  // Request payload (uses new userTypeId parameter) / リクエストのペイロード（新しいuserTypeIdパラメータを使用）
  const payload = {
    userTypeId: VIEWER_USER_TYPE_ID
  };

  // Arrays to track update results / 更新結果を追跡する配列
  const successfulUpdates = [];
  const permissionErrors = [];
  const otherErrors = [];

  // Process each user sequentially / 各ユーザーを順番に処理
  for (const userId of userIds) {
    try {
      const apiEndpoint = `${apiUrl}/users/${userId}`;
      Logger.log(`Attempting to update user with endpoint: / ユーザー更新を試行中、エンドポイント: ${apiEndpoint}`);

      const requestOptions = getRequestOptions(token, 'PUT', payload);
      const response = UrlFetchApp.fetch(apiEndpoint, requestOptions);

      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        successfulUpdates.push(userId);
        Logger.log(`Successfully updated user ${userId} to Viewer user type (${VIEWER_USER_TYPE_ID}). / ユーザー ${userId} を正常にユーザータイプViewer (${VIEWER_USER_TYPE_ID}) に更新しました。`);
      } else {
        // Categorize error type / エラーの種類を分類
        if (responseCode === 403 && responseText.includes('not_allowed')) {
          permissionErrors.push(userId);
          Logger.log(`Permission error updating user ${userId}: / ユーザー ${userId} の更新中に権限エラーが発生しました: Code ${responseCode}, Response: ${responseText}`);
        } else {
          otherErrors.push(userId);
          Logger.log(`Error updating user ${userId}: / ユーザー ${userId} の更新中にエラーが発生しました: Code ${responseCode}, Response: ${responseText}`);
        }
      }

      // Space out API requests (rate limit countermeasure) / APIリクエストの間隔を空ける（レート制限対策）
      Utilities.sleep(100);

    } catch (error) {
      Logger.log(`Exception occurred while updating user ${userId}: / ユーザー ${userId} の更新中に例外が発生しました: ${error.message}`);
      otherErrors.push(userId);
    }
  }

  // Log statistical information about errors / エラーの統計情報をログに出力
  if (permissionErrors.length > 0) {
    Logger.log(`${permissionErrors.length} users could not be updated due to permission/license limitations. / 権限またはライセンスの制限により、${permissionErrors.length}人のユーザーを更新できませんでした。`);
  }

  if (otherErrors.length > 0) {
    Logger.log(`${otherErrors.length} users could not be updated due to other errors. / その他のエラーにより、${otherErrors.length}人のユーザーを更新できませんでした。`);
  }

  return successfulUpdates;
}

/**
 * Function to add users to a group / ユーザーをグループに追加する関数
 *
 * @param {Array} userIds - Array of user IDs / ユーザーIDの配列
 * @param {string} groupId - Group ID / グループID
 * @return {boolean} Whether the addition was successful / 追加に成功したかどうか
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

    // Add users to group using URL parameter method / URLパラメータ方式でユーザーをグループに追加
    const userIdsParam = userIds.join(',');
    const groupApiEndpoint = `${apiUrl}/groups/${groupId}?addMembers=[${userIdsParam}]`;

    Logger.log(`Attempting to add users to group with endpoint: / ユーザーのグループ追加を試行中、エンドポイント: ${groupApiEndpoint}`);

    const requestOptions = getRequestOptions(token, 'PUT');
    const groupResponse = UrlFetchApp.fetch(groupApiEndpoint, requestOptions);

    const responseCode = groupResponse.getResponseCode();
    const responseText = groupResponse.getContentText();

    Logger.log(`Group API response code: / グループAPIレスポンスコード: ${responseCode}`);
    Logger.log(`Group API response: / グループAPIレスポンス: ${responseText}`);

    if (responseCode === 200) {
      Logger.log(`Successfully added users to group ${groupId} using primary method. / プライマリメソッドを使用して、グループ ${groupId} にユーザーを正常に追加しました。`);
      return true;
    } else {
      Logger.log(`Error adding users to group: / グループへのユーザー追加中にエラーが発生しました: Code ${responseCode}, Response: ${responseText}`);

      // Try alternative method / 別の方法を試す
      if (responseCode !== 200) {
        Logger.log(`Trying alternative method for adding users to group ${groupId}... / グループ ${groupId} へのユーザー追加の代替メソッドを試行中...`);
        return tryAlternativeGroupAddMethod(userIds, groupId);
      }

      return false;
    }

  } catch (error) {
    Logger.log(`Exception occurred while adding users to group: / グループへのユーザー追加中に例外が発生しました: ${error.message}`);
    Logger.log(`Stack trace: / スタックトレース: ${error.stack}`);
    return false;
  }
}

/**
 * Function to add users to a group using an alternative method / 代替方法でユーザーをグループに追加する関数
 *
 * @param {Array} userIds - Array of user IDs / ユーザーIDの配列
 * @param {string} groupId - Group ID / グループID
 * @return {boolean} Whether the addition was successful / 追加に成功したかどうか
 */
function tryAlternativeGroupAddMethod(userIds, groupId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    // Alternative endpoint (JSON body method) / 代替エンドポイント（JSONボディ方式）
    const alternativeEndpoint = `${apiUrl}/groups/${groupId}`;
    Logger.log(`Trying alternative method with endpoint: / 代替メソッドを試行中、エンドポイント: ${alternativeEndpoint}`);

    // Alternative payload / 代替ペイロード
    const payload = {
      addMembers: userIds
    };

    const requestOptions = getRequestOptions(token, 'PUT', payload);
    Logger.log(`Alternative payload: / 代替ペイロード: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(alternativeEndpoint, requestOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Alternative method response code: / 代替メソッドのリスポンスコード: ${responseCode}`);
    Logger.log(`Alternative method response: / 代替メソッドのリスポンス: ${responseText}`);

    if (responseCode === 200) {
      Logger.log(`Successfully added users to group ${groupId} using alternative method. / 代替メソッドを使用して、グループ ${groupId} にユーザーを正常に追加しました。`);
      return true;
    } else {
      Logger.log(`Alternative method failed: / 代替メソッドが失敗しました: Code ${responseCode}, Response: ${responseText}`);
      return false;
    }

  } catch (error) {
    Logger.log(`Exception in alternative method: / 代替メソッドで例外が発生しました: ${error.message}`);
    Logger.log(`Stack trace: / スタックトレース: ${error.stack}`);
    return false;
  }
}

/**
 * Stand-alone function to retrieve User ID v4 from email addresses (for testing/individual execution) / メールアドレスからUser ID v4を取得する単独実行関数（テスト用・個別実行用）
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
