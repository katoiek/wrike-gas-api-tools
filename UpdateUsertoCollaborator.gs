/**
 * ユーザーを一括でコラボレーターに変換し、指定したグループに追加する関数
 *
 * この関数は以下の処理を行います：
 * 1. スプレッドシートからユーザーIDを取得
 * 2. 各ユーザーのロールをコラボレーターに変更
 * 3. 更新の成功・失敗に関わらず、すべてのユーザーを指定したグループに追加
 */
function UpdateUsertoCollaborator() {
  // 定数の定義
  const GROUP_ID = 'KX7XTCJ7'; // 追加先のグループID
  const USER_ROLE = 'Collaborator'; // 設定するユーザーロール
  const VERSION = '2023-05-15 16:00'; // スクリプトのバージョン

  // スクリプトの実行開始をログに記録
  Logger.log(`Running UpdateUsertoCollaborator script - Version: ${VERSION}`);
  Logger.log(`Target Group ID: ${GROUP_ID}`);

  try {
    // スプレッドシートからユーザーIDを取得
    const userIds = getUserIdsFromSpreadsheet();
    if (userIds.length === 0) {
      Logger.log('No valid user IDs found in the spreadsheet.');
      return;
    }

    Logger.log(`Found ${userIds.length} user IDs in spreadsheet: ${userIds.join(', ')}`);

    // ユーザーをコラボレーターに変更
    const updatedUsers = updateUsersToCollaborator(userIds, USER_ROLE);
    Logger.log(`Successfully updated ${updatedUsers.length} users to ${USER_ROLE} role.`);

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
 * スプレッドシートからユーザーIDを取得する関数
 *
 * @return {Array} 有効なユーザーIDの配列
 */
function getUserIdsFromSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userListSheet = ss.getSheetByName('Userlist');

    if (!userListSheet) {
      throw new Error("'Userlist' sheet not found in the spreadsheet.");
    }

    // A列のUser ID v4を配列として取得（1行目をスキップ）
    const userIdRange = userListSheet.getRange('A2:A' + userListSheet.getLastRow());
    const userIdValues = userIdRange.getValues();

    // 空でないユーザーIDのみをフラット化して返す
    return userIdValues
      .map(row => row[0])
      .filter(id => id !== '');

  } catch (error) {
    Logger.log(`Error getting user IDs from spreadsheet: ${error.message}`);
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
 * ユーザーのロールをコラボレーターに変更する関数
 *
 * @param {Array} userIds - ユーザーIDの配列
 * @param {string} role - 設定するロール（デフォルト: 'Collaborator'）
 * @return {Array} 更新に成功したユーザーIDの配列
 */
function updateUsersToCollaborator(userIds, role = 'Collaborator') {
  const accountId = scriptProperties.getProperty('account_Id');
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  // リクエストのペイロード
  const payload = {
    profile: {
      accountId: accountId,
      role: role
    }
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
        Logger.log(`Successfully updated user ${userId} to ${role} role.`);
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
