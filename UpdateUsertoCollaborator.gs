/**
 * Function to bulk convert users to collaborators and add them to a specified group / ユーザーを一括でコラボレーターに変換し、指定したグループに追加する関数
 *
 * This function performs the following processes: / この関数は以下の処理を行います：
 * 1. Get user IDs from the spreadsheet / スプレッドシートからユーザーIDを取得
 * 2. Change each user's role to collaborator / 各ユーザーのロールをコラボレーターに変更
 * 3. Add all users to the specified group regardless of success/failure of the update / 更新の成功・失敗に関わらず、すべてのユーザーを指定したグループに追加
 */
function UpdateUsertoCollaborator() {
  // Constant definitions / 定数の定義
  const GROUP_ID = 'KX7XTCJ7'; // Target Group ID / 追加先のグループID
  const USER_ROLE = 'Collaborator'; // User role to set / 設定するユーザーロール
  const VERSION = '2023-05-15 16:00'; // Script version / スクリプトのバージョン

  // Log the start of script execution / スクリプトの実行開始をログに記録
  Logger.log(`Running UpdateUsertoCollaborator script - Version: ${VERSION}`);
  Logger.log(`Target Group ID: ${GROUP_ID}`);

  try {
    // Get user IDs from the spreadsheet / スプレッドシートからユーザーIDを取得
    const userIds = getUserIdsFromSpreadsheet();
    if (userIds.length === 0) {
      Logger.log('No valid user IDs found in the spreadsheet.');
      return;
    }

    Logger.log(`Found ${userIds.length} user IDs in spreadsheet: / スプレッドシートで${userIds.length}個のユーザーIDが見つかりました: ${userIds.join(', ')}`);

    // Change users to collaborators / ユーザーをコラボレーターに変更
    const updatedUsers = updateUsersToCollaborator(userIds, USER_ROLE);
    Logger.log(`Successfully updated ${updatedUsers.length} users to ${USER_ROLE} role. / ${updatedUsers.length}人のユーザーのロールを${USER_ROLE}に正常に更新しました。`);

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
 * Function to retrieve user IDs from the spreadsheet / スプレッドシートからユーザーIDを取得する関数
 *
 * @return {Array} Array of valid user IDs / 有効なユーザーIDの配列
 */
function getUserIdsFromSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userListSheet = ss.getSheetByName('Userlist');

    if (!userListSheet) {
      throw new Error("'Userlist' sheet not found in the spreadsheet. / スプレッドシートに'Userlist'シートが見つかりません。");
    }

    // Get User ID v4 from column A (skip row 1) / A列のUser ID v4を配列として取得（1行目をスキップ）
    const userIdRange = userListSheet.getRange('A2:A' + userListSheet.getLastRow());
    const userIdValues = userIdRange.getValues();

    // Flatten and return only non-empty user IDs / 空でないユーザーIDのみをフラット化して返す
    return userIdValues
      .map(row => row[0])
      .filter(id => id !== '');

  } catch (error) {
    Logger.log(`Error getting user IDs from spreadsheet: / スプレッドシートからのユーザーID取得中にエラーが発生しました: ${error.message}`);
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
 * Function to change user roles to collaborators / ユーザーのロールをコラボレーターに変更する関数
 *
 * @param {Array} userIds - Array of user IDs / ユーザーIDの配列
 * @param {string} role - Role to set (default: 'Collaborator') / 設定するロール（デフォルト: 'Collaborator'）
 * @return {Array} Array of user IDs successfully updated / 更新に成功したユーザーIDの配列
 */
function updateUsersToCollaborator(userIds, role = 'Collaborator') {
  const accountId = scriptProperties.getProperty('account_Id');
  const apiUrl = scriptProperties.getProperty('api_url');
  const token = scriptProperties.getProperty('token');

  // Request payload / リクエストのペイロード
  const payload = {
    profile: {
      accountId: accountId,
      role: role
    }
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
        Logger.log(`Successfully updated user ${userId} to ${role} role. / ユーザー ${userId} のロールを ${role} に正常に更新しました。`);
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
