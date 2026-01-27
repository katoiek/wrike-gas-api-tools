/**
 * GAS for copying custom fields between spaces / カスタムフィールドをスペース間でコピーするためのGAS
 * Copies space-scoped custom fields to another space / スペーススコープのカスタムフィールドを別のスペースにコピーします
 */

/**
 * Function to get SpaceID from permalink / パーマリンクからSpaceIDを取得する関数
 *
 * @param {string} permalink - Permalink of Wrike space / Wrikeスペースのパーマリンク
 * @return {string|null} SpaceID or null (in case of error) / SpaceID または null（エラーの場合）
 */
function getSpaceIdFromPermalink(permalink) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API configuration or token not found. Please check authentication. / APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    const apiEndpoint = `${apiUrl}/folders?permalink=${encodeURIComponent(permalink)}`;
    console.log(`API Endpoint: / APIエンドポイント: ${apiEndpoint}`);

    const response = UrlFetchApp.fetch(apiEndpoint, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
      muteHttpExceptions: true
    });

    console.log(`Response Code: / レスポンスコード: ${response.getResponseCode()}`);
    const responseText = response.getContentText();
    console.log(`Response Content: / レスポンス内容: ${responseText}`);

    if (response.getResponseCode() !== 200) {
      // In case of logical folder error, extract ID and search from space list / logical folderエラーの場合は、IDを抽出してスペース一覧から検索
      if (responseText.includes('logical folder')) {
        console.log('Logical folder error occurred. Searching from space list. / Logical folderエラーが発生。スペース一覧から検索します。');

        // Extract ID from permalink / パーマリンクからIDを抽出
        const idMatch = permalink.match(/id=(\d+)/);
        if (!idMatch) {
          throw new Error('Failed to extract ID from permalink. / パーマリンクからIDを抽出できませんでした。');
        }

        const folderId = idMatch[1];
        console.log(`Extracted Folder ID: / 抽出されたフォルダID: ${folderId}`);

        // Get list of spaces / スペース一覧を取得
        const spacesResponse = UrlFetchApp.fetch(`${apiUrl}/folders?project=false`, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (spacesResponse.getResponseCode() === 200) {
          const spacesJson = JSON.parse(spacesResponse.getContentText());
          const allFolders = spacesJson.data;
          const spaces = allFolders.filter(f => f.scope === 'WsSpace');

          console.log(`Searching through ${spaces.length} spaces... / ${spaces.length} 個のスペースを検索中...`);

          // Search for a space with matching ID / IDが一致するスペースを検索
          for (const space of spaces) {
            if (space.id === folderId) {
              console.log(`Matching space found: / 一致するスペースが見つかりました: ${space.title} (ID: ${space.id})`);
              return space.id;
            }
          }
        }
      }

      throw new Error(`API call failed. Response code: / API呼び出しに失敗しました。レスポンスコード: ${response.getResponseCode()}, Error: / エラー: ${responseText}`);
    }

    const myJson = JSON.parse(responseText);
    const data = myJson.data;

    if (!data || data.length === 0) {
      throw new Error('Folder for the specified permalink not found. / 指定されたパーマリンクのフォルダが見つかりません。');
    }

    // Folder information retrieved by permalink / パーマリンクで取得されたフォルダ情報
    const folder = data[0];
    console.log(`Retrieved folder: / 取得されたフォルダ: ${folder.title}, Scope: / スコア: ${folder.scope}, ID: ${folder.id}`);

    // The ID retrieved by permalink query is used as the space ID / パーマリンクで取得されたフォルダのIDがスペースIDとして使用される
    // (Due to Wrike specifications, the ID returned by permalink query is the space ID) / （Wrikeの仕様により、パーマリンククエリで返されるIDがスペースID）
    console.log(`Using ID retrieved by permalink as space ID: / パーマリンクで取得されたIDをスペースIDとして使用: ${folder.id}`);
    return folder.id;

  } catch (error) {
    console.error('SpaceID retrieval error: / SpaceID取得エラー:', error.message);
    SpreadsheetApp.getUi().alert('Error / エラー', `SpaceID retrieval failed: / SpaceID取得に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Function to retrieve custom fields for the specified space / 指定されたスペースのカスタムフィールドを取得する関数
 *
 * @param {string} spaceId - Space ID / スペースID
 * @return {Array|null} Array of custom fields or null (in case of error) / カスタムフィールド配列 または null（エラーの場合）
 */
function getCustomFieldsBySpaceId(spaceId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API configuration or token not found. Please check authentication. / APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    const apiEndpoint = `${apiUrl}/customfields`;

    const response = UrlFetchApp.fetch(apiEndpoint, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`API call failed. Response code: / API呼び出しに失敗しました。レスポンスコード: ${response.getResponseCode()}`);
    }

    const myJson = JSON.parse(response.getContentText());
    const allCustomFields = myJson.data;

    // Filter only custom fields for the specified space ID / 指定されたスペースIDのカスタムフィールドのみをフィルタリング
    const spaceCustomFields = allCustomFields.filter(field => field.spaceId === spaceId);

    // Exclude non-copyable field types (Calculated fields, database links, etc.) / コピー不可能なフィールドタイプを除外（計算フィールド、データベースリンクなど）
    const unsupportedTypes = ['CalculatedNumeric', 'CalculatedDate', 'LinkToDatabase'];
    const copyableFields = spaceCustomFields.filter(field => !unsupportedTypes.includes(field.type));

    const skippedCount = spaceCustomFields.length - copyableFields.length;

    console.log(`Retrieved ${spaceCustomFields.length} custom fields for space / スペース ${spaceId} のカスタムフィールドを ${spaceCustomFields.length} 件取得しました`);
    if (skippedCount > 0) {
      console.log(`${skippedCount} fields are not eligible for copying (Calculated fields, database links, etc.) / ${skippedCount} 件のフィールドはコピー対象外です（計算フィールド、データベースリンクなど）`);
    }
    console.log(`Copyable fields: / コピー可能なフィールド: ${copyableFields.length} fields`);

    return copyableFields;

  } catch (error) {
    console.error('Custom field retrieval error: / カスタムフィールド取得エラー:', error.message);
    SpreadsheetApp.getUi().alert('Error / エラー', `Failed to retrieve custom fields: / カスタムフィールド取得に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Function to create a custom field in a new space / カスタムフィールドを新しいスペースに作成する関数
 *
 * @param {Object} customField - Original custom field object / 元のカスタムフィールドオブジェクト
 * @param {string} targetSpaceId - Destination space ID / コピー先のスペースID
 * @return {Object|null} Created custom field or null (in case of error) / 作成されたカスタムフィールド または null（エラーの場合）
 */
function createCustomField(customField, targetSpaceId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');

    if (!apiUrl || !token) {
      throw new Error('API configuration or token not found. Please check authentication. / APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    // Build payload for custom field creation / カスタムフィールド作成用のペイロードを構築
    const payload = {
      title: customField.title,
      type: customField.type,
      spaceId: targetSpaceId
    };

    // Add description if present / 説明があれば追加
    if (customField.description) {
      payload.description = customField.description;
    }

    // Add settings if present as a settings object / 設定（settings）があれば、settingsオブジェクトとして追加
    if (customField.settings) {
      payload.settings = {};

      // Add each property in settings object to settings / settingsオブジェクトの各プロパティをsettings内に追加
      if (customField.settings.currency !== undefined) {
        payload.settings.currency = customField.settings.currency;
      }
      if (customField.settings.aggregation !== undefined) {
        payload.settings.aggregation = customField.settings.aggregation;
      }
      if (customField.settings.decimalPlaces !== undefined) {
        payload.settings.decimalPlaces = customField.settings.decimalPlaces;
      }
      if (customField.settings.useThousandsSeparator !== undefined) {
        payload.settings.useThousandsSeparator = customField.settings.useThousandsSeparator;
      }
      if (customField.settings.options !== undefined) {
        payload.settings.options = customField.settings.options;
      }
      if (customField.settings.values !== undefined) {
        payload.settings.values = customField.settings.values;
      }
      if (customField.settings.optionColorsEnabled !== undefined) {
        payload.settings.optionColorsEnabled = customField.settings.optionColorsEnabled;
      }
      if (customField.settings.allowOtherValues !== undefined) {
        payload.settings.allowOtherValues = customField.settings.allowOtherValues;
      }
    }

    const apiEndpoint = `${apiUrl}/customfields`;

    const response = UrlFetchApp.fetch(apiEndpoint, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload)
    });

    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      throw new Error(`Failed to create custom field. Response code: / カスタムフィールド作成に失敗しました。レスポンスコード: ${response.getResponseCode()}, Error: / エラー: ${errorText}`);
    }

    const myJson = JSON.parse(response.getContentText());
    const createdField = myJson.data[0];

    console.log(`Created custom field / カスタムフィールド "${customField.title}" を作成しました。ID: ${createdField.id}`);

    // Wait for rate limiting compliance / レート制限対応のための待機
    Utilities.sleep(200);

    return createdField;

  } catch (error) {
    console.error('Custom field creation error: / カスタムフィールド作成エラー:', error.message);
    throw error;
  }
}

/**
 * Main function to copy custom fields between spaces / カスタムフィールドをスペース間でコピーするメイン関数
 */
function copyCustomFieldsBetweenSpaces() {
  try {
    // Get UI inputs / UI入力を取得
    const ui = SpreadsheetApp.getUi();

    const sourcePermalinkResponse = ui.prompt(
      'Custom Field Copy / カスタムフィールドコピー',
      'Please enter the permalink of the source space: / コピー元スペースのパーマリンクを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );

    if (sourcePermalinkResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const sourcePermalink = sourcePermalinkResponse.getResponseText().trim();
    if (!sourcePermalink) {
      ui.alert('Error / エラー', 'Source space permalink has not been entered. / コピー元スペースのパーマリンクが入力されていません。', ui.ButtonSet.OK);
      return;
    }

    const targetPermalinkResponse = ui.prompt(
      'Custom Field Copy / カスタムフィールドコピー',
      'Please enter the permalink of the destination space: / コピー先スペースのパーマリンクを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );

    if (targetPermalinkResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const targetPermalink = targetPermalinkResponse.getResponseText().trim();
    if (!targetPermalink) {
      ui.alert('Error / エラー', 'Destination space permalink has not been entered. / コピー先スペースのパーマリンクが入力されていません。', ui.ButtonSet.OK);
      return;
    }

    // Initialize properties / プロパティを初期化
    if (!scriptProperties.getProperty('api_url') || !scriptProperties.getProperty('token')) {
      registKeys();
    }

    // Processing start message / 処理開始メッセージ
    createModelessDialog('Processing custom field copy... / カスタムフィールドコピー処理を開始します...', 'Processing / 処理中');

    // Retrieve source space ID / コピー元スペースIDを取得
    console.log('Retrieving source space ID... / コピー元スペースIDを取得中...');
    const sourceSpaceId = getSpaceIdFromPermalink(sourcePermalink);
    if (!sourceSpaceId) {
      return;
    }

    // Retrieve destination space ID / コピー先スペースIDを取得
    console.log('Retrieving destination space ID... / コピー先スペースIDを取得中...');
    const targetSpaceId = getSpaceIdFromPermalink(targetPermalink);
    if (!targetSpaceId) {
      return;
    }

    // Error if same space / 同じスペースの場合はエラー
    if (sourceSpaceId === targetSpaceId) {
      ui.alert('Error / エラー', 'Source and destination are the same space. Please specify different spaces. / コピー元とコピー先が同じスペースです。異なるスペースを指定してください。', ui.ButtonSet.OK);
      return;
    }

    // Retrieve custom fields from source space / コピー元スペースのカスタムフィールドを取得
    console.log('Retrieving custom fields from source space... / コピー元スペースのカスタムフィールドを取得中...');
    const sourceCustomFields = getCustomFieldsBySpaceId(sourceSpaceId);
    if (!sourceCustomFields) {
      return;
    }

    if (sourceCustomFields.length === 0) {
      ui.alert('Information / 情報', 'No custom fields found in the source space. / コピー元スペースにカスタムフィールドが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }

    // Confirmation dialog / 確認ダイアログ
    const confirmResponse = ui.alert(
      'Confirmation / 確認',
      `Copy ${sourceCustomFields.length} custom fields? / ${sourceCustomFields.length} 個のカスタムフィールドをコピーしますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      return;
    }

    // Calculate number of skipped fields / スキップされたフィールド数を計算
    const allCustomFieldsResponse = UrlFetchApp.fetch(`${scriptProperties.getProperty('api_url')}/customfields`, {
      headers: { Authorization: `Bearer ${scriptProperties.getProperty('token')}` }
    });
    const allFieldsJson = JSON.parse(allCustomFieldsResponse.getContentText());
    const allSourceFields = allFieldsJson.data.filter(field => field.spaceId === sourceSpaceId);
    const skippedCount = allSourceFields.length - sourceCustomFields.length;

    // Copy custom fields sequentially / カスタムフィールドを順次コピー
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    console.log('Starting custom field copy... / カスタムフィールドのコピーを開始...');

    for (const customField of sourceCustomFields) {
      try {
        console.log(`Copying custom field / カスタムフィールド "${customField.title}" をコピー中...`);
        createCustomField(customField, targetSpaceId);
        successCount++;
      } catch (error) {
        console.error(`Failed to copy custom field / カスタムフィールド "${customField.title}" のコピーに失敗:`, error.message);
        errorCount++;
        errors.push(`${customField.title}: ${error.message}`);
      }
    }

    // Result message / 結果メッセージ
    let resultMessage = `Custom field copy complete / カスタムフィールドコピー完了\n\n`;
    resultMessage += `Success: / 成功: ${successCount} \n`;
    resultMessage += `Failure: / 失敗: ${errorCount} `;
    if (skippedCount > 0) {
      resultMessage += `\nSkipped: / スキップ: ${skippedCount} items (Calculated fields, database links, etc.) / (計算フィールド、データベースリンクなど)`;
    }

    if (errors.length > 0) {
      resultMessage += `\n\nFailure Details: / 失敗詳細:\n${errors.slice(0, 5).join('\n')}`;
      if (errors.length > 5) {
        resultMessage += `\n... and ${errors.length - 5} others / 他 ${errors.length - 5} 件`;
      }
    }

    // Display completion message in dialog / 完了メッセージをダイアログで表示
    const completionHtml = `
      <div style="padding: 20px;">
        <h3>Custom field copy complete / カスタムフィールドコピー完了</h3>
        <pre style="white-space: pre-wrap; font-family: monospace; font-size: 12px;">${resultMessage}</pre>
      </div>
    `;
    createModelessDialog(completionHtml, 'Completion / 完了');
    console.log('Custom field copy process complete / カスタムフィールドコピー処理完了');

  } catch (error) {
    console.error('Main process error: / メイン処理エラー:', error.message);
    SpreadsheetApp.getUi().alert('Error / エラー', `An error occurred during processing: / 処理中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Function to execute custom field copy from menu / メニューからカスタムフィールドコピーを実行するための関数
 */
function showCopyCustomFieldsDialog() {
  copyCustomFieldsBetweenSpaces();
}

/**
 * For debugging: Function to test SpaceID retrieval for specified permalink / デバッグ用：指定されたパーマリンクのSpaceID取得をテストする関数
 */
function testGetSpaceId() {
  const permalink = 'https://www.wrike.com/open.htm?id=1282475585';
  console.log(`Test started: / テスト開始: ${permalink}`);

  // Initialize properties / プロパティを初期化
  if (!scriptProperties.getProperty('api_url') || !scriptProperties.getProperty('token')) {
    registKeys();
  }

  const spaceId = getSpaceIdFromPermalink(permalink);
  console.log(`Result: / 結果: ${spaceId}`);

  if (spaceId) {
    SpreadsheetApp.getUi().alert('Success / 成功', `Space ID: / スペースID: ${spaceId}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Failure / 失敗', 'Failed to retrieve space ID / スペースIDの取得に失敗しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
