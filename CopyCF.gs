/**
 * カスタムフィールドをスペース間でコピーするためのGAS
 * スペーススコープのカスタムフィールドを別のスペースにコピーします
 */

/**
 * パーマリンクからSpaceIDを取得する関数
 * 
 * @param {string} permalink - Wrikeスペースのパーマリンク
 * @return {string|null} SpaceID または null（エラーの場合）
 */
function getSpaceIdFromPermalink(permalink) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');
    
    if (!apiUrl || !token) {
      throw new Error('APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    const apiEndpoint = `${apiUrl}/folders?permalink=${encodeURIComponent(permalink)}`;
    console.log(`APIエンドポイント: ${apiEndpoint}`);
    
    const response = UrlFetchApp.fetch(apiEndpoint, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
      muteHttpExceptions: true
    });

    console.log(`レスポンスコード: ${response.getResponseCode()}`);
    const responseText = response.getContentText();
    console.log(`レスポンス内容: ${responseText}`);

    if (response.getResponseCode() !== 200) {
      // logical folderエラーの場合は、IDを抽出してスペース一覧から検索
      if (responseText.includes('logical folder')) {
        console.log('Logical folderエラーが発生。スペース一覧から検索します。');
        
        // パーマリンクからIDを抽出
        const idMatch = permalink.match(/id=(\d+)/);
        if (!idMatch) {
          throw new Error('パーマリンクからIDを抽出できませんでした。');
        }
        
        const folderId = idMatch[1];
        console.log(`抽出されたフォルダID: ${folderId}`);

        // スペース一覧を取得
        const spacesResponse = UrlFetchApp.fetch(`${apiUrl}/folders?project=false`, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (spacesResponse.getResponseCode() === 200) {
          const spacesJson = JSON.parse(spacesResponse.getContentText());
          const allFolders = spacesJson.data;
          const spaces = allFolders.filter(f => f.scope === 'WsSpace');
          
          console.log(`${spaces.length} 個のスペースを検索中...`);

          // IDが一致するスペースを検索
          for (const space of spaces) {
            if (space.id === folderId) {
              console.log(`一致するスペースが見つかりました: ${space.title} (ID: ${space.id})`);
              return space.id;
            }
          }
        }
      }
      
      throw new Error(`API呼び出しに失敗しました。レスポンスコード: ${response.getResponseCode()}, エラー: ${responseText}`);
    }

    const myJson = JSON.parse(responseText);
    const data = myJson.data;

    if (!data || data.length === 0) {
      throw new Error('指定されたパーマリンクのフォルダが見つかりません。');
    }

    // パーマリンクで取得されたフォルダ情報
    const folder = data[0];
    console.log(`取得されたフォルダ: ${folder.title}, スコープ: ${folder.scope}, ID: ${folder.id}`);
    
    // パーマリンクで取得されたフォルダのIDがスペースIDとして使用される
    // （Wrikeの仕様により、パーマリンククエリで返されるIDがスペースID）
    console.log(`パーマリンクで取得されたIDをスペースIDとして使用: ${folder.id}`);
    return folder.id;
    
  } catch (error) {
    console.error('SpaceID取得エラー:', error.message);
    SpreadsheetApp.getUi().alert('エラー', `SpaceID取得に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * 指定されたスペースのカスタムフィールドを取得する関数
 * 
 * @param {string} spaceId - スペースID
 * @return {Array|null} カスタムフィールド配列 または null（エラーの場合）
 */
function getCustomFieldsBySpaceId(spaceId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');
    
    if (!apiUrl || !token) {
      throw new Error('APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    const apiEndpoint = `${apiUrl}/customfields`;
    
    const response = UrlFetchApp.fetch(apiEndpoint, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`API呼び出しに失敗しました。レスポンスコード: ${response.getResponseCode()}`);
    }

    const myJson = JSON.parse(response.getContentText());
    const allCustomFields = myJson.data;

    // 指定されたスペースIDのカスタムフィールドのみをフィルタリング
    const spaceCustomFields = allCustomFields.filter(field => field.spaceId === spaceId);
    
    // コピー不可能なフィールドタイプを除外（計算フィールド、データベースリンクなど）
    const unsupportedTypes = ['CalculatedNumeric', 'CalculatedDate', 'LinkToDatabase'];
    const copyableFields = spaceCustomFields.filter(field => !unsupportedTypes.includes(field.type));
    
    const skippedCount = spaceCustomFields.length - copyableFields.length;
    
    console.log(`スペース ${spaceId} のカスタムフィールドを ${spaceCustomFields.length} 件取得しました`);
    if (skippedCount > 0) {
      console.log(`${skippedCount} 件のフィールドはコピー対象外です（計算フィールド、データベースリンクなど）`);
    }
    console.log(`コピー可能なフィールド: ${copyableFields.length} 件`);
    
    return copyableFields;
    
  } catch (error) {
    console.error('カスタムフィールド取得エラー:', error.message);
    SpreadsheetApp.getUi().alert('エラー', `カスタムフィールド取得に失敗しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * カスタムフィールドを新しいスペースに作成する関数
 * 
 * @param {Object} customField - 元のカスタムフィールドオブジェクト
 * @param {string} targetSpaceId - コピー先のスペースID
 * @return {Object|null} 作成されたカスタムフィールド または null（エラーの場合）
 */
function createCustomField(customField, targetSpaceId) {
  try {
    const apiUrl = scriptProperties.getProperty('api_url');
    const token = scriptProperties.getProperty('token');
    
    if (!apiUrl || !token) {
      throw new Error('APIの設定またはトークンが見つかりません。認証を確認してください。');
    }

    // カスタムフィールド作成用のペイロードを構築
    const payload = {
      title: customField.title,
      type: customField.type,
      spaceId: targetSpaceId
    };

    // 説明があれば追加
    if (customField.description) {
      payload.description = customField.description;
    }

    // 設定（settings）があれば、settingsオブジェクトとして追加
    if (customField.settings) {
      payload.settings = {};
      
      // settingsオブジェクトの各プロパティをsettings内に追加
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
      throw new Error(`カスタムフィールド作成に失敗しました。レスポンスコード: ${response.getResponseCode()}, エラー: ${errorText}`);
    }

    const myJson = JSON.parse(response.getContentText());
    const createdField = myJson.data[0];
    
    console.log(`カスタムフィールド "${customField.title}" を作成しました。ID: ${createdField.id}`);
    
    // レート制限対応のための待機
    Utilities.sleep(200);
    
    return createdField;
    
  } catch (error) {
    console.error('カスタムフィールド作成エラー:', error.message);
    throw error;
  }
}

/**
 * カスタムフィールドをスペース間でコピーするメイン関数
 */
function copyCustomFieldsBetweenSpaces() {
  try {
    // UI入力を取得
    const ui = SpreadsheetApp.getUi();
    
    const sourcePermalinkResponse = ui.prompt(
      'カスタムフィールドコピー',
      'コピー元スペースのパーマリンクを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sourcePermalinkResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const sourcePermalink = sourcePermalinkResponse.getResponseText().trim();
    if (!sourcePermalink) {
      ui.alert('エラー', 'コピー元スペースのパーマリンクが入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    const targetPermalinkResponse = ui.prompt(
      'カスタムフィールドコピー',
      'コピー先スペースのパーマリンクを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (targetPermalinkResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const targetPermalink = targetPermalinkResponse.getResponseText().trim();
    if (!targetPermalink) {
      ui.alert('エラー', 'コピー先スペースのパーマリンクが入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // プロパティを初期化
    if (!scriptProperties.getProperty('api_url') || !scriptProperties.getProperty('token')) {
      registKeys();
    }
    
    // 処理開始メッセージ
    createModelessDialog('カスタムフィールドコピー処理を開始します...', '処理中');
    
    // コピー元スペースIDを取得
    console.log('コピー元スペースIDを取得中...');
    const sourceSpaceId = getSpaceIdFromPermalink(sourcePermalink);
    if (!sourceSpaceId) {
      return;
    }
    
    // コピー先スペースIDを取得
    console.log('コピー先スペースIDを取得中...');
    const targetSpaceId = getSpaceIdFromPermalink(targetPermalink);
    if (!targetSpaceId) {
      return;
    }
    
    // 同じスペースの場合はエラー
    if (sourceSpaceId === targetSpaceId) {
      ui.alert('エラー', 'コピー元とコピー先が同じスペースです。異なるスペースを指定してください。', ui.ButtonSet.OK);
      return;
    }
    
    // コピー元スペースのカスタムフィールドを取得
    console.log('コピー元スペースのカスタムフィールドを取得中...');
    const sourceCustomFields = getCustomFieldsBySpaceId(sourceSpaceId);
    if (!sourceCustomFields) {
      return;
    }
    
    if (sourceCustomFields.length === 0) {
      ui.alert('情報', 'コピー元スペースにカスタムフィールドが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // 確認ダイアログ
    const confirmResponse = ui.alert(
      '確認',
      `${sourceCustomFields.length} 個のカスタムフィールドをコピーしますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // スキップされたフィールド数を計算
    const allCustomFieldsResponse = UrlFetchApp.fetch(`${scriptProperties.getProperty('api_url')}/customfields`, {
      headers: { Authorization: `Bearer ${scriptProperties.getProperty('token')}` }
    });
    const allFieldsJson = JSON.parse(allCustomFieldsResponse.getContentText());
    const allSourceFields = allFieldsJson.data.filter(field => field.spaceId === sourceSpaceId);
    const skippedCount = allSourceFields.length - sourceCustomFields.length;
    
    // カスタムフィールドを順次コピー
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    console.log('カスタムフィールドのコピーを開始...');
    
    for (const customField of sourceCustomFields) {
      try {
        console.log(`カスタムフィールド "${customField.title}" をコピー中...`);
        createCustomField(customField, targetSpaceId);
        successCount++;
      } catch (error) {
        console.error(`カスタムフィールド "${customField.title}" のコピーに失敗:`, error.message);
        errorCount++;
        errors.push(`${customField.title}: ${error.message}`);
      }
    }
    
    // 結果メッセージ
    let resultMessage = `カスタムフィールドコピー完了\n\n`;
    resultMessage += `成功: ${successCount} 個\n`;
    resultMessage += `失敗: ${errorCount} 個`;
    if (skippedCount > 0) {
      resultMessage += `\nスキップ: ${skippedCount} 個 (計算フィールド、データベースリンクなど)`;
    }
    
    if (errors.length > 0) {
      resultMessage += `\n\n失敗詳細:\n${errors.slice(0, 5).join('\n')}`;
      if (errors.length > 5) {
        resultMessage += `\n... 他 ${errors.length - 5} 件`;
      }
    }
    
    // 完了メッセージをダイアログで表示
    const completionHtml = `
      <div style="padding: 20px;">
        <h3>カスタムフィールドコピー完了</h3>
        <pre style="white-space: pre-wrap; font-family: monospace; font-size: 12px;">${resultMessage}</pre>
      </div>
    `;
    createModelessDialog(completionHtml, '完了');
    console.log('カスタムフィールドコピー処理完了');
    
  } catch (error) {
    console.error('メイン処理エラー:', error.message);
    SpreadsheetApp.getUi().alert('エラー', `処理中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * メニューからカスタムフィールドコピーを実行するための関数
 */
function showCopyCustomFieldsDialog() {
  copyCustomFieldsBetweenSpaces();
}

/**
 * デバッグ用：指定されたパーマリンクのSpaceID取得をテストする関数
 */
function testGetSpaceId() {
  const permalink = 'https://www.wrike.com/open.htm?id=1282475585';
  console.log(`テスト開始: ${permalink}`);
  
  // プロパティを初期化
  if (!scriptProperties.getProperty('api_url') || !scriptProperties.getProperty('token')) {
    registKeys();
  }
  
  const spaceId = getSpaceIdFromPermalink(permalink);
  console.log(`結果: ${spaceId}`);
  
  if (spaceId) {
    SpreadsheetApp.getUi().alert('成功', `スペースID: ${spaceId}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('失敗', 'スペースIDの取得に失敗しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}