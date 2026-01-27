const scriptProperties = PropertiesService.getScriptProperties();
// const service = getService_();

/**
 * Register API keys and parameters from spreadsheet / スプレッドシートからAPIキーとパラメータを登録する
 */
function registKeys() {
  scriptProperties.deleteAllProperties(); // Initialize script properties / スクリプトプロパティの初期化
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parametersSheet = spreadsheet.getSheetByName('parameters'); // Specify parameters sheet / パラメーターシートの指定
  scriptProperties.setProperty('api_url', parametersSheet.getRange('B1').getValue());
  scriptProperties.setProperty('client_id', parametersSheet.getRange('B2').getValue());
  scriptProperties.setProperty('client_secret', parametersSheet.getRange('B3').getValue());
  scriptProperties.setProperty('token_url', parametersSheet.getRange('B4').getValue());
  scriptProperties.setProperty('auth_url', parametersSheet.getRange('B5').getValue());
  scriptProperties.setProperty('scope', parametersSheet.getRange('B6').getValue());
  scriptProperties.setProperty('account_Id', parametersSheet.getRange('B7').getValue());
  // scriptProperties.setProperty('token',parametersSheet.getRange('B8').getValue());
  console.log(scriptProperties.getProperty('api_url'));
  console.log(scriptProperties.getProperty('client_id'));
  console.log(scriptProperties.getProperty('client_secret'));
  console.log(scriptProperties.getProperty('token_url'));
  console.log(scriptProperties.getProperty('auth_url'));
  console.log(scriptProperties.getProperty('scope'));
  console.log(scriptProperties.getProperty('account_Id'));
  // console.log('param check\n' + scriptProperties.getProperty('token'));
}

/**
 * Create OAuth2 service for Wrike API / Wrike API用のOAuth2サービスを作成する
 * @return {OAuth2Service} OAuth2 service instance / OAuth2サービスインスタンス
 */
function getService_() {
  return OAuth2.createService('Wrike')
    .setAuthorizationBaseUrl(scriptProperties.getProperty('auth_url'))
    .setTokenUrl(scriptProperties.getProperty('token_url'))
    .setClientId(scriptProperties.getProperty('client_id'))
    .setClientSecret(scriptProperties.getProperty('client_secret'))
    .setScope(scriptProperties.getProperty('scope'))
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCallbackFunction('authCallback')
}

/**
 * Check token status / トークンの状態をチェックする関数
 *
 * @return {Object} Token status {isValid: boolean, message: string} / トークンの状態 {isValid: boolean, message: string}
 */
function checkTokenStatus() {
  try {
    // Get token from script properties / スクリプトプロパティからトークンを取得
    registKeys(); // Update parameters / パラメータを最新化

    const service = getService_();
    const hasAccess = service.hasAccess();

    if (!hasAccess) {
      return {
        isValid: false,
        message: 'Wrike APIのアクセストークンが設定されていないか、期限切れです。認証を行ってください。'
      };
    }

    // Get token and save to script properties / トークンを取得してスクリプトプロパティに保存
    const token = service.getAccessToken();
    scriptProperties.setProperty('token', token);

    return {
      isValid: true,
      message: 'トークンは有効です。'
    };
  } catch (error) {
    console.error('トークンチェック中にエラーが発生しました: ' + error.message);
    return {
      isValid: false,
      message: 'トークンチェック中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * Function executed when spreadsheet is opened / スプレッドシートが開かれたときに実行される関数
 */
function onOpen() {
  // Create menu / メニューの作成
  SpreadsheetApp.getUi()
    .createMenu('Wrike API連携')
    .addItem('⓪Wrike認証', 'showAuth')
    .addItem('①テスト用My Id Info', 'getInformationAboutMe')
    .addItem('②UserIDv4取得', 'getUserIdv4')
    .addItem('③スペース一覧取得', 'getSpaceList')
    .addItem('④フォルダ＆スペース一覧取得', 'GetFolderProjectList')
    .addItem('⑤Get User all user tipe', 'GetUsertypeList')
    .addItem('⑤Update to Collaborator', 'updateUsertoCollaborator')
    .addItem('⑤-2 Bulk Update to Viewer', 'UpdateUsertoViewer')
    .addItem('⑥ユーザー一括招待シート作成', 'initBulkUserInviteSheet')
    .addItem('⑦ユーザー一括招待実行', 'bulkInviteUsers')
    .addItem('⑧スペース間のカスタムフィールドコピー', 'showCopyCustomFieldsDialog')
    .addItem('⑨Backlog→Wrike インポート', 'ImportBacklogToWrike')
    .addItem('⑩Get Timelog', 'GetTimelog')
    // .addItem('🔧SpaceIDテスト', 'testGetSpaceId')
    .addItem('ログアウト', 'logout')
    .addToUi();

  // Set callback URL / コールバックURLの設定
  const scriptId = ScriptApp.getScriptId();
  const url = `https://script.google.com/macros/d/${scriptId}/usercallback`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('parameters');
  sheet.getRange("B11").setValue(url);

  // Automatic authentication process / 自動認証処理
  try {
    const tokenStatus = checkTokenStatus();

    if (!tokenStatus.isValid) {
      // Display message if token is invalid / トークンが無効な場合はメッセージを表示
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Wrike API認証が必要です',
        tokenStatus.message + '\n\n認証処理を実行しますか？',
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        showAuth();
      }
    }
  } catch (error) {
    console.error('自動認証処理中にエラーが発生しました: ' + error.message);
  }
}

/**
 * Create modeless dialog / モーダレスダイアログを作成する関数
 *
 * @param {string} html - HTML to display in dialog / ダイアログに表示するHTML
 * @param {string} title - Dialog title / ダイアログのタイトル
 */
function createModelessDialog(html, title) {
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}

/**
 * Perform Wrike API authentication / Wrike API認証を行う関数
 */
function showAuth() {
  try {
    registKeys();
    var service = getService_();

    if (!service.hasAccess()) {
      // 認証が必要な場合
      var authorizationUrl = service.getAuthorizationUrl();
      var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Wrike APIの認可をします。'
      );
      template.authorizationUrl = authorizationUrl;
      var page = template.evaluate();
      const title = 'Wrikeアプリの認可処理';
      createModelessDialog(page, title);
    } else {
      // 既に認証済みの場合
      const token = service.getAccessToken();
      scriptProperties.setProperty('token', token);

      // 認証済みメッセージを表示
      createModelessDialog(
        'Wrike APIは既に認証されています。トークンが更新されました。',
        '認証状態'
      );

      console.log('access token already provided.\n' + token);
    }
  } catch (error) {
    // エラーが発生した場合
    console.error('認証処理中にエラーが発生しました: ' + error.message);
    createModelessDialog(
      '認証処理中にエラーが発生しました: ' + error.message,
      'エラー'
    );
  }
}

/**
 * 認証コールバック処理を行う関数
 *
 * @param {Object} request - リクエストオブジェクト
 * @return {HtmlOutput} 認証結果を表示するHTML
 */
function authCallback(request) {
  var service = getService_();
  var isAuthorized = service.handleCallback(request);
  var accessToken = service.getAccessToken();

  if (accessToken) {
    // 認証成功
    scriptProperties.setProperty('token', accessToken);
    return HtmlService.createHtmlOutput('認証認可に成功しました! 本タブは閉じてください。');
  } else {
    // 認証失敗または拒否
    return HtmlService.createHtmlOutput('認証認可に失敗しました! 本タブは閉じてください。');
  }
}

/**
 * ログアウト処理を行う関数
 */
function logout() {
  try {
    var service = getService_();
    service.reset();
    scriptProperties.setProperty('token', '');

    const mes = 'Wrikeアプリからログアウトしました。';
    const logoutTitle = 'ログアウト終了';
    createModelessDialog(mes, logoutTitle);
  } catch (error) {
    console.error('ログアウト処理中にエラーが発生しました: ' + error.message);
    createModelessDialog(
      'ログアウト処理中にエラーが発生しました: ' + error.message,
      'エラー'
    );
  }
}
