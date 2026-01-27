# Wrike GAS API Tools

Wrikeの管理タスクを自動化するためのGoogle Apps Script（GAS）ツールコレクションです。このリポジトリには、Wrike APIを使用してユーザー管理、フォルダ/プロジェクト管理、カスタムフィールド操作などを行うための複数のスクリプトが含まれています。

## 機能

- **認証管理**: OAuth2を使用したWrike APIへの認証
- **データインポート**: BacklogデータをWrikeタスクとしてインポート
- **ユーザー管理**: ユーザーIDの取得、ユーザー招待、権限変更、ロール変更（コラボレーターまたはビューアーへ）、グループへの一括割り当て
- **ユーザータイプ管理**: ユーザータイプ一覧とユーザーロール情報の取得
- **グループ管理**: ユーザーのグループへの追加とグループ権限の管理
- **フォルダ/プロジェクト管理**: フォルダとプロジェクトの一覧取得
- **スペース管理**: スペース一覧の取得
- **カスタムフィールド操作**: スペース間でのカスタムフィールドのコピーと管理
- **コンタクト管理**: 全コンタクトとユーザー情報の取得
- **タイムログ管理**: Wrikeのタイムログデータを取得し、タスク名、ユーザー名、カテゴリー、作業時間、日付などを一覧出力

## 構成

- **/*.gs**: Google Apps Scriptファイル（各機能ごとに分割）

## 言語とランタイム

- **言語**: JavaScript (Google Apps Script)
- **ランタイム**: Google Apps Script環境
- **APIバージョン**: Wrike API v4

## 主要ファイル

- **main.gs**: メインスクリプト、認証処理、メニュー設定
- **ImportBacklogToWrike.gs**: BacklogデータをWrikeタスクとしてインポート機能
- **BulkInviteUsers.gs**: ユーザー一括招待機能
- **UpdateUsertoCollaborator.gs**: ユーザーロールをコラボレーターに一括変更
- **UpdateUsertoViewer.gs**: ユーザーロールをビューアーに一括変更
- **GetUserIdv4.gs**: ユーザーIDの取得（APIv2からv4への変換）
- **GetFolderProjectList.gs**: フォルダとプロジェクトの一覧取得
- **GetSpaceList.gs**: スペース一覧の取得
- **GetAllCustomFields.gs**: カスタムフィールド一覧の取得
- **CopyCF.gs**: カスタムフィールドのコピー機能
- **GetAllContacts.gs**: コンタクト一覧の取得
- **Get Information about me.gs**: 現在のユーザー情報取得
- **GetUsertypeList.gs**: ユーザータイプ一覧の取得
- **GetTimelog.gs**: タイムログ取得機能

## セットアップ

### 1. Googleスプレッドシートの作成

新しいGoogleスプレッドシートを作成し、以下のシートを追加してください：

- **parameters**: API接続設定
- **BacklogData**: Backlogインポートデータ（オプション、Backlog→Wrikeインポート用）
- **BulkUserInvite**: ユーザー一括招待用シート
- **Userlist**: ユーザーID変換用シート
- **GetSpaceList**: スペース一覧出力用
- **GetFolderProjectList**: フォルダ/プロジェクト一覧出力用
- **GetAllCustomFields**: カスタムフィールド一覧出力用
- **GetAllContacts**: コンタクト一覧出力用
- **GetInfoAboutMe**: 現在のユーザー情報出力用
- **GetUsertypeList**: ユーザータイプ一覧出力用
- **GetTimeLog**: タイムログ出力用

### 2. parametersシートの設定

**parameters**シートのB列に以下の設定を行ってください：

| 行 | パラメータ | 説明 |
|----|-----------|------|
| B1 | API URL   | Wrike APIのベースURL |
| B2 | Client ID | OAuth2クライアントID |
| B3 | Client Secret | OAuth2クライアントシークレット |
| B4 | Token URL | OAuth2トークンURL |
| B5 | Auth URL  | OAuth2認証URL |
| B6 | Scope     | APIスコープ |
| B7 | Account ID | WrikeアカウントID |

### 3. スクリプトの追加

1. Google Apps Script (script.google.com) を開く
2. 新しいプロジェクトを作成
3. このリポジトリの各.gsファイルをコピー＆ペースト
4. プロジェクトを保存

### 4. OAuth2ライブラリの有効化

1. Google Apps Scriptで「ライブラリ」に移動
2. OAuth2ライブラリを追加: `1B7FSrxWpon2RwlyvwSuGakdnoIRRd3Q9lLWVtbWkMsV4Z1u2wHBJ1iUH`
3. 最新バージョンを選択して保存

## 使用方法

### 認証

1. Googleスプレッドシートを開く
2. 「Wrike API連携」メニューが表示される
3. 「⓪Wrike認証」をクリックしてWrike APIで認証

### Backlogデータをwrikeにインポート

BacklogのデータをWrikeのタスクとしてインポートするには：

1. スプレッドシートに **BacklogData** シートを作成
2. セル **B1** にタスクを作成するWrikeフォルダのパーマリンクを貼り付け
3. 2行目にヘッダーを追加（必須列を含む）：
   - **件名** (タイトル) - 必須
   - **詳細** (説明) - オプション
   - **開始日** (開始日) - オプション（形式：YYYY/MM/DD）
   - **期限日** (期限日) - オプション（形式：YYYY/MM/DD）
   - **コメント1, コメント2, ...** (コメント) - オプション（最大200列まで対応）
4. 3行目からBacklogデータを追加
5. メニュー項目 **⑨Backlog→Wrike インポート** を使用してインポート

#### 同期データの仕様

以下のBacklogデータがWrikeタスクに同期されます：

| Backlog列 | Wrikeフィールド | 動作 |
|---|---|---|
| **件名** (タイトル) | タスクタイトル | 必須、タスク名として使用 |
| **詳細** (説明) | タスクの説明 | タスクの説明フィールドにコピー |
| **開始日** | タスク開始日 | 開始日と期限日の両方が存在する場合は開始日を使用、開始日のみ存在する場合は期限日にも使用、存在しない場合はWrikeのデフォルト動作 |
| **期限日** | タスク期限日 | 両方の日付が存在する場合は期限日を使用、期限日のみ存在する場合は開始日を空にして期限日のみ設定、両方ない場合は日付を設定しない |
| **コメント1-200** | タスクの説明 | 説明フィールドの最後に追加、形式：**【コメント#】** |

例：
```
B1: https://www.wrike.com/folder/123456789
2行目: 件名 | 詳細 | 開始日 | 期限日 | コメント1 | コメント2 | ...
3行目以降: タスクデータ...
```

### 利用可能な機能

スプレッドシートのメニューから以下の機能が利用できます：

- **⓪Wrike認証**: OAuth2を使用してWrike APIで認証
- **①テスト用My Id Info**: 現在のユーザー情報取得（テスト用）
- **②UserIDv4取得**: メールアドレスからユーザーIDをAPI v2からv4形式に変換
- **③スペース一覧取得**: Wrikeの全スペース一覧を取得
- **④フォルダ＆スペース一覧取得**: フォルダとプロジェクト一覧を階層構造で取得
- **⑤Get User all user type**: システム内の全ユーザータイプとロールを取得
- **⑤Update to Collaborator**: 複数ユーザーのロールを一括でコラボレーターに変更
- **⑤-2 Bulk Update to Viewer**: 複数ユーザーのロールを一括でビューアーに変更し、指定グループに追加
- **⑤-3 Email→UserID取得のみ**: メールアドレスからユーザーIDを取得のみ実行（ロール変更なし）
- **⑥ユーザー一括招待シート作成**: ユーザー一括招待用シートの初期化と作成
- **⑦ユーザー一括招待実行**: ユーザーの一括招待をロール割り当て付きで実行
- **⑧スペース間のカスタムフィールドコピー**: スペース間でカスタムフィールドをコピー
- **⑨Backlog→Wrike インポート**: Backlogデータをタイトル、説明、開始日、期限日、コメントと共にWrikeタスクにインポート
- **⑩タイムログ取得**: Wrikeからタイムログデータを取得し、指定されたフォーマットで出力
- **ログアウト**: 認証トークンをクリアしてログアウト

## 設定要件

ツールを使用する前に、以下を確認してください：

1. **Wrike APIアクセス**: APIアクセス権限を持つ有効なWrikeアカウント
2. **OAuth2認証情報**: WrikeからのクライアントIDとクライアントシークレット
3. **適切な権限**: Wrikeアカウントでの適切な権限

## エラーハンドリング

スクリプトには包括的なエラーハンドリングとログ機能が含まれています。問題が発生した場合は、Google Apps Scriptのログで詳細なエラー情報を確認してください。

## 貢献

1. リポジトリをフォーク
2. 機能ブランチを作成
3. 変更を加える
4. 英語と日本語の両方で適切なコメントを追加
5. プルリクエストを送信

## ライセンス

このプロジェクトはオープンソースです。詳細はライセンスファイルを確認してください。

---

# Wrike GAS API Tools

A collection of Google Apps Script (GAS) tools for automating Wrike management tasks. This repository contains multiple scripts for user management, folder/project management, custom field operations, and more using the Wrike API.

## Features

- **Authentication Management**: OAuth2 authentication for Wrike API
- **Data Import**: Import Backlog data as Wrike tasks
- **User Management**: User ID retrieval, bulk user invitation, permission changes, role conversion (to Collaborator or Viewer), bulk group assignment
- **User Type Management**: Retrieve user type lists and user role information
- **Group Management**: Add users to groups and manage group permissions
- **Folder/Project Management**: Retrieve folder and project lists
- **Space Management**: Retrieve space lists
- **Custom Field Operations**: Copy and manage custom fields between spaces
- **Contact Management**: Retrieve all contacts and user information
- **Timelog Management**: Retrieve Wrike timelog data, including task names, user names, categories, work hours, and dates

## Structure

- **/*.gs**: Google Apps Script files (separated by functionality)

## Language & Runtime

- **Language**: JavaScript (Google Apps Script)
- **Runtime**: Google Apps Script environment
- **API Version**: Wrike API v4

## Key Files

- **main.gs**: Main script, authentication processing, menu setup
- **ImportBacklogToWrike.gs**: Import Backlog data as Wrike tasks with full column mapping
- **BulkInviteUsers.gs**: Bulk user invitation functionality with role assignment
- **UpdateUsertoCollaborator.gs**: Bulk user role conversion to Collaborator role
- **UpdateUsertoViewer.gs**: Bulk user role conversion to Viewer role and automatic group assignment
- **GetUserIdv4.gs**: User ID retrieval and conversion (API v2 to v4 format)
- **GetFolderProjectList.gs**: Folder and project list retrieval
- **GetSpaceList.gs**: Space list retrieval
- **GetAllCustomFields.gs**: Custom field list retrieval with metadata
- **CopyCF.gs**: Custom field copy functionality between spaces
- **GetAllContacts.gs**: Contact list retrieval
- **Get Information about me.gs**: Current user information retrieval
- **GetUsertypeList.gs**: User type and role list retrieval
- **GetTimelog.gs**: Timelog retrieval functionality

## Setup

### 1. Create Google Spreadsheet

Create a new Google Spreadsheet and add the following sheets:

- **parameters**: API connection settings
- **BacklogData**: Backlog import data (optional, for Backlog to Wrike import)
- **BulkUserInvite**: Bulk user invitation sheet
- **Userlist**: User ID conversion sheet
- **GetSpaceList**: Space list output
- **GetFolderProjectList**: Folder/project list output
- **GetAllCustomFields**: Custom field list output
- **GetAllContacts**: Contact list output
- **GetInfoAboutMe**: Current user information output
- **GetUsertypeList**: User type list output
- **GetTimeLog**: Timelog output

### 2. Configure Parameters Sheet

In the **parameters** sheet, set up the following configuration in column B:

| Row | Parameter | Description |
|-----|-----------|-------------|
| B1  | API URL   | Wrike API base URL |
| B2  | Client ID | OAuth2 Client ID |
| B3  | Client Secret | OAuth2 Client Secret |
| B4  | Token URL | OAuth2 Token URL |
| B5  | Auth URL  | OAuth2 Authorization URL |
| B6  | Scope     | API Scope |
| B7  | Account ID | Wrike Account ID |

### 3. Add Scripts

1. Open Google Apps Script (script.google.com)
2. Create a new project
3. Copy and paste each .gs file from this repository
4. Save the project

### 4. Enable OAuth2 Library

1. In Google Apps Script, go to Libraries
2. Add the OAuth2 library: `1B7FSrxWpon2RwlyvwSuGakdnoIRRd3Q9lLWVtbWkMsV4Z1u2wHBJ1iUH`
3. Select the latest version and save

## Usage

### Authentication

1. Open your Google Spreadsheet
2. The "Wrike API連携" menu will appear
3. Click "⓪Wrike認証" to authenticate with Wrike API

### Importing Backlog Data to Wrike

To import Backlog data as Wrike tasks:

1. Create a **BacklogData** sheet in your spreadsheet
2. In cell **B1**, paste the Wrike folder permalink where tasks should be created
3. Add headers in row 2, including the following columns:s
   - **詳細** (Description) - Optional
   - **開始日** (Start Date) - Optional (format: YYYY/MM/DD)
   - **期限日** (Due Date) - Optional (format: YYYY/MM/DD)
   - **コメント1, コメント2, ...** (Comments) - Optional (up to 200 comment columns supported)
4. Add your Backlog data starting from row 3
5. Use menu item **⑨Backlog→Wrike インポート** to import

#### Synchronized Data Specifications

The following Backlog data is synced to Wrike tasks:

| Backlog Column | Wrike Field | Behavior |
|---|---|---|
| **件名** (Title) | Task Title | Required, used as task name |
| **詳細** (Description) | Task Description | Copied to task description field |
| **開始日** (Start Date) | Task Start Date | If both start and due dates exist, uses start date; if only start date exists, uses it as due date; if missing, Wrike default applies |
| **期限日** (Due Date) | Task Due Date | If both dates exist, uses due date; if only due date exists, leaves start date empty; if both missing, no dates set |
| **コメント1-20** (Comments) | Task Description | Appended to the end of description field with format: **【Comment#】** |

Example structure:
```
B1: https://www.wrike.com/folder/123456789
Row 2: 件名 | 詳細 | 開始日 | 期限日 | コメント1 | コメント2 | ...
Row 3+: Task data...
```

### Available Functions

The following functions are available through the spreadsheet menu:

- **⓪Wrike認証**: Authenticate with Wrike API using OAuth2
- **①テスト用My Id Info**: Get current user information (for testing purposes)
- **②UserIDv4取得**: Convert user IDs from API v2 to v4 format using email addresses
- **③スペース一覧取得**: Retrieve list of all spaces in Wrike
- **④フォルダ＆スペース一覧取得**: Retrieve folder and project lists with hierarchy
- **⑤Get User all user type**: Get all user types and roles in the system
- **⑤Update to Collaborator**: Bulk update user roles to Collaborator
- **⑤-2 Bulk Update to Viewer**: Bulk update user roles to Viewer and add to specified group
- **⑤-3 Email→UserID取得のみ**: Extract user IDs from email addresses only (no role update)
- **⑥ユーザー一括招待シート作成**: Initialize and create bulk user invitation sheet
- **⑦ユーザー一括招待実行**: Execute bulk user invitations with role assignment
- **⑧スペース間のカスタムフィールドコピー**: Copy custom fields between spaces
- **⑨Backlog→Wrike インポート**: Import Backlog data as Wrike tasks with start/due date and comment synchronization
- **⑩タイムログ取得**: Retrieve timelogs from Wrike and export in a specified format
- **ログアウト**: Logout and clear authentication token

## Configuration Requirements

Before using the tools, ensure you have:

1. **Wrike API Access**: Valid Wrike account with API access
2. **OAuth2 Credentials**: Client ID and Client Secret from Wrike
3. **Proper Permissions**: Appropriate permissions in your Wrike account

## Error Handling

The scripts include comprehensive error handling and logging. Check the Google Apps Script logs for detailed error information if issues occur.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add appropriate comments in both English and Japanese
5. Submit a pull request

## License

This project is open source. Please check the license file for details.
