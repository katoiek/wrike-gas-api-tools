# Wrike GAS API Tools

A collection of Google Apps Script (GAS) tools for automating Wrike management tasks. This repository contains multiple scripts for user management, folder/project management, custom field operations, and more using the Wrike API.

## Features

- **Authentication Management**: OAuth2 authentication for Wrike API
- **User Management**: User ID retrieval, bulk user invitation, permission changes
- **Folder/Project Management**: Retrieve folder and project lists
- **Space Management**: Retrieve space lists
- **Custom Field Operations**: Copy and manage custom fields
- **Contact Management**: Retrieve all contacts and user information

## Structure

- **/*.gs**: Google Apps Script files (separated by functionality)

## Language & Runtime

- **Language**: JavaScript (Google Apps Script)
- **Runtime**: Google Apps Script environment
- **API Version**: Wrike API v4

## Key Files

- **main.gs**: Main script, authentication processing, menu setup
- **BulkInviteUsers.gs**: Bulk user invitation functionality
- **GetUserIdv4.gs**: User ID retrieval (API v2 to v4 conversion)
- **GetFolderProjectList.gs**: Folder and project list retrieval
- **GetSpaceList.gs**: Space list retrieval
- **GetAllCustomFields.gs**: Custom field list retrieval
- **CopyCF.gs**: Custom field copy functionality
- **GetAllContacts.gs**: Contact list retrieval
- **Get Information about me.gs**: Current user information retrieval
- **GetUsertypeList.gs**: User type list retrieval

## Setup

### 1. Create Google Spreadsheet

Create a new Google Spreadsheet and add the following sheets:

- **parameters**: API connection settings
- **BulkUserInvite**: Bulk user invitation sheet
- **Userlist**: User ID conversion sheet
- **GetSpaceList**: Space list output
- **GetFolderProjectList**: Folder/project list output
- **GetAllCustomFields**: Custom field list output
- **GetAllContacts**: Contact list output
- **GetInfoAboutMe**: Current user information output
- **GetUsertypeList**: User type list output

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

### Available Functions

The following functions are available through the spreadsheet menu:

- **⓪Wrike認証**: Wrike API authentication
- **①テスト用My Id Info**: Get current user information (for testing)
- **②UserIDv4取得**: Convert user IDs from API v2 to v4 format
- **③スペース一覧取得**: Retrieve space list
- **④フォルダ＆スペース一覧取得**: Retrieve folder and project list
- **⑤Update to Collaborator**: Update user permissions to collaborator
- **⑥ユーザー一括招待シート作成**: Create bulk user invitation sheet
- **⑦ユーザー一括招待実行**: Execute bulk user invitation
- **⑧カスタムフィールドコピー**: Copy custom fields

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

---

# Wrike GAS API Tools

Wrikeの管理タスクを自動化するためのGoogle Apps Script（GAS）ツールコレクションです。このリポジトリには、Wrike APIを使用してユーザー管理、フォルダ/プロジェクト管理、カスタムフィールド操作などを行うための複数のスクリプトが含まれています。

## 機能

- **認証管理**: OAuth2を使用したWrike APIへの認証
- **ユーザー管理**: ユーザーIDの取得、ユーザー招待、権限変更
- **フォルダ/プロジェクト管理**: フォルダとプロジェクトの一覧取得
- **スペース管理**: スペース一覧の取得
- **カスタムフィールド操作**: カスタムフィールドのコピーと管理
- **コンタクト管理**: 全コンタクトとユーザー情報の取得

## 構成

- **/*.gs**: Google Apps Scriptファイル（各機能ごとに分割）

## 言語とランタイム

- **言語**: JavaScript (Google Apps Script)
- **ランタイム**: Google Apps Script環境
- **APIバージョン**: Wrike API v4

## 主要ファイル

- **main.gs**: メインスクリプト、認証処理、メニュー設定
- **BulkInviteUsers.gs**: ユーザー一括招待機能
- **GetUserIdv4.gs**: ユーザーIDの取得（APIv2からv4への変換）
- **GetFolderProjectList.gs**: フォルダとプロジェクトの一覧取得
- **GetSpaceList.gs**: スペース一覧の取得
- **GetAllCustomFields.gs**: カスタムフィールド一覧の取得
- **CopyCF.gs**: カスタムフィールドのコピー機能
- **GetAllContacts.gs**: コンタクト一覧の取得
- **Get Information about me.gs**: 現在のユーザー情報取得
- **GetUsertypeList.gs**: ユーザータイプ一覧の取得

## セットアップ

### 1. Googleスプレッドシートの作成

新しいGoogleスプレッドシートを作成し、以下のシートを追加してください：

- **parameters**: API接続設定
- **BulkUserInvite**: ユーザー一括招待用シート
- **Userlist**: ユーザーID変換用シート
- **GetSpaceList**: スペース一覧出力用
- **GetFolderProjectList**: フォルダ/プロジェクト一覧出力用
- **GetAllCustomFields**: カスタムフィールド一覧出力用
- **GetAllContacts**: コンタクト一覧出力用
- **GetInfoAboutMe**: 現在のユーザー情報出力用
- **GetUsertypeList**: ユーザータイプ一覧出力用

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

### 利用可能な機能

スプレッドシートのメニューから以下の機能が利用できます：

- **⓪Wrike認証**: Wrike API認証
- **①テスト用My Id Info**: 現在のユーザー情報取得（テスト用）
- **②UserIDv4取得**: ユーザーIDをAPI v2からv4形式に変換
- **③スペース一覧取得**: スペース一覧の取得
- **④フォルダ＆スペース一覧取得**: フォルダとプロジェクト一覧の取得
- **⑤Update to Collaborator**: ユーザー権限をコラボレーターに更新
- **⑥ユーザー一括招待シート作成**: ユーザー一括招待用シートの作成
- **⑦ユーザー一括招待実行**: ユーザー一括招待の実行
- **⑧カスタムフィールドコピー**: カスタムフィールドのコピー

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
