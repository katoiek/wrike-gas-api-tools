---
description: Repository Information Overview
alwaysApply: true
---

# Wrike GAS API Tools Information

## Summary
Wrikeの管理タスクを自動化するためのGoogle Apps Script（GAS）ツールコレクションです。このリポジトリには、Wrike APIを使用してユーザー管理、フォルダ/プロジェクト管理、カスタムフィールド操作などを行うための複数のスクリプトが含まれています。

## Structure
- **/*.gs**: Google Apps Scriptファイル（各機能ごとに分割）
- **notes/**: プロジェクト関連のメモや補足情報
- **.vscode/**: VSCode設定ファイル

## Language & Runtime
**Language**: JavaScript (Google Apps Script)
**Runtime**: Google Apps Script環境
**API Version**: Wrike API v4

## Main Features
- **認証管理**: OAuth2を使用したWrike APIへの認証
- **ユーザー管理**: ユーザーIDの取得、ユーザー招待、権限変更
- **フォルダ/プロジェクト管理**: フォルダとプロジェクトの一覧取得
- **スペース管理**: スペース一覧の取得
- **カスタムフィールド操作**: カスタムフィールドのコピーと管理

## Key Files
- **main.gs**: メインスクリプト、認証処理、メニュー設定
- **BulkInviteUsers.gs**: ユーザー一括招待機能
- **GetUserIdv4.gs**: ユーザーIDの取得（APIv2からv4への変換）
- **GetFolderProjectList.gs**: フォルダとプロジェクトの一覧取得
- **GetSpaceList.gs**: スペース一覧の取得
- **GetAllCustomFields.gs**: カスタムフィールド一覧の取得
- **CopyCF.gs**: カスタムフィールドのコピー機能

## Usage & Operations
このツールはGoogle Spreadsheetのアドオンとして動作します。スプレッドシートを開くと「Wrike API連携」メニューが表示され、以下の機能が利用可能になります：

1. Wrike認証: OAuth2認証を実行
2. ユーザー管理機能: ユーザーIDの取得、一括招待、権限変更
3. フォルダ/プロジェクト管理: 一覧取得
4. スペース管理: 一覧取得
5. カスタムフィールド操作: 取得、コピー

## Configuration
スプレッドシート内の「parameters」シートに以下の設定が必要です：
- API URL
- クライアントID
- クライアントシークレット
- トークンURL
- 認証URL
- スコープ
- アカウントID

## Data Structure
- **BulkUserInvite**: ユーザー一括招待用のシート
- **Userlist**: ユーザーID変換用のシート
- **parameters**: API接続設定用のシート

## Authentication Flow
1. OAuth2認証を使用してWrike APIに接続
2. 認証トークンをスクリプトプロパティに保存
3. トークンの有効期限が切れた場合は自動的に再認証を促す
