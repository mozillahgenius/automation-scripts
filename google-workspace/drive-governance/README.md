# drive-governance
> Google Drive の命名規則、重複ファイル、権限、アーカイブを包括的に管理するエンタープライズ向けガバナンスシステム

## 概要
Google 共有ドライブ（Daily / External-Share / Archive の3ドライブ構成）を対象とした包括的な管理システムです。ファイル・フォルダの命名規則チェック、重複ファイル検知（MD5ハッシュベース）、未更新ファイルのリマインド、zzzz_ プレフィックスファイルの自動アーカイブ移動、外部共有の権限管理・自動失効、ドライブメンバーの可視化・監査など、9つの主要機能を提供します。全ての管理結果は1つのスプレッドシートの9枚のシートに記録されます。

## 主な機能
- **命名規則チェック**: `XX_`(2桁数字), `YYYYMMDD_`, `zzzz_`, `keep_`, `X.`(1文字+ドット) の5パターンを許可。違反をメール通知
- **禁止語チェック**: ファイル名に「関連」「その他」が含まれる場合を検出
- **Microsoft ファイル形式検出**: Word/Excel/PowerPoint 等16種の拡張子を検出し、Google 形式への変換を推奨
- **重複ファイル検知**: MD5 ハッシュによる同一ファイル検出、スプレッドシート上で黄色ハイライト表示
- **ファイル更新リマインド**: 設定日数以上未更新のファイルをメール通知（keep_/zzzz_ プレフィックスは除外）
- **zzzz_ ファイル自動アーカイブ**: Daily ドライブから Archive ドライブへの自動移動
- **権限自動失効**: ExternalShareLog に登録された有効期限切れの外部共有権限を自動削除
- **ドライブメンバー管理**: 全ドライブの権限情報を DriveMembers シートに記録し、問題権限を赤ハイライト
- **メンバーダッシュボード**: ドライブ種別ごとの権限分布をサマリー表示（外部ユーザー数、Archive 編集権限の警告）
- **権限監査レポート**: Archive への編集権限、外部ユーザーへの編集権限を検出しメール通知
- **緊急外部共有停止**: External-Share ドライブの全外部権限を一括削除する緊急対応機能
- **容量レポート**: 四半期ごとのドライブ別ファイル数・容量レポート生成

## アーキテクチャ
```
[時間トリガー: 毎日実行推奨]
  |
  v
mainCheckExtended()
  |
  +-- scanForDuplicates()
  |     +-- _getTargetFolderIds(): チェック頻度に基づき対象フォルダを選別
  |     +-- _collectFiles(): 再帰的にファイル収集（MD5ハッシュ含む）
  |     +-- _highlightDuplicates(): 重複行を黄色ハイライト
  |
  +-- scanNamingViolations()
  |     +-- _checkFolderRules(): 再帰的に命名規則・禁止語・MSファイルチェック
  |     +-- _createViolationEmailBody(): HTML メールレポート生成
  |     +-- MailApp.sendEmail(): 管理者にメール通知
  |
  +-- sendFileReviewReminders()
  |     +-- _collectReminderFiles(): 閾値超過の未更新ファイルを収集
  |     +-- _sendReminderEmail(): 担当者にHTML形式のリマインドメール送信
  |
  +-- moveZzzzFilesToArchive()
  |     +-- Daily ドライブの zzzz_ ファイルを Archive ドライブへ移動
  |     +-- AuditLog に記録
  |
  +-- revokeExpiredPermissions()
  |     +-- ExternalShareLog の有効期限切れエントリの権限を削除
  |     +-- AuditLog に記録
  |
  +-- updateDriveMembersList()
  |     +-- Drive Permissions API で全ドライブの権限情報を取得
  |     +-- 外部/内部ドメイン判定
  |     +-- 問題権限を赤ハイライト
  |
  +-- updateMemberDashboard()
        +-- ドライブ種別ごとの権限サマリーを生成

[管理用スプレッドシート: 9シート構成]
  +-- FolderConfig: 監視対象フォルダ設定
  +-- DuplicateScan: 重複ファイルスキャン結果
  +-- RuleViolations: 命名規則違反一覧
  +-- FileReminders: リマインド送信履歴
  +-- DriveConfig: 3ドライブ（Daily/External/Archive）設定
  +-- DriveMembers: ドライブ権限一覧
  +-- MemberDashboard: 権限サマリー
  +-- AuditLog: 監査ログ（権限変更、アーカイブ移動等）
  +-- ExternalShareLog: 外部共有の承認・期限管理
```

## ファイル構成
| ファイル | 説明 |
|---|---|
| `Code.js` | 全ロジック（1,891行）: 命名規則、重複検知、リマインド、アーカイブ移動、権限管理、監査、レポート |
| `appsscript.json` | GAS プロジェクト設定（タイムゾーン: Asia/Singapore、V8 ランタイム） |

## 主要関数

### メイン・セットアップ
| 関数名 | 説明 |
|---|---|
| `mainCheckExtended()` | 全管理機能を順次実行するメイン関数。エラー時は管理者にメール通知 |
| `mainCheck()` | 旧バージョン互換。`mainCheckExtended()` を呼び出す |
| `setupExtended()` | 全9シートのヘッダー行・サンプルデータを作成する初期化関数 |
| `setup()` | 旧バージョン互換。`setupExtended()` を呼び出す |

### 重複ファイル検知
| 関数名 | 説明 |
|---|---|
| `scanForDuplicates()` | チェック対象フォルダの全ファイルを収集し、DuplicateScan シートに記録 |
| `_collectFiles(folderId, path, ts, rows)` | 再帰的にファイル情報（ID, 名前, パス, サイズ, MD5, 更新日）を収集 |
| `_highlightDuplicates(sht)` | 同一 MD5 ハッシュの行を黄色（#FFF2CC）でハイライト |

### 命名規則チェック
| 関数名 | 説明 |
|---|---|
| `scanNamingViolations()` | 対象フォルダ内の全ファイル・フォルダの命名規則をチェックし、違反をメール通知 |
| `_isValidNamingConvention(name)` | 5つの命名パターン（2桁数字, 日付8桁, zzzz_, keep_, 1文字+ドット）をチェック |
| `_checkFolderRules(folderId, path, rows, ts)` | 再帰的に命名規則、禁止語、Microsoft ファイル形式をチェック |
| `_checkMicrosoftFileFormat(fileName)` | 16種の Microsoft ファイル拡張子を検出し、推奨 Google 形式を返す |
| `_createViolationEmailBody(violationRows)` | HTML テーブル形式のルール違反レポートメールを生成 |
| `_getViolationRowColor(ruleType)` | 違反種別に応じた行背景色を返す |
| `_createViolationSummary(violationRows)` | 違反種別ごとの件数サマリーを生成 |

### ファイル更新リマインド
| 関数名 | 説明 |
|---|---|
| `sendFileReviewReminders()` | 全監視フォルダのリマインド対象ファイルを確認し、メール送信 |
| `_collectReminderFiles(folder, path, cutoff, reminderFiles)` | 閾値日以上未更新のファイルを再帰収集（keep_/zzzz_ は除外） |
| `_sendReminderEmail(folder, reminderFiles, notifyEmail, reminderThreshold)` | 対応指示付き HTML リマインドメール送信 |
| `_shouldSendReminderEmail(lastReminderDate, reminderFrequency, now)` | リマインド頻度に基づく送信判定 |

### フォルダ設定管理
| 関数名 | 説明 |
|---|---|
| `_getFolderConfigs(configSheet)` | FolderConfig シートから監視フォルダ設定を全件取得 |
| `_getTargetFolderIds(configSheet)` | チェック頻度に基づき今回チェックが必要なフォルダ ID のみ返す |
| `_shouldCheckFolder(lastCheckDate, reminderFrequency, now)` | フォルダのチェック要否を判定 |
| `_getDaysUntilNextCheck(lastCheckDate, reminderFrequency, now)` | 次回チェックまでの残日数を計算 |

### アーカイブ管理
| 関数名 | 説明 |
|---|---|
| `moveZzzzFilesToArchive()` | Daily ドライブの `zzzz_` プレフィックスファイルを Archive ドライブへ自動移動し、AuditLog に記録 |

### 権限管理
| 関数名 | 説明 |
|---|---|
| `revokeExpiredPermissions()` | ExternalShareLog の有効期限切れエントリの権限を自動削除 |
| `updateDriveMembersList()` | Drive Permissions API で全ドライブの権限情報を DriveMembers シートに記録 |
| `updateMemberDashboard()` | ドライブ種別ごとの権限分布サマリーを MemberDashboard シートに生成 |
| `_highlightProblematicPermissions(membersSheet)` | Archive の編集権限、外部ユーザーの編集権限を赤（#ffcdd2）でハイライト |

### 監査・レポート
| 関数名 | 説明 |
|---|---|
| `auditAccounts()` | Archive への編集権限、外部ユーザーへの編集権限を検出しメールレポート送信 |
| `_createAuditReportEmail(issues)` | 監査レポートの HTML メール本文を生成 |
| `emergencyDisableExternalSharing()` | External-Share ドライブの全外部権限を一括削除する緊急対応関数 |
| `recordExternalShare(fileId, email, permission, expiryDays, approver)` | 外部共有を ExternalShareLog に手動記録 |
| `generateStorageReport()` | 四半期容量レポートをメール送信 |
| `generateReportSummary()` | 全シートの統計サマリーを生成（違反数、リマインド数、重複数、メンバー数、監査指摘数） |

### ヘルパー・ユーティリティ
| 関数名 | 説明 |
|---|---|
| `_getDriveConfigs(driveConfigSheet)` | DriveConfig シートからドライブ設定を取得 |
| `_countFilesInFolder(folder)` | フォルダ内のファイル数を再帰的にカウント |
| `_countDuplicateFiles(dupSheet)` | 重複ファイル数（同一MD5が2件以上）をカウント |
| `_countRecentAuditIssues(auditSheet)` | 過去30日間の監査指摘件数をカウント |
| `_getSheetGid(sheetName)` | シートの GID を取得（メール内リンク用） |
| `validateConfiguration()` | フォルダ・ドライブ設定の妥当性チェック（アクセス可否、メールアドレス有効性など） |
| `checkAndGrantSpreadsheetAccess()` | 管理スプレッドシートの権限情報を確認 |
| `grantEditAccessToEmail(email)` | 指定メールアドレスにスプレッドシートの編集権限を付与 |

### テスト用関数
| 関数名 | 説明 |
|---|---|
| `testFileReminders()` | リマインド機能のテスト実行 |
| `testRuleCheck()` | ルール違反チェックのテスト実行 |
| `testDuplicateCheck()` | 重複ファイルスキャンのテスト実行 |
| `testArchiveMove()` | アーカイブ移動のテスト実行 |
| `testMemberManagement()` | メンバー管理機能のテスト実行 |
| `showCheckStatus()` | 全フォルダのチェックステータスをログ出力 |
| `forceCheckAllFolders()` | 全フォルダの最終チェック日時をクリアし、強制的に全チェック実行 |

## 使用サービス・API
- **Google Drive API** (DriveApp + Drive Advanced Service) - ファイル検索、権限管理、ファイル移動
- **Google Sheets API** (SpreadsheetApp) - 9シートの管理データ読み書き
- **Google Mail API** (MailApp) - ルール違反通知、リマインド、監査レポート、緊急通知メール送信
- **Drive Permissions API** (Drive.Permissions) - ドライブ権限の一覧取得・削除
- **Session Service** (Session) - 実行ユーザーのメールアドレス取得（内部/外部判定用）

## セットアップ手順
1. Google Apps Script プロジェクトを作成し、`Code.js` をコピー
2. Apps Script エディタで「サービス」から **Drive API** を追加
3. 紐づける GCP プロジェクトで **Drive API** と **Admin SDK Directory API** を有効化
4. `CONFIG_SHEET_ID` を管理用スプレッドシートの ID に変更
5. `ADMIN_EMAIL` を管理者のメールアドレスに変更
6. スクリプト実行者が対象の共有ドライブ全ての「管理者」権限を持っていることを確認
7. `setupExtended()` を実行して全9シートを初期化
8. FolderConfig シートに監視対象フォルダ ID と通知先メールを設定
9. DriveConfig シートに Daily / External-Share / Archive の3ドライブ ID を設定
10. `mainCheckExtended()` を時間トリガー（毎日実行推奨）に登録

## 設定定数
| 定数名 | 説明 |
|---|---|
| `CONFIG_SHEET_ID` | 管理用スプレッドシートの ID |
| `CONFIG_SHEET_NAME` | `FolderConfig` - 監視フォルダ設定シート |
| `DUP_SHEET_NAME` | `DuplicateScan` - 重複スキャン結果シート |
| `VIOL_SHEET_NAME` | `RuleViolations` - 命名規則違反シート |
| `REMINDER_SHEET_NAME` | `FileReminders` - リマインド送信履歴シート |
| `DRIVE_CONFIG_SHEET_NAME` | `DriveConfig` - ドライブ設定シート |
| `DRIVE_MEMBERS_SHEET_NAME` | `DriveMembers` - ドライブ権限一覧シート |
| `MEMBER_DASHBOARD_SHEET_NAME` | `MemberDashboard` - 権限サマリーシート |
| `AUDIT_LOG_SHEET_NAME` | `AuditLog` - 監査ログシート |
| `EXTERNAL_SHARE_SHEET_NAME` | `ExternalShareLog` - 外部共有管理シート |
| `ADMIN_EMAIL` | 管理者メールアドレス |
| `DRIVE_TYPES` | `{ DAILY, EXTERNAL, ARCHIVE }` - ドライブ種別 |
| `PERMISSION_EXPIRY_DAYS` | `{ CONTRIBUTOR: 90, VIEWER: 30 }` - 権限自動失効日数 |

## スプレッドシート構成（9シート）

### FolderConfig
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 監視対象フォルダID | Google Drive フォルダ ID |
| B | 通知先メール | リマインドメール送信先 |
| C | リマインド閾値（日） | 未更新日数の閾値（デフォルト60日） |
| D | リマインド頻度（日） | リマインド送信間隔（デフォルト7日） |
| E | 最終リマインド送信日時 | 自動更新 |
| F | 最終チェック日時 | 自動更新 |
| G | 備考 | 自由記入 |

### DuplicateScan
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | スキャン日時 | スキャン実行日時 |
| B | ファイルID | Google Drive ファイル ID |
| C | ファイル名 | ファイル名 |
| D | フォルダパス | ファイルのフルパス |
| E | サイズ(byte) | ファイルサイズ |
| F | MD5ハッシュ | ファイルの MD5 チェックサム（重複判定キー） |
| G | 最終更新日 | ファイル最終更新日 |

### RuleViolations
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 検知日時 | 違反検知日時 |
| B | パス | ファイル/フォルダのパス |
| C | 種別 | 「ファイル」または「フォルダ」 |
| D | 名前 | ファイル/フォルダ名 |
| E | ルール種別 | 「命名規則違反」または「Microsoftファイル形式」 |
| F | 詳細 | 違反内容の詳細説明 |
| G | URL | Google Drive 上の URL |

### FileReminders
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 送信日時 | リマインドメール送信日時 |
| B | 対象フォルダ | フォルダ名 |
| C | ファイル名 | ファイル名 |
| D | 最終更新日 | ファイル最終更新日 |
| E | 経過日数 | 最終更新からの経過日数 |
| F | ファイルパス | ファイルのフルパス |
| G | ファイルURL | Google Drive 上の URL |
| H | 送信先メール | リマインドメール送信先 |

### DriveConfig
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | ドライブ種別 | `Daily` / `External-Share` / `Archive` |
| B | ドライブID | 共有ドライブ ID |
| C | ドライブ名 | 表示名 |
| D | Gemini検索 | 「はい」/「いいえ」 |
| E | 備考 | ドライブの用途説明 |

### DriveMembers
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 更新日時 | 権限情報取得日時 |
| B | ドライブID | 共有ドライブ ID |
| C | ドライブ名 | 表示名 |
| D | メールアドレス | ユーザーのメールアドレス |
| E | 権限タイプ | `owner` / `writer` / `reader` 等 |
| F | 最終アクセス | （将来使用） |
| G | 有効期限 | 権限の有効期限（設定されている場合） |
| H | 外部ドメイン | 「外部」/「内部」 |

### MemberDashboard
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 更新日時 / サマリー内容 | ドライブ種別ごとの権限分布サマリー |
| B | サマリー | 人数等 |

### AuditLog
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 監査日時 | アクション実行日時 |
| B | ドライブ名 | 対象ドライブ |
| C | 対象者 | メールアドレス |
| D | 権限 | 権限タイプまたはアクション種別 |
| E | アクション | `zzzz_自動アーカイブ` / `権限失効` / `監査指摘` / `緊急権限削除` 等 |
| F | 理由 | アクションの理由 |
| G | 実行者 | `システム自動` / `システム監査` / ユーザーメール |

### ExternalShareLog
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 共有日時 | 外部共有を承認した日時 |
| B | ファイルID | Google Drive ファイル ID |
| C | ファイル名 | ファイル名 |
| D | 共有先メール | 外部共有先のメールアドレス |
| E | 権限 | 付与した権限タイプ |
| F | 有効期限 | 共有の有効期限（自動失効の判定に使用） |
| G | 承認者 | 共有を承認したユーザー |
| H | 備考 | 失効時は「失効済み」が記録される |

## 命名規則ルール
| パターン | 正規表現 | 例 |
|---|---|---|
| 2桁数字プレフィックス | `^\d{2}_.+` | `01_議事録.docx` |
| 日付プレフィックス | `^\d{8}_.+` (1900-2100年の有効な日付) | `20240101_レポート.pdf` |
| アーカイブプレフィックス | `^zzzz_.+` (大文字小文字不問) | `zzzz_旧版資料.xlsx` |
| 永久保存プレフィックス | `^keep_.+` (大文字小文字不問) | `keep_契約書.pdf` |
| 1文字+ドット | `^[a-zA-Z]\..+` | `A.設計書.pdf` |

## トリガー設定
| 対象関数 | 推奨間隔 | 説明 |
|---|---|---|
| `mainCheckExtended()` | 毎日1回 | 全管理機能の定期実行 |
| `auditAccounts()` | 週1回（任意） | 権限監査レポート |
| `generateStorageReport()` | 四半期1回（任意） | 容量レポート |
