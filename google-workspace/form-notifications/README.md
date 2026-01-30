# form-notifications / Google Form 通知システム（教育移住・ベビーシッター）

## 概要

Google フォームの回答送信時に、自動で関係者にメール通知を送信するシステム。教育移住相談フォーム、顧問サービス相談フォーム、ベビーシッターフォームの 3 種類に対応。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Forms API, Gmail API |
| Triggers | フォーム送信時（onFormSubmit） |

## ファイル構成

- `babysitter-form-notification.gs` - ベビーシッターフォームの通知処理
- `Googleform(顧問用)/formNotification.gs` - 顧問サービス相談フォームの通知処理
- `Googleform(今後移住想定)/formNotification.gs` - 教育移住相談フォームの通知処理

## セットアップ

1. 各 Google フォームに紐づくスプレッドシートのスクリプトエディタを開く
2. 対応する `.gs` ファイルの内容をコピー
3. 通知先メールアドレスをスクリプト内に設定
4. フォーム送信時トリガー（`onFormSubmit`）を設定
5. 各フォームごとに上記手順を繰り返す

## 使い方

- フォームが送信されると、トリガーにより自動的に通知処理が実行される
- 回答内容を整形したメールが指定の宛先に送信される
- 各フォームの通知テンプレート・宛先はスクリプト内で個別に設定可能
