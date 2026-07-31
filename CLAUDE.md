# CLAUDE.md

このファイルは、Claude Code がこのリポジトリを扱うときのプロジェクト前提と作業ルールです。

## プロジェクト概要

このリポジトリは、kintone上で動作する「Excel帳票出力プラグイン」のプロトタイプです。

目的は、複数のkintoneアプリからログインユーザー権限でデータを取得し、既存Excelテンプレートの元データ用シートへ値を書き込み、見た目重視の `.xlsx` 帳票を出力することです。

APIトークン、パスワード、実データは使用しません。

## 基本方針

- Excel出力形式は `.xlsx`
- Excelテンプレート方式を採用
- 集計シートの数式・罫線・書式・レイアウトはExcel側に残す
- プラグイン側では集計シートの数式を計算しない
- ブラウザ側ライブラリはExcelJSを使用
- kintoneデータ取得は `kintone.api()` を使用
- APIトークンは使わない
- 基準日は通常、実行日の前日を使う。「入力基準日でExcel出力」ボタンではレコードの手入力基準日を使う
- 期間ルールは各日付フィルターの `dateRule` で指定する

## 現在のプロトタイプ範囲

- プラグイン1本
- モード選択
  - テンプレート管理モード
  - 帳票出力モード
- 帳票1種類を想定
- 取得元アプリはJSON設定で複数指定可能
- 帳票ごとにテンプレート管理レコード側の取得元アプリ設定JSONを指定可能
- テンプレート管理アプリでは紐づく帳票出力アプリIDを設定し、抽出条件で出力アプリの任意フィールドを参照できる
- 初期Excelテンプレート作成
- 完成版テンプレート添付の取得
- 設定シートへの出力条件書き込み
- 元データ用シートへのkintoneレコード書き込み
- 集計シートの数式保持
- 500件超のレコード取得にカーソルAPIで対応

## 主なファイル

- `src/desktop.ts`: kintone画面側のメイン処理
- `src/config.ts`: プラグイン設定画面の処理
- `src/shared/excel.ts`: Excelテンプレート生成・編集・ダウンロード
- `src/shared/kintoneApi.ts`: kintone API取得処理
- `src/shared/dateRules.ts`: 基準日・期間ルール計算
- `src/shared/config.ts`: プラグイン設定の読み書き
- テンプレート管理レコードの `sources_json`: 帳票ごとの取得元アプリ設定JSON
- `SOURCES_JSON_EXAMPLES.md`: 取得元アプリ設定JSONの利用者向けサンプル
- テンプレート管理レコード追加/編集画面には、複数取得元アプリ対応のJSON簡易生成UIがある
- 簡易生成UIは、アプリIDからフィールド定義を取得し、型推定、選択、削除、上下移動ができる
- 抽出条件は `filters`、並び順は `sorts` の配列だけを正式仕様として扱う
- 旧 `periodRule` / `additionalQuery` / `valueFrom: "period"` は互換対応しない
- `src/shared/defaults.ts`: 初期設定値
- `plugin/manifest.json`: kintoneプラグインmanifest
- `public/config.html`: プラグイン設定画面HTML
- `public/css/`: 設定画面・デスクトップ画面CSS
- `scripts/build.mjs`: Viteビルド
- `scripts/build-plugin.mjs`: kintoneプラグインzip生成

## 開発コマンド

PowerShellでは `npm` ではなく `npm.cmd` を使うことがあります。

```powershell
npm.cmd install
npm.cmd run typecheck
npm.cmd test
npm.cmd run build
npm.cmd run package
```

成果物:

```text
dist/kintone-excel-report-plugin.zip
```

## 実装時の注意

- 実データ、APIトークン、パスワードをコードやドキュメントへ書かない
- kintoneフィールドコードやアプリIDはダミーまたは設定値として扱う
- Excelの集計シートは原則触らない
- 元データ用シートの行差し替えは `src/shared/excel.ts` に集約する
- 期間ルール追加は `src/shared/dateRules.ts` に集約する
- kintone APIのクエリ生成・カーソルAPI処理は `src/shared/kintoneApi.ts` に集約する
- 仕様拡張時も、まず1帳票・少数アプリで検証できる状態を保つ

## 優先したい次の拡張

相談済みの要望は、`SPEC.md` の「今後の拡張候補」と「次の実装順」を正とする。

直近の優先順:

1. 店舗マスタなどのマスタ参照
2. 出力アプリの基準日/期間初期値と出力履歴
3. テンプレート・設定検証
4. エラー表示と権限チェック
5. 大量データ取得時の進捗表示
