# AGENTS.md

このファイルは、Codex などのコーディングエージェントがこのリポジトリを扱うときの作業ガイドです。

## 目的

kintoneの複数アプリからデータを取得し、Excelテンプレートへ流し込む帳票出力プラグインを開発する。

既存Excelの見た目を重視するため、集計シートの数式・書式・レイアウトはExcel側に残し、プラグインは元データ用シートへの値書き込みを担当する。

## 作業原則

- APIトークンや認証情報は扱わない
- `kintone.api()` を使い、ログインユーザー権限で取得する
- Excel数式の再計算はExcel側に任せる
- 集計シートは原則変更しない
- 基準日は通常、実行日の前日を使う。「入力基準日でExcel出力」ボタンではレコードの手入力基準日を使う
- 期間ルールは取得元全体ではなく、各日付フィルターの `dateRule` に持たせる
- 変更は既存の責務分割に沿って小さく行う
- TypeScriptの型チェックとパッケージ作成を確認してから完了する

## バージョン運用

- バージョン更新とナンバリングは、変更内容に応じてAIエージェントが判断して行う
- 原則としてSemVerを目安にする
  - patch: 不具合修正、小さな文言修正
  - minor: 後方互換のある機能追加、設定項目追加
  - major: 既存設定や利用手順に破壊的変更がある場合
- バージョン更新時は `package.json`、`package-lock.json`、`plugin/manifest.json`、`CHANGELOG.md`、必要に応じてREADME/SPEC内の成果物名を揃える
- `npm.cmd run package` でバージョン付きzipと最新版zipの両方を確認する

## アーキテクチャ

```text
src/
  desktop.ts              kintoneレコード画面の処理
  config.ts               プラグイン設定画面の処理
  shared/
    config.ts             設定のparse/serialize
    dateRules.ts          基準日・期間ルール
    defaults.ts           初期設定
    excel.ts              Excel生成・編集
    kintoneApi.ts         kintone API処理
    types.ts              共通型
```

## モード

### テンプレート管理モード

テンプレート管理アプリに適用する。

現在の主機能:

- 初期Excelテンプレート作成
- 設定シート、元データ用シート、集計シートの生成

### 帳票出力モード

帳票出力アプリに適用する。

現在の主機能:

- 帳票種類、店舗をレコードから取得し、基準日は出力時に実行日の前日として確定
- テンプレート管理アプリから完成版Excelを取得
- 取得元アプリの日付フィルターごとに期間計算
- テンプレート管理レコード側の取得元アプリ設定JSON取得
- kintoneレコード取得
- Excel元データ用シートへ書き込み
- `.xlsx` ダウンロード

## 開発コマンド

PowerShell環境では `npm.cmd` を優先する。

```powershell
npm.cmd run typecheck
npm.cmd test
npm.cmd run package
```

`npm.cmd run package` は以下を行う。

- Viteで `desktop.js` と `config.js` を生成
- `plugin-package/` にkintoneプラグイン構成を配置
- `dist/kintone-excel-report-plugin.zip` を生成

## 変更時の目安

- 相談済みの今後実装する内容は `SPEC.md` の「今後の拡張候補」と「次の実装順」を正とする
- PL Management案件を継続する場合は、先に `SPEC.md` の「PL Management 導入・引継ぎ状況」を確認する。帳票ID、テンプレート、検証済み範囲、残作業をここで引き継ぐ
- PL案件の帳票設定を変更した場合は、完成版テンプレート、取得元アプリ設定JSON、`SPEC.md` の引継ぎ状況を同じ変更単位で更新する。環境固有のアプリID・接続先・認証情報は書かない
- 単月帳票のフィルターで `aggregation_status` を指定する場合は、Kintoneのドロップダウン仕様に合わせて演算子を `in` にする。`=` は `GAIA_IQ03` になる
- 期間ルールを増やす場合: `src/shared/dateRules.ts`
- Excelシート書き込みを変える場合: `src/shared/excel.ts`
- kintoneの取得条件やカーソルAPI処理を変える場合: `src/shared/kintoneApi.ts`
- 設定項目を増やす場合:
  - `src/shared/types.ts`
  - `src/shared/defaults.ts`
  - `src/shared/config.ts`
  - `src/config.ts`
  - `public/config.html`
- 帳票ごとの取得元アプリ設定JSONを変える場合: テンプレート管理アプリの `sources_json`
- 取得元アプリ設定JSONの利用者向けサンプル: `SOURCES_JSON_EXAMPLES.md`
- テンプレート管理レコード追加/編集画面には、複数取得元アプリ対応のJSON簡易生成UIがある
- テンプレート管理アプリでは紐づく帳票出力アプリIDを設定し、抽出条件で出力アプリの任意フィールドを参照できる
- 簡易生成UIは、アプリIDからフィールド定義を取得し、型推定、選択、削除、上下移動ができる
- 抽出条件は `filters`、並び順は `sorts` の配列だけを正式仕様として扱う
- 旧 `periodRule` / `additionalQuery` / `valueFrom: "period"` は互換対応せず、簡易生成UIで作り直す
- manifestや同梱ファイルを変える場合: `plugin/manifest.json` と `scripts/build-plugin.mjs`

## 検証

最低限、以下を実行する。

```powershell
npm.cmd run typecheck
npm.cmd test
npm.cmd run package
```

既知の警告:

- `kintone-plugin-packer` が `homepage_url.ja` / `homepage_url.en` 未設定の警告を出すことがある。zip生成が成功していれば現時点では許容。

## 禁止事項

- 実運用のAPIトークン、パスワード、個人情報、実データを保存しない
- 集計シートの数式をTypeScript側で再現しようとしない
- unrelatedなリファクタリングを混ぜない
- 生成済みExcelやzipをソース管理対象にしない
